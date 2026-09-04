"""确定性结构探查：候选表头识别、列概况与基础质量统计。"""

from __future__ import annotations

from collections import Counter

import pandas as pd

from e_master.core.loader.loader import RawSheet, RawWorkbook
from e_master.core.models.profile import (
    ColumnProfile,
    DataIssue,
    HeaderCandidate,
    IssueLevel,
    SheetProfile,
    WorkbookProfile,
)
from e_master.core.text import column_letter, is_blank_text, looks_numeric, normalize_key

# 表头候选扫描：从首个有内容行往后最多扫描的行数。
# 表头通常出现在文件前部，过长会引入数据行干扰排序。
_HEADER_SCAN_LIMIT = 15


class WorkbookAnalyzer:
    """把原始表格转换为结构画像。"""

    def analyze(self, raw: RawWorkbook) -> WorkbookProfile:
        profile = WorkbookProfile(path=raw.path, file_type=raw.file_type, total_sheets=len(raw.sheets))
        for raw_sheet in raw.sheets:
            sheet_profile = self._analyze_sheet(raw_sheet)
            profile.sheets.append(sheet_profile)
            profile.issues.extend(sheet_profile.issues)
        return profile

    @staticmethod
    def _analyze_sheet(raw_sheet: RawSheet) -> SheetProfile:
        profile = SheetProfile(sheet_name=raw_sheet.sheet_name)
        frame = raw_sheet.data
        if frame.shape[0] == 0 or frame.shape[1] == 0:
            profile.issues.append(
                DataIssue(
                    code="EMPTY_SHEET",
                    message="该 Sheet 没有任何内容",
                    level=IssueLevel.WARNING,
                    sheet=profile.sheet_name,
                )
            )
            return profile

        blank = frame.map(is_blank_text)
        row_has_content = (~blank).any(axis=1)
        col_has_content = (~blank).any(axis=0)
        if not bool(row_has_content.any()):
            profile.issues.append(
                DataIssue(
                    code="EMPTY_SHEET",
                    message="该 Sheet 没有任何内容",
                    level=IssueLevel.WARNING,
                    sheet=profile.sheet_name,
                )
            )
            return profile

        used_row_positions = frame.index[row_has_content].to_list()
        used_col_positions = frame.columns[col_has_content].to_list()
        used_first_row = int(used_row_positions[0]) + 1
        used_last_row = int(used_row_positions[-1]) + 1
        used_first_col = int(used_col_positions[0]) + 1
        used_last_col = int(used_col_positions[-1]) + 1
        profile.used_first_row = used_first_row
        profile.used_last_row = used_last_row
        profile.used_first_col = used_first_col
        profile.used_last_col = used_last_col

        first_col_idx = used_first_col - 1
        last_col_idx = used_last_col - 1
        scan_end_row = min(used_last_row, used_first_row + _HEADER_SCAN_LIMIT - 1)

        candidates = [
            candidate
            for row_number in range(used_first_row, scan_end_row + 1)
            if (candidate := WorkbookAnalyzer._score_header_row(frame, row_number, first_col_idx, last_col_idx)) is not None
        ]
        candidates.sort(key=lambda item: (-item.score, item.row_number))
        profile.header_candidates = candidates[:5]
        profile.recommended_header = candidates[0] if candidates and candidates[0].score >= 0.3 else None

        if profile.recommended_header is None:
            profile.issues.append(
                DataIssue(
                    code="NO_HEADER_CANDIDATE",
                    message="未找到可信的候选表头行，后续字段映射需要人工确认",
                    level=IssueLevel.WARNING,
                    sheet=profile.sheet_name,
                    details={"scanned_rows": scan_end_row - used_first_row + 1},
                )
            )
            return profile

        recommended = profile.recommended_header
        header_pos = recommended.row_number - 1
        profile.data_start_row = recommended.row_number + 1
        data_start_pos = recommended.row_number  # 物理行号转 0 起始下标
        data_end_pos = used_last_row

        if data_start_pos >= data_end_pos:
            profile.issues.append(
                DataIssue(
                    code="NO_DATA_BELOW_HEADER",
                    message=f"推荐表头为第 {recommended.row_number} 行，但未发现其下方的数据行",
                    level=IssueLevel.INFO,
                    sheet=profile.sheet_name,
                )
            )
            return profile

        profile.data_row_count = data_end_pos - data_start_pos
        self_row_count = data_end_pos - data_start_pos
        data_rows = frame.iloc[data_start_pos:data_end_pos, first_col_idx : last_col_idx + 1]
        data_blank = data_rows.map(is_blank_text)
        profile.empty_cell_count = int(data_blank.sum().sum())

        # 列概况
        header_names: list[str] = []
        for col_num in range(used_first_col, used_last_col + 1):
            col_idx = col_num - 1
            excel_column = column_letter(col_idx)
            header_value = frame.iat[header_pos, col_idx]
            name = str(header_value).strip() if not is_blank_text(header_value) else ""
            header_names.append(name)
            series = data_rows.iloc[:, col_num - used_first_col]
            non_empty = int((~data_blank.iloc[:, col_num - used_first_col]).sum())
            samples = tuple(
                value
                for value in series.tolist()
                if not is_blank_text(value)
            )[:3]
            profile.columns.append(
                ColumnProfile(
                    excel_column=excel_column,
                    name=name or f"列{excel_column}",
                    non_empty_count=non_empty,
                    sample_values=samples,
                )
            )
            if not name:
                profile.issues.append(
                    DataIssue(
                        code="EMPTY_HEADER_CELL",
                        message=f"表头行中第 {excel_column} 列缺少字段名，已使用占位名",
                        level=IssueLevel.WARNING,
                        sheet=profile.sheet_name,
                        row=recommended.row_number,
                        column=excel_column,
                    )
                )
            elif looks_numeric(name):
                profile.issues.append(
                    DataIssue(
                        code="NUMERIC_HEADER_CELL",
                        message=f"第 {excel_column} 列表头 '{name}' 是纯数字，可能不是字段名",
                        level=IssueLevel.WARNING,
                        sheet=profile.sheet_name,
                        row=recommended.row_number,
                        column=excel_column,
                    )
                )
        profile.column_count = len(profile.columns)

        # 表头前疑似标题/说明内容
        content_before_header = 0
        if header_pos > used_first_row - 1:
            rows_above = frame.iloc[used_first_row - 1 : header_pos, first_col_idx : last_col_idx + 1]
            content_before_header = int((~rows_above.map(is_blank_text)).any(axis=1).sum())
        if content_before_header:
            profile.issues.append(
                DataIssue(
                    code="CONTENT_ABOVE_HEADER",
                    message=f"推荐表头前有 {content_before_header} 行内容，疑似标题或说明文字",
                    level=IssueLevel.INFO,
                    sheet=profile.sheet_name,
                    details={"row_count": content_before_header},
                )
            )

        # 重复表头名
        normalized_headers = [normalize_key(name) for name in header_names if name]
        duplicated_headers = sorted({name for name, count in Counter(normalized_headers).items() if count > 1})
        if duplicated_headers:
            profile.issues.append(
                DataIssue(
                    code="DUPLICATE_HEADER_NAME",
                    message="存在重复字段名: " + "、".join(duplicated_headers),
                    level=IssueLevel.WARNING,
                    sheet=profile.sheet_name,
                    row=recommended.row_number,
                    details={"names": duplicated_headers},
                )
            )

        # 重复数据行
        if self_row_count > 0:
            keys = data_rows.apply(
                lambda row: "|".join(normalize_key(value) for value in row),
                axis=1,
            )
            has_content = (~data_blank).any(axis=1)
            keys_with_content = keys[has_content]
            profile.duplicate_data_rows = int(keys_with_content.duplicated(keep="first").sum())
            if profile.duplicate_data_rows:
                profile.issues.append(
                    DataIssue(
                        code="DUPLICATE_DATA_ROWS",
                        message=f"数据区发现 {profile.duplicate_data_rows} 行完全重复的数据",
                        level=IssueLevel.WARNING,
                        sheet=profile.sheet_name,
                        details={"duplicate_count": profile.duplicate_data_rows},
                    )
                )

        return profile

    @staticmethod
    def _score_header_row(
        frame: pd.DataFrame,
        row_number: int,
        first_col_idx: int,
        last_col_idx: int,
    ) -> HeaderCandidate | None:
        """给一行打分，判断其“像表头”的程度。"""
        row = frame.iloc[row_number - 1, first_col_idx : last_col_idx + 1]
        values = [str(value) for value in row.tolist() if not is_blank_text(value)]
        total = last_col_idx - first_col_idx + 1
        if not values:
            return None

        populated_ratio = len(values) / total
        text_count = sum(1 for value in values if not looks_numeric(value))
        text_ratio = text_count / len(values)
        unique_ratio = len({normalize_key(value) for value in values}) / len(values)
        score = 0.45 * populated_ratio + 0.35 * text_ratio + 0.20 * unique_ratio
        reasons = (
            f"填充率 {populated_ratio:.0%}",
            f"文本占比 {text_ratio:.0%}",
            f"唯一值占比 {unique_ratio:.0%}",
        )
        return HeaderCandidate(
            row_number=row_number,
            score=round(score, 3),
            reasons=reasons,
            sample_values=tuple(values[:8]),
        )
