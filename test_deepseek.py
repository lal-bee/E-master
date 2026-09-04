from openai import OpenAI
from dotenv import load_dotenv
import os


load_dotenv()


client = OpenAI(
    api_key=os.getenv("DEEPSEEK_API_KEY"),
    base_url=os.getenv("DEEPSEEK_BASE_URL")
)


response = client.chat.completions.create(
    model=os.getenv("MODEL"),
    messages=[
        {
            "role": "user",
            "content": "你好，请回复：DeepSeek连接成功"
        }
    ]
)


print(response.choices[0].message.content)