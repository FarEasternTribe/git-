from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

client = OpenAI()

response = client.responses.create(
    model="gpt-5",
    input="12345×6789を計算して、途中式も説明してください。"
)

print(response.output_text)