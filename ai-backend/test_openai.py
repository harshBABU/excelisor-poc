from openai import OpenAI

client = OpenAI()

response = client.chat.completions.create(
    model="gpt-5-nano",
    messages=[
        {"role": "user", "content": "WRITE BACK HI!"}
    ]
)

print(response.choices[0].message.content)
