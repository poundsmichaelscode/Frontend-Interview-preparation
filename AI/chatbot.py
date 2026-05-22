from openai import OpenAI

client = OpenAI

response = client.response.create(


input= "write a poem of four lines",
model = "gpt-4"
)

print (response.output_text)
