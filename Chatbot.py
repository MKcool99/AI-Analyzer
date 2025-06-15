from openai import OpenAI

# You can also set the API key using an environment variable
# os.environ["OPENAI_API_KEY"] = "your-api-key-here"

client = OpenAI(api_key=""

def ask_openai(prompt, model="gpt-3.5-turbo", temperature=0.7, max_tokens=300):
    try:
        chat_completion = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ],
            temperature=temperature,
            max_tokens=max_tokens,
        )
        return chat_completion.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"

if __name__ == "__main__":
    print("Welcome to your helpful assistant! Type 'exit' to quit.")
    while True:
        user_input = input("\nQuestion: ")
        if user_input.lower() in ['exit', 'quit']:
            print("Goodbye!")
            break
        response = ask_openai(user_input)
        print(f"Assistant: {response}")
