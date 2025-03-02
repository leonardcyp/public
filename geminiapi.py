import os
import dotenv
from google import genai

dotenv.load_dotenv()
api_key = os.getenv("LLM_API_KEY") #Input your API Key to interact with Gemini AI
client = genai.Client(api_key=api_key)

def generate_response(user_input):
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=user_input,
    )
    return response.text# if 'text' in response else "No text field found"

def main():
    print("Welcome to the Chatbot! Type 'exit' to quit.")
    while True:
        user_input = input("You: ").strip()
        if user_input.lower() == 'exit':
            print("Goodbye!")
            break
        bot_response = generate_response(user_input)
        print("Bot:", bot_response)

if __name__ == "__main__":
    main()