import os
from pathlib import Path
from dotenv import load_dotenv
from agno.agent import Agent
from agno.models.openai import OpenAILike

# Build the path to the .env file relative to the script's location
script_dir = Path(__file__).parent
dotenv_path = script_dir / 'config' / '.env'

# Load environment variables from the constructed path
load_dotenv(dotenv_path=dotenv_path)

# --- Zhipu AI PaaS/Subscription Model Setup with Agno ---
zhipu_api_key = os.getenv("ZHIPU_API_KEY")
zhipu_base_url = "https://open.bigmodel.cn/api/coding/paas/v4"

# Create an agent using Agno's OpenAILike for Zhipu
zhipu_agent = None
if zhipu_api_key and zhipu_api_key != "YOUR_ZHIPU_API_KEY":
    try:
        zhipu_model = OpenAILike(
            id="glm-4.6",  # Model ID for Zhipu, e.g., "glm-4"
            api_key=zhipu_api_key,
            base_url=zhipu_base_url
        )
        zhipu_agent = Agent(model=zhipu_model)
    except Exception as e:
        print(f"Error initializing Zhipu agent: {e}")

# Define a simple task
def simple_task(agent, prompt):
    if agent:
        try:
            response = agent.run(prompt).content
            return response
        except Exception as e:
            return f"An error occurred during chat: {e}"
    return "Agent not initialized."

# Run the task
if __name__ == "__main__":
    print("--- Running Zhipu AI Task with Agno ---")
    if not zhipu_agent:
        print("Zhipu agent could not be initialized. Please check your ZHIPU_API_KEY in 'apps/exp/config/.env'.")
    else:
        prompt = "你好，介绍一下你自己。"
        response = simple_task(zhipu_agent, prompt)
        print(f"Zhipu AI Response: {response}")
