"""Central LLM configuration constants."""
import os
import contextvars

DEFAULT_OPENAI_MODEL = "gpt-5-nano"

# Store the API key for the current request
api_key_ctx = contextvars.ContextVar("api_key_ctx", default=None)

# Store the selected model for the current request
model_ctx = contextvars.ContextVar("model_ctx", default=None)

def get_model() -> str:
    # Use selected model or fallback to default
    return model_ctx.get() or DEFAULT_OPENAI_MODEL

def get_openai_client():
    from openai import OpenAI
    key = api_key_ctx.get()
    
    if key:
        return OpenAI(api_key=key)
        
    # Fallback to local environment if no request context key is passed
    env_key = os.getenv("OPENAI_API_KEY")
    if env_key:
        return OpenAI(api_key=env_key)
        
    # If nothing is found, it will try to initialize but might fail if OpenAI doesn't find the env variable either
    return OpenAI()
