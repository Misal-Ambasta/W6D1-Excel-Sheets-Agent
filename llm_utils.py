import os
from typing import Dict, Any, Optional
from langchain_google_genai import GoogleGenerativeAI
from langchain.prompts import PromptTemplate
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def get_gemini_llm():
    """Return a Gemini LLM instance using langchain-google-genai."""
    model_name = os.environ.get("GEMINI_MODEL_NAME", "gemini-pro")
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        raise ValueError("GOOGLE_API_KEY is not set in environment variables.")
    return GoogleGenerativeAI(model=model_name, google_api_key=api_key)

# Prompt templates for different query types
PROMPT_TEMPLATES = {
    "filter": PromptTemplate(
        input_variables=["query", "columns"],
        template=(
            "You are an expert data analyst. Given the following columns: {columns}, "
            "translate the user's request into a pandas filter expression. "
            "User request: '{query}'\n"
            "Return only the pandas filter code (e.g., df[df['Region'] == 'Delhi'])"
        ),
    ),
    "aggregate": PromptTemplate(
        input_variables=["query", "columns"],
        template=(
            "Given columns: {columns}, translate the user's aggregation query into a pandas groupby/agg expression. "
            "User request: '{query}'\n"
            "Return only the pandas code."
        ),
    ),
    "sort": PromptTemplate(
        input_variables=["query", "columns"],
        template=(
            "Given columns: {columns}, translate the user's sorting/grouping query into a pandas sort_values or groupby expression. "
            "User request: '{query}'\n"
            "Return only the pandas code."
        ),
    ),
    "pivot": PromptTemplate(
        input_variables=["query", "columns"],
        template=(
            "Given columns: {columns}, translate the user's pivot table query into a pandas pivot_table expression. "
            "User request: '{query}'\n"
            "Return only the pandas code."
        ),
    ),
}

def get_prompt(query_type: str, query: str, columns: list) -> str:
    """Render the appropriate prompt for the query type."""
    if query_type not in PROMPT_TEMPLATES:
        raise ValueError(f"Unknown query type: {query_type}")
    return PROMPT_TEMPLATES[query_type].format(query=query, columns=columns)

def parse_llm_response(response: str) -> Optional[str]:
    """Extract pandas code from LLM response."""
    # Simple extraction; can be improved
    code = response.strip()
    if 'df' in code:
        return code
    return None

def run_gemini_query(query_type: str, query: str, columns: list) -> Optional[str]:
    """Send prompt to Gemini LLM and return pandas code string."""
    llm = get_gemini_llm()
    prompt = get_prompt(query_type, query, columns)
    try:
        response = llm(prompt)
        code = parse_llm_response(response)
        return code
    except Exception as e:
        print(f"Error calling Gemini LLM: {e}")
        return None
