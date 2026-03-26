"""
Prompt Loader — loads prompt text from external .txt files.

Usage:
    from src.utils.prompt_loader import load_prompt
    prompt = load_prompt("01_identify_drawing_layout")
"""
import os
from functools import lru_cache

from src.utils.logger import get_logger

logger = get_logger(__name__)

# prompts/ directory at project root
_PROMPTS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "prompts")


@lru_cache(maxsize=32)
def load_prompt(name: str) -> str:
    """
    Load a prompt from prompts/<name>.txt and return its content.

    Args:
        name: Prompt filename without extension, e.g. "01_identify_drawing_layout"

    Returns:
        The prompt text (stripped of leading/trailing whitespace).

    Raises:
        FileNotFoundError: if the prompt file doesn't exist.
    """
    path = os.path.join(_PROMPTS_DIR, f"{name}.txt")
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Prompt file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        text = f.read().strip()
    logger.debug(f"Loaded prompt '{name}' ({len(text)} chars)")
    return text
