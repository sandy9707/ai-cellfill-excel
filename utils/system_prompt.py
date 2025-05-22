import os

def initialize_system_prompt(system_prompt_file="systemprompt.txt"):
    """
    Checks if the system prompt file exists. If not, creates an empty file.
    This ensures the script doesn't crash if the file is missing.
    """
    if not os.path.exists(system_prompt_file):
        print(f"Creating '{system_prompt_file}'...")
        with open(system_prompt_file, "w", encoding="utf-8") as f:
            pass  # Create empty file
        print(f"'{system_prompt_file}' created.")

def read_system_prompt(system_prompt_file="systemprompt.txt"):
    """
    Reads the content of the system prompt file.
    Ensures the file exists by calling initialize_system_prompt first.
    Returns the content as a string, or an empty string if reading fails.
    """
    initialize_system_prompt(system_prompt_file)  # Ensure file exists
    try:
        with open(system_prompt_file, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception as e:
        print(f"Error reading '{system_prompt_file}': {e}")
        return ""