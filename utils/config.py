import os
import configparser

def read_config(config_file=".config"):
    """
    Reads all API configurations (KEY, ENDPOINT, MODEL, NAME, TYPE, ENABLED)
    from sections starting with [API_] in the .config file.
    Returns a list of dictionaries, each representing an LLM config,
    or an empty list if no valid sections are found or an error occurs.
    """
    config = configparser.ConfigParser()
    llm_configs = []
    if not os.path.exists(config_file):
        print(f"Error: Configuration file '{config_file}' not found.")
        return llm_configs
    try:
        config.read(config_file, encoding="utf-8")  # Specify encoding
        for section in config.sections():
            if section.startswith("API_"):
                api_config = {
                    "KEY": config.get(section, "KEY", fallback=None),
                    "ENDPOINT": config.get(section, "ENDPOINT", fallback=None),
                    "MODEL": config.get(section, "MODEL", fallback=None),
                    "NAME": config.get(
                        section, "NAME", fallback=section[4:]
                    ),  # Default name from section
                    "TYPE": config.get(
                        section, "TYPE", fallback="openai"
                    ).lower(),  # Default type is openai, ensure lowercase
                    "ENABLED": config.getboolean(section, "ENABLED", fallback=True),
                }
                # Basic validation
                if (
                    api_config["KEY"]
                    and api_config["ENDPOINT"]
                    and api_config["MODEL"]
                    and api_config["ENABLED"]
                ):
                    llm_configs.append(api_config)
                    print(
                        f"Loaded config for: {api_config['NAME']} (Type: {api_config['TYPE']}, Enabled: {api_config['ENABLED']})"
                    )
                else:
                    print(
                        f"Warning: Incomplete or disabled configuration in section '{section}'. Skipping."
                    )
        if not llm_configs:
            print(
                f"Error: No valid [API_*] sections found or configured in '{config_file}'."
            )
        return llm_configs
    except configparser.Error as e:
        print(f"Error reading config file '{config_file}': {e}")
        return []
    except Exception as e:
        print(f"An unexpected error occurred reading config: {e}")
        return []