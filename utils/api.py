import requests
import json

def call_api(api_config, system_prompt, user_prompt, api_timeout=180):
    """
    Calls the appropriate AI API based on api_config['TYPE'].
    Handles request setup, execution, and error handling for different API types.
    Returns the generated text content or an error message string.
    """
    api_name = api_config["NAME"]
    api_type = api_config["TYPE"]
    print(
        f"  Calling API '{api_name}' (Type: {api_type}) for prompt: '{user_prompt[:30]}...'"
    )

    try:
        if api_type == "openai":
            # --- OpenAI Compatible API Call ---
            headers = {
                "Authorization": f"Bearer {api_config['KEY']}",
                "Content-Type": "application/json",
            }
            data = {
                "model": api_config["MODEL"],
                "messages": [],
            }
            # Add system prompt if provided and not empty
            if system_prompt:
                data["messages"].append({"role": "system", "content": system_prompt})
            data["messages"].append({"role": "user", "content": user_prompt})

            response = requests.post(
                f"{api_config['ENDPOINT']}/chat/completions",  # Assume /chat/completions endpoint
                headers=headers,
                json=data,
                timeout=api_timeout,
            )
            response.raise_for_status()
            result = response.json()

            # Extract content - handle potential variations in response structure
            if result.get("choices") and len(result["choices"]) > 0:
                message = result["choices"][0].get("message", {})
                content = message.get("content")
                if content:
                    print(f"  API call '{api_name}' successful.")
                    return content.strip()
                else:
                    # Handle cases like function calls if needed in the future
                    print(
                        f"  API Error ({api_name}): No 'content' found in message: {message}"
                    )
                    return f"Error ({api_name}): No content in response message."
            else:
                print(f"  API Error ({api_name}): Unexpected response format: {result}")
                return f"Error ({api_name}): Unexpected API response format."

        elif api_type == "google":
            # --- Google Gemini API Call ---
            headers = {"Content-Type": "application/json"}
            # Construct the specific URL for Google Gemini
            url = f"{api_config['ENDPOINT']}/{api_config['MODEL']}:generateContent?key={api_config['KEY']}"
            # Construct the specific JSON body for Google Gemini
            # Note: System prompt handling might differ for Gemini.
            # This basic implementation only sends the user prompt.
            # More complex scenarios might require adjusting the 'contents' structure.
            data = {"contents": [{"parts": [{"text": user_prompt}]}]}
            if system_prompt:
                # Basic system prompt integration (may need refinement based on Gemini best practices)
                data["systemInstruction"] = {"parts": [{"text": system_prompt}]}

            response = requests.post(
                url, headers=headers, json=data, timeout=api_timeout
            )
            response.raise_for_status()
            result = response.json()

            # Extract content from Google Gemini response
            if result.get("candidates") and len(result["candidates"]) > 0:
                candidate = result["candidates"][0]
                if (
                    candidate.get("content")
                    and candidate["content"].get("parts")
                    and len(candidate["content"]["parts"]) > 0
                ):
                    content = candidate["content"]["parts"][0].get("text")
                    if content:
                        print(f"  API call '{api_name}' successful.")
                        return content.strip()
                    else:
                        print(
                            f"  API Error ({api_name}): No 'text' found in content part: {candidate['content']['parts'][0]}"
                        )
                        return f"Error ({api_name}): No text in response part."
                else:
                    # Handle safety ratings, finish reasons etc. if needed
                    finish_reason = candidate.get("finishReason", "UNKNOWN")
                    safety_ratings = candidate.get("safetyRatings", [])
                    print(
                        f"  API Warning/Error ({api_name}): No content/parts found. Finish Reason: {finish_reason}. Safety: {safety_ratings}"
                    )
                    return f"Error ({api_name}): No content/parts in response. Finish: {finish_reason}"
            else:
                print(f"  API Error ({api_name}): Unexpected response format: {result}")
                return f"Error ({api_name}): No candidates in API response."

        else:
            print(
                f"  API Error ({api_name}): Unsupported API TYPE '{api_type}' in config."
            )
            return f"Error: Unsupported API type '{api_type}'"

    except requests.exceptions.Timeout:
        print(f"  API Request Error ({api_name}): Timeout after {api_timeout} seconds.")
        return f"Error ({api_name}): API request timed out."
    except requests.exceptions.RequestException as e:
        error_message = f"Error ({api_name}): API request failed."
        if e.response is not None:
            try:
                # Try to get more specific error from response body
                error_detail = e.response.json()
                error_message += (
                    f" Status: {e.response.status_code}. Detail: {error_detail}"
                )
            except json.JSONDecodeError:
                error_message += (
                    f" Status: {e.response.status_code}. Response: {e.response.text}"
                )
        else:
            error_message += f" Exception: {e}"
        print(f"  API Request Error ({api_name}): {e}")
        return error_message
    except Exception as e:
        print(f"  Error during API call processing ({api_name}): {e}")
        # Consider logging the full traceback here for debugging
        # import traceback
        # traceback.print_exc()
        return f"Error ({api_name}): Processing API response failed. {e}"