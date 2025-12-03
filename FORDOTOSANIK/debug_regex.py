import re

def find_value(text, pattern):
    clean_pattern = pattern.strip()
    if clean_pattern.endswith(":"):
        clean_pattern = clean_pattern[:-1].strip()
        
    escaped_chars = [re.escape(c) for c in clean_pattern]
    flexible_pattern = r"[ \t]*".join(escaped_chars)
    final_regex_base = fr"{flexible_pattern}[ \t]*:?"
    
    print(f"Pattern: {pattern}")
    print(f"Regex Base: {final_regex_base}")
    
    # Multiline regex
    regex = fr"{final_regex_base}[ \t]*[\r\n]+\s*([^\n]+)"
    print(f"Full Regex: {regex}")
    
    match = re.search(regex, text, re.IGNORECASE)
    if match:
        print(f"Match found: '{match.group(1)}'")
        return match.group(1).strip()
    else:
        print("No match found")
        return ""

text = "TC:\n98765432109\n"
val = find_value(text, "TC:")
print(f"Result: '{val}'")
