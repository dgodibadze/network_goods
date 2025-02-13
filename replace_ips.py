import re
import sys

def load_mapping(filename):
    """
    Reads the hostname-to-IP mapping file and returns a dictionary.
    Expected file format (ignoring commented/empty lines):
        hostname,IP address,description
    """
    mapping = {}
    try:
        with open(filename, 'r') as f:
            for line in f:
                line = line.strip()
                # Skip empty lines or lines that start with a comment symbol
                if not line or line.startswith('#'):
                    continue
                # Split by comma and strip extra spaces
                parts = [part.strip() for part in line.split(',')]
                if len(parts) >= 3:
                    hostname, ip, description = parts[0], parts[1], parts[2]
                    mapping[ip] = (hostname, description)
    except FileNotFoundError:
        print(f"Mapping file '{filename}' not found.")
        sys.exit(1)
    return mapping

def get_user_input():
    """
    Read multi-line input from the user until a line with '==' is encountered.
    """
    print("Paste your text below. When finished, type '==' on a new line and press Enter:")
    lines = []
    while True:
        line = sys.stdin.readline()
        # Remove trailing whitespace and check if it's the termination signal
        if line.rstrip() == "==":
            break
        lines.append(line)
    return "".join(lines)

def main():
    # Load the mapping from the file
    mapping = load_mapping("hostname_to_ip_mapping.txt")

    # Get the text input from the user
    text = get_user_input()

    # Regular expression to match IPv4 addresses (simple version)
    ip_regex = re.compile(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b')

    # Dictionary to keep track of replacements {original_ip: (new_text, description)}
    swapped_ips = {}

    def replace_ip(match):
        ip = match.group(0)
        if ip in mapping:
            hostname, description = mapping[ip]
            new_text = f"{ip}({hostname})"
            swapped_ips[ip] = (new_text, description)
            return new_text
        return ip

    # Replace all matching IPs in the text using our replacement function.
    modified_text = ip_regex.sub(replace_ip, text)

    # Output the modified text
    print("\n--- Modified Text ---")
    print(modified_text)

    # Output the list of swapped IPs along with the description.
    if swapped_ips:
        print("\n--- Swapped IPs ---")
        for ip, (new_text, description) in swapped_ips.items():
            print(f"{ip} -> {new_text} - {description}")
    else:
        print("\nNo IP addresses were replaced.")

if __name__ == "__main__":
    main()
