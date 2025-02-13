#!/usr/bin/env python3
"""
replace_ips.py v0.3

This script reads a mapping file (hostname_to_ip_mapping.txt) that contains lines in the format:
    hostname, IP address, description
Comments (lines starting with '#') and empty lines are ignored.

Usage:
    Run the script:
        python replace_ips.py
    Then, paste your text. When finished, type "==" on a new line to signal the end of input.

The script will scan the pasted text for IP addresses. If an IP is found in the mapping file,
it will be replaced with the corresponding hostname. After processing, the script prints:
    - The modified text.
    - A summary of all substitutions in the format:
        original_IP -> hostname - description
"""

import re
import sys

def load_mapping(mapping_file: str) -> dict:
    """
    Load the hostname to IP mapping from a file.

    Each valid line in the file should have three comma-separated values:
      hostname, IP address, description

    Lines starting with '#' or empty lines are skipped.

    Args:
        mapping_file (str): The path to the mapping file.

    Returns:
        dict: A dictionary with IP addresses as keys and a tuple (hostname, description) as values.
    """
    mapping = {}
    try:
        with open(mapping_file, 'r') as file:
            for line in file:
                line = line.strip()
                # Skip empty lines or comment lines
                if not line or line.startswith('#'):
                    continue
                # Split the line into parts and strip any extra whitespace
                parts = [part.strip() for part in line.split(',')]
                if len(parts) >= 3:
                    hostname, ip, description = parts[0], parts[1], parts[2]
                    mapping[ip] = (hostname, description)
    except FileNotFoundError:
        print(f"Error: Mapping file '{mapping_file}' not found.")
        sys.exit(1)
    return mapping

def get_user_input() -> str:
    """
    Prompt the user to paste text input. Input ends when a line with '==' is entered.

    Returns:
        str: The complete text input provided by the user.
    """
    print("Paste your text below. When finished, type '==' on a new line and press Enter:")
    input_lines = []
    while True:
        # Read a line from standard input
        line = sys.stdin.readline()
        # Check if the termination signal is entered (line with '==')
        if line.rstrip() == "==":
            break
        input_lines.append(line)
    return "".join(input_lines)

def main():
    # Define the mapping file name
    MAPPING_FILE = "hostname_to_ip_mapping.txt"

    # Load the IP-to-hostname mapping from the file
    ip_to_hostname = load_mapping(MAPPING_FILE)

    # Get the text input from the user
    original_text = get_user_input()

    # Regular expression pattern to match IPv4 addresses.
    # This pattern matches four groups of 1-3 digits separated by dots.
    ip_pattern = re.compile(r'\b(?:\d{1,3}\.){3}\d{1,3}\b')

    # Dictionary to store details of IPs that were replaced.
    # Key: original IP; Value: tuple (hostname, description)
    replaced_ips = {}

    def replace_ip(match: re.Match) -> str:
        """
        Replacement function used with re.sub().
        If the matched IP is in the mapping, return the hostname;
        otherwise, return the original IP.
        """
        ip_address = match.group(0)
        if ip_address in ip_to_hostname:
            hostname, description = ip_to_hostname[ip_address]
            replaced_ips[ip_address] = (hostname, description)
            return hostname  # Replace with hostname only
        return ip_address

    # Replace all matching IP addresses in the original text.
    modified_text = ip_pattern.sub(replace_ip, original_text)

    # Display the modified text.
    print("\n--- Modified Text ---")
    print(modified_text)

    # Display a summary of all substitutions made.
    if replaced_ips:
        print("\n--- Replaced IPs Summary ---")
        for ip, (hostname, description) in replaced_ips.items():
            print(f"{ip} -> {hostname} - {description}")
    else:
        print("\nNo IP addresses were replaced.")

if __name__ == "__main__":
    main()
