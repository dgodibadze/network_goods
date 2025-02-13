#!/usr/bin/env python3
"""
replace_ips.py (v0.1)

This script reads a mapping file (hostname_mapping.txt) containing hostnames, IP addresses, and descriptions.
It then accepts multi-line text input from the user (terminated by a line containing only "=="),
replaces any IP addresses in the text with the corresponding hostname if a match is found,
and prints the modified text.
"""

import sys
import re

def load_mapping(filename):
    """
    Load the mapping of IP addresses to hostnames from the file.
    Each valid line should contain three comma-separated values:
    hostname, IP address, and description.
    Lines starting with '#' or empty lines are ignored.

    Args:
        filename (str): The mapping file name.

    Returns:
        dict: A dictionary mapping IP addresses (str) to hostnames (str).
    """
    mapping = {}
    try:
        with open(filename, 'r') as file:
            for line in file:
                line = line.strip()
                # Skip comments and empty lines
                if not line or line.startswith('#'):
                    continue
                parts = line.split(',')
                if len(parts) != 3:
                    continue  # Skip malformed lines
                hostname = parts[0].strip()
                ip_address = parts[1].strip()
                # We only need the hostname for v0.1; description is ignored.
                mapping[ip_address] = hostname
    except FileNotFoundError:
        print(f"Error: Mapping file '{filename}' not found.")
        sys.exit(1)
    return mapping

def get_user_input():
    """
    Read multi-line input from the user until a line containing only "==" is entered.

    Returns:
        str: The complete text input.
    """
    print("Enter your text. End with a line containing only '==':")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.strip() == "==":
            break
        lines.append(line)
    return "\n".join(lines)

def replace_ips(text, mapping):
    """
    Replace IPv4 addresses in the text with corresponding hostnames if found in the mapping.

    Args:
        text (str): The original text.
        mapping (dict): A dictionary mapping IP addresses to hostnames.

    Returns:
        str: The modified text.
    """
    # Regular expression to match IPv4 addresses
    ip_regex = re.compile(r'\b((?:\d{1,3}\.){3}\d{1,3})\b')
    # Replace found IP with hostname if available, otherwise keep the original IP.
    return ip_regex.sub(lambda match: mapping.get(match.group(1), match.group(1)), text)

def main():
    mapping_file = "hostname_mapping.txt"
    ip_mapping = load_mapping(mapping_file)
    user_text = get_user_input()
    modified_text = replace_ips(user_text, ip_mapping)

    print("\nModified Text:")
    print(modified_text)

if __name__ == '__main__':
    main()
