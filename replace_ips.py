#!/usr/bin/env python3
"""
replace_ips.py (v0.2)

This script reads a mapping file (hostname_mapping.txt) containing hostnames, IP addresses, and descriptions.
It then accepts multi-line text input from the user (terminated by a line containing only "=="),
replaces any IPv4 addresses found in the text with corresponding hostnames (if a match exists),
and prints the modified text along with a summary of all substitutions made.

Usage:
    python3 replace_ips.py
"""

import sys
import re

def load_mapping(filename):
    """
    Load the hostname mapping from a file.
    Each non-comment line should contain exactly three comma-separated values:
    hostname, IP address, and description.

    Args:
        filename (str): The mapping file name.

    Returns:
        dict: A dictionary mapping IP addresses (str) to a tuple (hostname, description).
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
                    # Skip malformed lines
                    continue
                hostname = parts[0].strip()
                ip_address = parts[1].strip()
                description = parts[2].strip()
                mapping[ip_address] = (hostname, description)
    except FileNotFoundError:
        print(f"Error: Mapping file '{filename}' not found.")
        sys.exit(1)
    return mapping

def get_user_input():
    """
    Reads multi-line text input from the user.
    The user must finish input by entering a line that contains only "==".

    Returns:
        str: The complete text input joined by newline characters.
    """
    print("Paste your text below. End input with a line containing only '==':")
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

def replace_ips_in_text(text, ip_mapping):
    """
    Replace any IPv4 addresses in the text with the corresponding hostname if found in the mapping.

    Args:
        text (str): The original text input.
        ip_mapping (dict): Dictionary mapping IP addresses to (hostname, description).

    Returns:
        tuple:
            - modified_text (str): The text after substitutions.
            - substitutions (dict): A dictionary of IP addresses replaced with (hostname, description).
    """
    substitutions = {}  # Records substitutions made

    # Regular expression to match IPv4 addresses
    ip_regex = re.compile(r'\b((?:\d{1,3}\.){3}\d{1,3})\b')

    def substitution(match):
        ip = match.group(1)
        if ip in ip_mapping:
            hostname, description = ip_mapping[ip]
            # Record substitution details (only record once per unique IP)
            if ip not in substitutions:
                substitutions[ip] = (hostname, description)
            return hostname
        return ip  # Leave unchanged if no match found

    modified_text = ip_regex.sub(substitution, text)
    return modified_text, substitutions

def print_summary(substitutions):
    """
    Print a summary of all substitutions made in the following format:
    original_IP -> hostname - description

    Args:
        substitutions (dict): Dictionary where key is original IP and value is (hostname, description).
    """
    print("\nSubstitution Summary:")
    if not substitutions:
        print("No IP addresses were replaced.")
    else:
        for ip, (hostname, description) in substitutions.items():
            print(f"{ip} -> {hostname} - {description}")

def main():
    mapping_file = "hostname_mapping.txt"
    ip_mapping = load_mapping(mapping_file)

    # Get multi-line text input from the user
    user_text = get_user_input()

    # Process the text to replace IP addresses with hostnames
    modified_text, substitutions = replace_ips_in_text(user_text, ip_mapping)

    # Display the modified text
    print("\nModified Text:")
    print(modified_text)

    # Print the substitution summary
    print_summary(substitutions)

if __name__ == '__main__':
    main()
