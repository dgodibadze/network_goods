#!/usr/bin/env python3
"""
BGP Data Swapper v0.2

Description:
    This script demonstrates swapping BGP neighbor mapping data. The original data is provided
    as a dictionary mapping IP addresses to hostnames. In this version (v0.2), the swapped data
    output is simplified to display the hostname only—omitting the associated IP addresses—even if
    multiple IPs map to the same hostname.

Usage:
    python bgp_data_swapper.py

Author: Your Name
Version: 0.2
Date: 2025-02-13
"""

def swap_bgp_data(neighbor_mapping):
    """
    Swap the neighbor mapping data to use hostname as the primary key.

    Args:
        neighbor_mapping (dict): Mapping of BGP neighbors with IP addresses as keys and
                                 hostnames as values.
                                 Example: {'10.1.1.1': 'routerA', '10.1.1.2': 'routerB'}

    Returns:
        dict: A new mapping with hostnames as keys. If multiple IPs map to the same hostname,
              they are aggregated in a list.
    """
    swapped = {}
    for ip, hostname in neighbor_mapping.items():
        # If the hostname already exists, aggregate the IP addresses
        if hostname in swapped:
            # Convert existing value to list if it is not already one
            if not isinstance(swapped[hostname], list):
                swapped[hostname] = [swapped[hostname]]
            swapped[hostname].append(ip)
        else:
            swapped[hostname] = ip
    return swapped

def display_swapped_data(swapped_data):
    """
    Display the swapped data focusing on hostnames only.

    Although the swapped data might include the original IP addresses in the background,
    this function prints only the hostname. This aligns with the v0.2 requirement.

    Args:
        swapped_data (dict): Mapping with hostnames as keys.
    """
    print("Swapped Data (Hostname Only):")
    for hostname in swapped_data.keys():
        print(hostname)

def main():
    """
    Main function that defines the BGP neighbor mapping, performs the data swap,
    and displays the results.
    """
    # Example BGP neighbor mapping: IP address -> Hostname.
    # Note: Multiple IP addresses can map to the same hostname.
    bgp_neighbors = {
        '10.1.1.1': 'routerA',
        '10.1.1.2': 'routerB',
        '10.1.1.3': 'routerC',
        '10.1.1.4': 'routerA',  # routerA has more than one IP.
    }

    # Swap the mapping to use hostnames as the key.
    swapped_mapping = swap_bgp_data(bgp_neighbors)

    # Display the swapped data (hostnames only).
    display_swapped_data(swapped_mapping)

if __name__ == "__main__":
    main()
