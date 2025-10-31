# Juniper Mist - EVPN IP CLOS Fabric Builder

**Version:** 0.1  
**Author:** Lukas Eisenberger (leisenberger@juniper.net)

> ⚠️ **USE AT YOUR OWN RISK!** ⚠️  
> This tool is in early development (v0.1). Please thoroughly test in a non-production environment before deploying.

## Overview

EVPN IP CLOS Fabric Builder is a tool designed to automate the configuration and deployment of EVPN-based IP CLOS fabric architectures for Juniper Mist. It simplifies the complex process of building scalable EVPN-VXLAN Campus Fabric networks with support for multi-pod architectures and various advanced networking features.

## Features

### Core Capabilities
- **Excel-based Configuration Input**
  - Reads configuration from Excel worksheets (FABRIC, SWITCHES, NETWORKS, WAN)
  - Provides a familiar interface for network engineers to define fabric parameters
  - Examples are given in the provided Excel worksheets
  - Pod-aware and support for serviceblock

### Architecture Support (SWITCH sheet)
- **Pod-aware Design with Multi-pod Support**
  - Scalable architecture supporting multiple pods
  - Simplified inter-pod connectivity management

### Hardware Configuration (SWITCH sheet)
- **Optic Port Configuration**
  - Configurable port speeds
  - Channelization support for breakout cables

### Network Protocol Support (FABRIC sheet)
  - IPv4 underlay support
  - IPv6 underlay support
  - Simultaneous IPv4 and IPv6 support in the overlay network

### Services (NETWORKS sheet)
- **DHCP Relay Support**
  - IPv4 DHCP relay configuration
  - IPv6 DHCP relay configuration

### WAN Integration (WAN sheet)
- **L3 Routing on Border/Core Switches**
  - WAN connectivity for border and core switches using eBGP
  - BGP peering configuration when WAN worksheet is populated
  - Seamless route-filtering according to networks which are defined in NETWORKS and enabled DHCP-Relays (only required prefixes are annouced to the WAN router)
  - Explicit deny support via `EXPORT_EXPLICIT_DENY` and `IMPORT_EXPLICIT_DENY` flags
  - Auto-generated export policy terms with prefix-based matching
  - Result: Granular control over route advertisement and acceptance

### Topology Management
- **Intelligent Topology Handling**
  - Automatically detects existing topologies by name
  - Updates existing topologies instead of creating duplicates
  - Preserves existing configurations where appropriate

## Prerequisites
- Python 3.14 tested
- Requirements: requests, openpyxl

## Installation

```bash
# Clone the repository
git clone https://github.com/eisenluk/mist-campus-fabric-ip-clos.git

# Navigate to the directory
cd mist-campus-fabric-ip-clos

# Install dependencies (if applicable)
# pip install -r requirements.txt
```

## Usage

1. Change values in excel spread sheet
2. Create an API token on your Mist Org
3. Copy API token and ORG-ID to the script mist-campus-fabric-ip-clos.py
4. Run the mist-campus-fabric-ip-clos.py

## Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

## License

- **MIT**

## Support

For questions, issues, or suggestions, please contact:
- **Author:** Lukas Eisenberger
- **Email:** leisenberger@juniper.net

## Disclaimer

This tool is provided as-is without any warranties. Always validate generated configurations before deploying to production environments. The author and contributors are not responsible for any network outages or issues resulting from the use of this tool.

---

*Last updated: [Date]*  
*Version: 0.1*
