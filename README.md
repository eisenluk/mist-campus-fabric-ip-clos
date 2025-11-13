# Juniper Mist - EVPN IP CLOS Fabric Builder

**Version:** 0.2
**Author:** Lukas Eisenberger (leisenberger@juniper.net)

> ⚠️ **USE AT YOUR OWN RISK!** ⚠️
> This tool is in early development. Please thoroughly test in a non-production environment before deploying.

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
- **Flexible Pod Topologies** ✨ NEW in v0.2
  - **3-tier topology**: Access → Distribution → Core (traditional CLOS architecture)
  - **2-tier topology**: Access → Core (spine-leaf style, no distribution layer)
  - **Mixed deployments**: Different pods can use different topologies in the same fabric
  - Distribution switches are now optional per pod
  - Automatic topology detection and validation

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

### Configuring Flexible Pod Topologies

Starting with v0.2, you can configure pods with or without distribution switches:

#### 2-Tier Topology (Access → Core)
For a simple spine-leaf deployment without a distribution layer:
- Define access switches with pod numbers
- Set access switch UPLINKS to point directly to core switches
- Do NOT create distribution switches for that pod

Example SWITCHES sheet configuration:
```
HOSTNAME    ROLE    POD    UPLINKS       UPLINK_PORTS        DOWNLINKS    DOWNLINK_PORTS
Core1       core           (leave empty) (leave empty)       A1,A2        ge-0/0/0,ge-0/0/1
Core2       core           (leave empty) (leave empty)       A1,A2        ge-0/0/0,ge-0/0/1
A1          access  1      Core1,Core2   ge-0/0/0,ge-0/0/1   (empty)      (empty)
A2          access  1      Core1,Core2   ge-0/0/0,ge-0/0/1   (empty)      (empty)
```

#### 3-Tier Topology (Access → Distribution → Core)
For traditional CLOS architecture with distribution layer:
- Define access switches with pod numbers
- Create distribution switches for that pod
- Set access switch UPLINKS to distribution switches
- Set distribution switch UPLINKS to core switches

Example SWITCHES sheet configuration:
```
HOSTNAME    ROLE          POD    UPLINKS       UPLINK_PORTS        DOWNLINKS    DOWNLINK_PORTS
Core1       core                 (leave empty) (leave empty)       D1,D2        ge-0/0/0,ge-0/0/1
Core2       core                 (leave empty) (leave empty)       D1,D2        ge-0/0/0,ge-0/0/1
D1          distribution  2      Core1,Core2   ge-0/0/0,ge-0/0/1   A3,A4        ge-0/0/2,ge-0/0/3
D2          distribution  2      Core1,Core2   ge-0/0/0,ge-0/0/1   A3,A4        ge-0/0/2,ge-0/0/3
A3          access        2      D1,D2         ge-0/0/0,ge-0/0/1   (empty)      (empty)
A4          access        2      D1,D2         ge-0/0/0,ge-0/0/1   (empty)      (empty)
```

#### Mixed Topology (Multiple Pods with Different Architectures)
You can combine both approaches in a single fabric:
```
HOSTNAME    ROLE          POD    UPLINKS       UPLINK_PORTS        DOWNLINKS         DOWNLINK_PORTS
Core1       core                 (leave empty) (leave empty)       D1,D2,A1,A2       ge-0/0/0-3
Core2       core                 (leave empty) (leave empty)       D1,D2,A1,A2       ge-0/0/0-3
D1          distribution  1      Core1,Core2   ge-0/0/0,ge-0/0/1   A5,A6             ge-0/0/2,ge-0/0/3
D2          distribution  1      Core1,Core2   ge-0/0/0,ge-0/0/1   A5,A6             ge-0/0/2,ge-0/0/3
A5          access        1      D1,D2         ge-0/0/0,ge-0/0/1   (empty)           (empty)
A6          access        1      D1,D2         ge-0/0/0,ge-0/0/1   (empty)           (empty)
A1          access        2      Core1,Core2   ge-0/0/0,ge-0/0/1   (empty)           (empty)
A2          access        2      Core1,Core2   ge-0/0/0,ge-0/0/1   (empty)           (empty)
```
In this example:
- **Pod 1** uses 3-tier: A5,A6 → D1,D2 → Core1,Core2
- **Pod 2** uses 2-tier: A1,A2 → Core1,Core2 (no distribution switches)

The script will automatically detect and validate the topology for each pod.

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

## Changelog

### Version 0.2
- Added flexible pod topology support
- Distribution switches are now optional per pod
- Support for 2-tier (Access → Core) and 3-tier (Access → Distribution → Core) topologies
- Mixed topology deployments in multi-pod environments
- Automatic topology detection and validation
- Enhanced validation messages showing topology type

### Version 0.1
- Initial release
- Multi-pod support
- IPv4/IPv6 underlay support
- DHCP relay configuration
- WAN integration with BGP
- Optic port configuration

---

*Last updated: 2025-01-13*
*Version: 0.2*
