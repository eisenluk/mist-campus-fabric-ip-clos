# Juniper Mist - EVPN IP CLOS Fabric Builder

**Version:** 0.2
**Author:** Lukas Eisenberger (leisenberger@juniper.net)

> ⚠️ **DISCLAIMER - USE AT YOUR OWN RISK!** ⚠️
> This tool is in early development and provided "AS IS" without warranty of any kind, express or implied. Users must thoroughly test in a non-production environment before any deployment. By using this tool, you acknowledge that you assume all risks associated with its use. The developers expressly disclaim all liability for any damages, losses, or consequences arising from the use, misuse, or inability to use this tool, including but not limited to direct, indirect, incidental, consequential, or special damages, data loss, system failures, or business interruption. Use at your own risk.

## Overview

EVPN IP CLOS Fabric Builder is a tool designed to automate the configuration and deployment of EVPN-based IP CLOS fabric architectures for Juniper Mist. It simplifies the process of building scalable EVPN-VXLAN Campus Fabric with support for multi-pod architectures and various advanced networking features.

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
- **Flexible Pod Topologies**
  - **3-tier topology**: Access → Distribution → Core (traditional CLOS architecture)
  - **2-tier topology**: Access → Core (spine-leaf style, no distribution layer)
  - **Mixed deployments**: Different pods can use different topologies in the same fabric (2-tier mixed with 3-tier)
  - Distribution switches are now optional per pod
  - Automatic topology detection and validation

### Hardware Configuration (SWITCH sheet)
- **Optic Port Configuration**
  - Configurable port speeds
  - Channelization support for breakout cables

### Network Protocol Support (FABRIC sheet)
  - IPv4 underlay support
  - IPv6 underlay support
  - Dual-Stack IPv4 and IPv6 support in the overlay network

### Services (NETWORKS sheet)
- **DHCP Relay Support**
  - IPv4 DHCP relay configuration
  - IPv6 DHCP relay configuration

### WAN Integration (WAN sheet)
- **L3 Routing on Border Switches to external L3-handhoff**
  - WAN connectivity for border and core switches using eBGP
  - BGP peering with VRF-lite configuration when WAN worksheet is populated
  - BFD support
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
pip install -r requirements.txt
```

## Usage

1. Change values in excel spread sheet
2. Create an API token on your Mist Org
3. Copy API token and ORG-ID to the script mist-campus-fabric-ip-clos.py
4. Run the mist-campus-fabric-ip-clos.py

#### 2-Tier Topology (Access → Core)
For a simple spine-leaf deployment without a distribution layer:
- Define access switches with pod numbers
- Set access switch UPLINKS to point directly to core switches
- Do NOT create distribution switches for that pod

Example SWITCHES sheet configuration:
```
HOSTNAME    ROLE    POD    UPLINKS       UPLINK_PORTS        DOWNLINKS        DOWNLINK_PORTS
Core1       core           (leave empty) (leave empty)       Access1,Access2  ge-0/0/0,ge-0/0/1
Core2       core           (leave empty) (leave empty)       Access1,Access2   ge-0/0/0,ge-0/0/1
Access1     access  1      Core1,Core2   ge-0/0/0,ge-0/0/1   (empty)          (empty)
Access2     access  1      Core1,Core2   ge-0/0/0,ge-0/0/1   (empty)          (empty)
```

#### 3-Tier Topology (Access → Distribution → Core)
For traditional CLOS architecture with distribution layer:
- Define access switches with pod numbers
- Create distribution switches for that pod
- Set access switch UPLINKS to distribution switches
- Set distribution switch UPLINKS to core switches

Example SWITCHES sheet configuration:
```
HOSTNAME       ROLE          POD    UPLINKS                      UPLINK_PORTS        DOWNLINKS                    DOWNLINK_PORTS
Core1          core                 (leave empty)                (leave empty)       Distribution1,Distribution2  ge-0/0/0,ge-0/0/1
Core2          core                 (leave empty)                (leave empty)       Distribution1,Distribution2  ge-0/0/0,ge-0/0/1
Distribution1  distribution  1      Core1,Core2                  ge-0/0/0,ge-0/0/1   Access1,Access2             ge-0/0/2,ge-0/0/3
Distribution2  distribution  1      Core1,Core2                  ge-0/0/0,ge-0/0/1   Access1,Access2             ge-0/0/2,ge-0/0/3
Access1        access        1      Distribution1,Distribution2  ge-0/0/0,ge-0/0/1   (empty)                      (empty)
Access2        access        1      Distribution1,Distribution2  ge-0/0/0,ge-0/0/1   (empty)                      (empty)
```

#### Mixed Topology (Multiple Pods with Different Architectures)
You can combine both approaches in a single fabric:
```
HOSTNAME       ROLE          POD    UPLINKS                      UPLINK_PORTS        DOWNLINKS                                          DOWNLINK_PORTS
Core1          core                 (leave empty)                (leave empty)       Distribution1,Distribution2,Access1,Access2        ge-0/0/0-3
Core2          core                 (leave empty)                (leave empty)       Distribution1,Distribution2,Access1,Access2        ge-0/0/0-3
Distribution1  distribution  1      Core1,Core2                  ge-0/0/0,ge-0/0/1   Access3,Access4                                    ge-0/0/2,ge-0/0/3
Distribution2  distribution  1      Core1,Core2                  ge-0/0/0,ge-0/0/1   Access3,Access4                                   ge-0/0/2,ge-0/0/3
Access1        access        2      Core1,Core2                  ge-0/0/0,ge-0/0/1   (empty)                                            (empty)
Access2        access        2      Core1,Core2                  ge-0/0/0,ge-0/0/1   (empty)                                            (empty)
Access3        access        1      Distribution1,Distribution2  ge-0/0/0,ge-0/0/1   (empty)                                            (empty)
Access4        access        1      Distribution1,Distribution2  ge-0/0/0,ge-0/0/1   (empty)                                            (empty)
```
In this example:
- **Pod 1** uses 3-tier: Access3,Access4 → Distribution1,Distribution2 → Core1,Core2
- **Pod 2** uses 2-tier: Access1,Access2 → Core1,Core2 (no distribution switches)

The script will automatically detect and validate the topology for each pod.


#### Dedicated Serviceblock (Access → Distribution → Core)
If "core_as_border" in FABRIC is TRUE, then a dedicated Serviceblock as border is used:
```
HOSTNAME        ROLE          POD    UPLINKS                      UPLINK_PORTS        DOWNLINKS                                            DOWNLINK_PORTS
Serviceblock1   border               (leave empty)                (leave empty)       Core1,Core2                                          ge-0/0/0,ge-0/0/1
Serviceblock2   border               (leave empty)                (leave empty)       Core1,Core2                                          ge-0/0/0,ge-0/0/1
Core1           core                 Serviceblock1,Serviceblock2  ge-0/0/4,ge-0/0/5   Distribution1,Distribution2,Distribution3,Distribution4  ge-0/0/0,ge-0/0/1,ge-0/0/2,ge-0/0/3
Core2           core                 Serviceblock1,Serviceblock2  ge-0/0/4,ge-0/0/5   Distribution1,Distribution2,Distribution3,Distribution4  ge-0/0/0,ge-0/0/1,ge-0/0/2,ge-0/0/3
Distribution1   distribution  1      Core1,Core2                  ge-0/0/0,ge-0/0/1   Access1,Access2                                      ge-0/0/2,ge-0/0/3
Distribution2   distribution  1      Core1,Core2                  ge-0/0/0,ge-0/0/1   Access1,Access2                                      ge-0/0/2,ge-0/0/3
Distribution3   distribution  2      Core1,Core2                  ge-0/0/0,ge-0/0/1   Access3,Access4                                      ge-0/0/2,ge-0/0/3
Distribution4   distribution  2      Core1,Core2                  ge-0/0/0,ge-0/0/1   Access3,Access4                                      ge-0/0/2,ge-0/0/3
Access1         access        1      Distribution1,Distribution2  ge-0/0/0,ge-0/0/1   (empty)                                              (empty)
Access2         access        1      Distribution1,Distribution2  ge-0/0/0,ge-0/0/1   (empty)                                              (empty)
Access3         access        2      Distribution3,Distribution4  ge-0/0/0,ge-0/0/1   (empty)                                              (empty)
Access4         access        2      Distribution3,Distribution4  ge-0/0/0,ge-0/0/1   (empty)                                              (empty)
```
In this example:
- **Pod 1** uses Access1,Access1 → Distribution1,Distribution2 → Core1,Core2 → Serviceblock1,Serviceblock2 
- **Pod 2** uses Access1,Access1 → Distribution3,Distribution4 → Core1,Core2 → Serviceblock1,Serviceblock2 

The script will automatically detect and validate the topology for each pod.


## Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

## License

- **MIT**

## Support

For questions, issues, or suggestions, please contact:
- **Author:** Lukas Eisenberger
- **Email:** leisenberger@juniper.net

---

## Changelog

### Version 0.2
- Added flexible pod topology support
- Distribution switches are now optional per pod
- Support for 2-tier (Access → Core) and 3-tier (Access → Distribution → Core) topologies
- Mixed topology deployments in multi-pod environments
- Automatic topology detection and validation
- Enhanced validation messages showing topology type
- Multi-pod support
- IPv4/IPv6 underlay support
- DHCP relay configuration
- WAN integration with BGP
- Optic port configuration

---

*Last updated: 2025-11-12*
*Version: 0.2*
