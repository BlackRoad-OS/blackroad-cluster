# BlackRoad Cluster

**Pure LAN Infrastructure • No Cloud Dependency**

```
┌─────────────────────────────────────────────────────────────┐
│                   BLACKROAD CLUSTER                         │
│              Raspberry Pi Mesh Network                      │
│                                                             │
│   lucidia ←──→ blackroad-pi ←──→ alice ←──→ aria           │
│   (primary)    (secondary)      (edge)      (edge)         │
│                                                             │
│         Your hardware. Your network. Your rules.           │
└─────────────────────────────────────────────────────────────┘
```

## Philosophy

**The Pis ARE the infrastructure.** Cloudflare is just a delivery layer - optional, not core.

- Direct LAN communication between all nodes
- ESP32 devices connect directly to Pis
- No internet dependency for local operations
- Services run on YOUR hardware

## Nodes

| Name | IP | Role | Services |
|------|-----|------|----------|
| lucidia | 192.168.4.38 | Primary | 300 web services |
| blackroad-pi | 192.168.4.64 | Secondary | TBD |
| alice | 192.168.4.49 | Edge | TBD |
| aria | 192.168.4.99 | Edge | TBD |

## Quick Start

### View Cluster Status
```bash
blackroad-cluster
```

### Access Services
```bash
# Dashboard (all 300 services)
open http://lucidia.local/

# Individual service
curl http://192.168.4.38/roadchat

# Health check
curl http://192.168.4.38/health
```

### From ESP32
```cpp
#include <HTTPClient.h>

HTTPClient http;
http.begin("http://192.168.4.38/api/status");
int code = http.GET();
// Direct to Pi - no internet needed
```

## Architecture

```
┌──────────────┐     ┌──────────────┐     ┌──────────────┐
│    ESP32     │────▶│     Pi       │────▶│   Browser    │
│  (CEO Hub)   │     │  (lucidia)   │     │   (local)    │
└──────────────┘     └──────────────┘     └──────────────┘
       │                    │                    │
       │         WiFi (192.168.4.x)              │
       └────────────────────┴────────────────────┘
                    Local Mesh
                    No Cloud
```

## Services Stack

On lucidia (192.168.4.38):
- **nginx** - Web server (80/443)
- **avahi** - mDNS (lucidia.local)
- **Java** - Hello World API (8888)
- **300 static sites** - /var/www/blackroad/

## Installation

### Install cluster command
```bash
cp bin/blackroad-cluster ~/bin/
chmod +x ~/bin/blackroad-cluster
```

### Set up new Pi node
```bash
./setup-node.sh <hostname> <ip>
```

## Files

```
blackroad-cluster/
├── bin/
│   └── blackroad-cluster     # Main status command
├── nginx/
│   ├── blackroad-lan         # HTTP config
│   └── blackroad-lan-ssl     # HTTPS config
├── setup/
│   ├── setup-node.sh         # New node setup
│   └── install-services.sh   # Service installer
└── README.md
```

## Access Methods

| Method | URL | Notes |
|--------|-----|-------|
| mDNS | http://lucidia.local | Requires avahi |
| Direct IP | http://192.168.4.38 | Always works |
| HTTPS | https://192.168.4.38 | Self-signed cert |
| SSH | ssh pi@lucidia.local | Ed25519 key |

## Related Repos

- [blackroad-os-iot-cluster](https://github.com/BlackRoad-OS/blackroad-os-iot-cluster) - ESP32 firmware
- [blackroad-mesh](https://github.com/BlackRoad-OS/blackroad-mesh) - Mesh networking

---

**BlackRoad Cluster** • Pure LAN • No Cloud • Your Infrastructure
