# 🔌 Network Port Inventory Tool

Automatically connects to Cisco switches via SSH, collects port-level data, and generates a formatted Excel report — one `.xlsx` file per switch, with four worksheets.

---

## 📋 What It Does

- Connects to each switch in your inventory over SSH
- Collects interface status, VLANs, duplex/speed, PoE draw, MAC addresses, IP addresses, last activity timestamps, and CDP neighbors
- Outputs a timestamped `.xlsx` file per switch into a `reports/` folder
- Each report has four sheets ready for verification and cleanup work

---

## 📄 Report Sheets

### Sheet 1 — Port Inventory
Full port-by-port breakdown of every interface on the switch.

| Column | Description |
|---|---|
| switch_name | Hostname of the switch |
| switch_ip | Management IP |
| snmp_location | SNMP location string |
| interface | Interface name (e.g. Gi1/0/1) |
| description | Interface description |
| connected | `up` or `down` — color coded green/red |
| trunk_mode | `trunk` or `access` |
| vlans | VLAN(s) assigned |
| voice_vlan | Voice VLAN if configured |
| duplex | Half / Full / Auto |
| speed | Port speed |
| power_watts | PoE watts drawn |
| max_sec_devices | Port security max MAC limit |
| last_input | Time since last inbound traffic |
| last_output | Time since last outbound traffic |
| end_device_mac | MAC(s) seen on port in dotted format, semicolon separated |
| end_device_ip | IP(s) resolved via ARP, semicolon separated |
| ip_source | `DHCP` or `Static` per MAC |

---

### Sheet 2 — Cleanup
Summary stats and a list of ports wasting capacity.

**Summary row (top):**
- Total Ports
- Used Ports (Up)
- Down Ports
- Utilization % (live formula)

**Idle port list:**
Any port that is `up` but where **both** `last_input` AND `last_output` are greater than 4 weeks. These are candidates for reclamation. Columns shown: Interface, Description, Connected, Last Input, Last Output, MAC Address.

---

### Sheet 3 — MAC Address Compare
One row per MAC address for easy verification. If a port has multiple MACs, each gets its own row.

| Column | Description |
|---|---|
| Interface | Port the MAC was learned on |
| Verification Command | `sh mac address-table address <MAC>` — ready to copy/paste |
| MAC Address | MAC in Cisco dotted format (e.g. `aabb.ccdd.eeff`) |
| VLAN | VLAN associated with that port |

> Example: if Gi1/0/3 has 3 MACs, it appears as 3 separate rows — one per MAC.

---

### Sheet 4 — CDP Neighbors
One row per CDP neighbor discovered on the switch, with a ready-to-run verification command.

| Column | Description |
|---|---|
| Neighbor Device ID | Device name as reported by CDP |
| Verification Command | `sh cdp neighbors \| include <NAME>` — ready to copy/paste |

---

## ⚙️ Requirements

- Python 3.8+
- SSH access to your switches
- The following Python packages:

```bash
pip install netmiko openpyxl pyyaml
```

---

## 🗂️ Inventory File Setup

Create a YAML file (e.g. `inventory.yaml`) — **do not commit this file to GitHub**, it contains credentials.

```yaml
credentials:
  username: admin
  password: yourpassword
  secret: yourenablesecret   # remove this line if enable secret is not needed

cores:
  - name: core-sw-01
    host: 192.168.1.1

switches:
  - name: access-sw-01
    host: 192.168.1.10
  - name: access-sw-02
    host: 192.168.1.11
```

> **Tip:** Add `inventory.yaml` to your `.gitignore` so it never gets accidentally uploaded.

---

## 🚀 Running the Script

```bash
python build_port_inventory.py --inventory inventory.yaml
```

Reports are saved to the `reports/` folder automatically:

```
reports/
  access-sw-01-20250427-143022.xlsx
  access-sw-02-20250427-143045.xlsx
```

Each file is named `<switch_name>-<timestamp>.xlsx` so runs never overwrite each other.

---

## ⚠️ Security Notes

- Never commit `inventory.yaml` to GitHub — it contains switch credentials
- Keep this repository **Private**
- If credentials need to be shared with a team, use a secrets manager or environment variables

---

## 🔧 Supported Hardware

Tested on Cisco IOS and IOS-XE switches. Interfaces recognized:

`GigabitEthernet`, `TenGigabitEthernet`, `FastEthernet`, `FortyGigabitEthernet`, `HundredGigE`, `FiveGigabitEthernet`, `Port-channel`
