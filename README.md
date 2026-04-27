# 🔌 Network Port Inventory Tool

Automatically connects to Cisco switches via SSH, collects port-level data, and generates a formatted Excel report — one `.xlsx` file per switch.

---

## 📋 What It Does

- Connects to each switch in your inventory over SSH
- Collects interface status, VLANs, duplex/speed, PoE draw, MAC addresses, IP addresses, and last activity timestamps
- Outputs a timestamped `.xlsx` file per switch into a `reports/` folder
- Each report has two sheets:
  - **Port Inventory** — full port-by-port breakdown
  - **Cleanup** — summary stats + list of ports that are UP but have been idle for more than 4 weeks

---

## 📊 Report Columns

| Column | Description |
|---|---|
| switch_name | Hostname of the switch |
| switch_ip | Management IP |
| snmp_location | SNMP location string |
| interface | Interface name (e.g. Gi1/0/1) |
| description | Interface description |
| connected | `up` or `down` |
| trunk_mode | `trunk` or `access` |
| vlans | VLAN(s) assigned |
| voice_vlan | Voice VLAN if configured |
| duplex | Half / Full / Auto |
| speed | Port speed |
| power_watts | PoE watts drawn |
| max_sec_devices | Port security max MAC limit |
| last_input | Time since last inbound traffic |
| last_output | Time since last outbound traffic |
| end_device_mac | MAC(s) seen on port (dotted format) |
| end_device_ip | IP(s) resolved via ARP |
| ip_source | `DHCP` or `Static` |

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
  secret: yourenablesecret

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
reports/
access-sw-01-20250427-143022.xlsx
access-sw-02-20250427-143045.xlsx

---

## 🧹 Cleanup Tab Logic

The **Cleanup** sheet flags ports that are wasting capacity:

- **Summary row:** Total ports, ports in use, ports down, and utilization %
- **Idle port list:** Any port that is `up` but where **both** `last_input` AND `last_output` are greater than 4 weeks — these are candidates for reclamation

---

## ⚠️ Security Notes

- Never commit `inventory.yaml` to GitHub — it contains switch credentials
- Keep this repository **Private**
- If credentials need to be shared with a team, use a secrets manager or environment variables

---

## 🔧 Supported Hardware

Tested on Cisco IOS and IOS-XE switches. Interfaces recognized:

`GigabitEthernet`, `TenGigabitEthernet`, `FastEthernet`, `FortyGigabitEthernet`, `HundredGigE`, `FiveGigabitEthernet`, `Port-channel`
