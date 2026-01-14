
# Network Automation Utilities — Cisco Access‑Port Audit ➜ Excel

A modular Python utility that connects to Cisco switches (optionally through an SSH jump host),
collects interface details, PoE information, and neighbor presence, then exports a clean,
**filters‑only** Excel workbook with a **SUMMARY** sheet and one sheet per device.

This README provides a full, practical guide to install, configure, run, and troubleshoot the tool.

---

## ✨ What this tool does
- **Audits access ports** across multiple Cisco switches in parallel.
- **Normalizes** interface names (e.g., `GigabitEthernet1/0/1` → `Gi1/0/1`) for cross‑command matching.
- **Enriches** interfaces with:
  - PoE draw and state (`show power inline`)
  - LLDP/CDP neighbor presence (`show lldp neighbors detail`, `show cdp neighbors detail`)
- **Classifies port mode & VLAN** using `show interfaces status` (access / trunk / routed).
- **Flags stale access ports** using conservative rules (see [Stale logic](#-stale-logic-how-ports-are-flagged)).
- **Exports Excel** with an at‑a‑glance **SUMMARY** and one sheet per device, with filters, frozen header, column auto‑size, and conditional formatting.
- **Shows a progress bar** while running concurrent device jobs.

> Design note: The workbook intentionally uses **filters only** (no Excel tables), and places **SUMMARY first**.

---

## 🧱 Project layout
```
.
├── main.py                 # CLI entry point
├── README.md               # This guide
└── Modules/
    ├── __init__.py
    ├── config.py           # Static configuration (e.g., JUMP_HOST)
    ├── credentials.py      # Secure credential retrieval + fallbacks
    ├── jump_manager.py     # SSH jump host (bastion) support
    └── netmiko_utils.py    # Netmiko connection wrapper(s)
```

> The `Modules/*.py` files encapsulate most environment‑specific behaviour (jump host,
> credential storage, and SSH connection settings).

---

## 📦 Requirements
- **Python**: 3.8+
- **Python packages** (install with pip):
  - `netmiko`
  - `paramiko`
  - `pandas`
  - `openpyxl`
  - `pywin32` *(Windows only; used for Windows Credential Manager integration)*

```bash
pip install netmiko paramiko pandas openpyxl pywin32
```

### Optional but recommended
- **TextFSM templates** (NTC templates) for robust parsing of `show interfaces` when `use_textfsm=True`.
  - If templates are available and the `NET_TEXTFSM` environment variable points to them, parsing accuracy improves.
  - If not available, the script still works and falls back where needed (e.g., it has its own
    fixed‑width parser for `show interfaces status`).

---

## 🔐 Credentials & Security
The script retrieves device credentials using `Modules/credentials.py`:

- **Primary**: Windows Credential Manager (target name expected by default: `MyApp/ADM`).
- **Fallback**: Interactive prompt for username and password (secure, not echoed).
- **Enable secret**: Retrieved by `get_enable_secret()` if configured, otherwise not required.

> If you are running on Linux/macOS, ensure `credentials.py` prompts for credentials or implements your
> preferred secure store. On Windows, `pywin32` enables Credential Manager access.

**Never** hard‑code credentials in the repository. Use the secure store or environment prompts.

---

## 🛰️ Jump host (bastion) behaviour
- `main.py` reads `JUMP_HOST` from `Modules/config.py`.
- **Default**: the script **uses the jump host** if `--direct` is *not* supplied.
- `--direct` will skip the jump host entirely and attempt direct SSH connections.

```python
# Modules/config.py (example)
JUMP_HOST = "jump-gateway.example.com"  # or None to disable by default
```

> The `JumpManager` maintains a persistent SSH session to the bastion and proxies device connections through it.

---

## 🗂️ Device list file
Provide a plain‑text file with **one device per line**. Lines that are blank or start with `#` are ignored.

```
# devices.txt
10.10.10.11
10.10.10.12  # inline comments are not parsed; this whole token must be a host/IP only
core-switch-01
edge-sw-22
```

> Hostnames must be resolvable from the machine (or via the jump host, depending on your SSH setup).

---

## 🚀 Quick start
1. **Install dependencies** (see [Requirements](#-requirements)).
2. **Create `devices.txt`** with your targets (see [Device list file](#%EF%B8%8F-device-list-file)).
3. **(Optional) Configure** `Modules/config.py` with your `JUMP_HOST`.
4. **Run** the audit:

```bash
# Using jump host from Modules/config.py
python -m main --devices devices.txt --output access_port_audit.xlsx

# Direct connections (no bastion), 5 workers, different stale threshold
python -m main --direct -w 5 --stale-days 60 -d devices.txt -o results.xlsx

# Verbose debugging
python -m main --debug --devices devices.txt
```

The script prints an event‑driven progress bar like:
```
Progress: [██████████░░░░░░░░░░░░░░] 12/30 started: 15/30
```

On completion, the Excel workbook is written to the filename you specify (default `audit.xlsx`).

---

## 🧭 CLI reference
`main.py` exposes the following command‑line options:

```text
--devices, -d    (required)  Path to the devices file (one IP/hostname per line; '#' comments allowed)
--output,  -o    (optional)  Output Excel file name. Default: audit.xlsx
--workers, -w    (optional)  Max concurrent device sessions (threads). Default: 10
--stale-days     (optional)  Days threshold for stale access ports. 0 disables stale flagging. Default: 30
--direct         (optional)  Connect directly (do not use jump host)
--debug          (optional)  Enable verbose logging/prints
```

### Required vs optional
- **Required**: `--devices`
- **Optional**: everything else

---

## 🧪 What the script collects
For each device the script attempts to gather:

- **Hostname** (from `show running-config | include ^hostname` or CLI prompt fallback)
- **Interfaces** via TextFSM (`show interfaces`) when available
- **Port mode & VLAN** via a robust, fixed‑width parser of `show interfaces status`
- **PoE details**: admin/oper state, power draw (W), class, device (`show power inline`)
- **Neighbor presence**: LLDP/CDP seen on the port (boolean)
- **Error counters**: input, output, CRC (from `show interfaces` parsed data)
- **Activity indicator**: "Last input" time (seconds parsed when present)

---

## 🛑 Stale logic — how ports are flagged
A conservative approach is used **only for ports in `access` mode** and when `--stale-days > 0`:

- **If Status = `connected`** → mark **stale = True** **only** if `Last input ≥ <stale-days>`.
- **If Status ≠ `connected`** → mark **stale = True** when **both** conditions hold:
  1. **No PoE draw** (PoE power is blank/`-`/0.0), and
  2. **No LLDP/CDP neighbor present** on the port.

This tends to avoid false positives on trunk/routed ports and on access ports actively in use.

> You can disable stale flagging entirely by setting `--stale-days 0`.

---

## 📤 Excel output structure
The workbook contains:

### 1) `SUMMARY` sheet (first)
- One row per device, with a final **TOTAL** row (sums numeric columns).
- Columns include:
  - `Device`, `Mgmt IP`, `Total Ports (phy)`
  - `Access Ports`, `Trunk Ports`, `Routed Ports`
  - `Connected`, `Not Connected`, `Admin Down`, `Err-Disabled`
  - `% Access of Total`, `% Trunk of Total`, `% Routed of Total`, `% Connected of Total`

### 2) One sheet per device
Columns typically include (when available):
- `Device`, `Mgmt IP`, `Interface` (long form), `Description`
- `Status` (normalized: connected / notconnect / administratively down / err-disabled)
- `AdminDown`, `Connected`, `ErrDisabled` (booleans for quick filters)
- `Mode` (access/trunk/routed), `VLAN`
- `Duplex`, `Speed`, `Type`
- `Input Errors`, `Output Errors`, `CRC Errors`
- `Last Input` (raw text)
- `PoE Power (W)`, `PoE Oper`, `PoE Admin`
- `LLDP/ CDP Neighbor` (boolean)
- `Stale (≥<N> d)` (boolean)

### Formatting
- **Frozen header** (`A2`) and **AutoFilter** across all columns
- **Auto‑sized columns** with sensible min/max widths
- **Conditional formatting**:
  - `Status = connected` → green
  - `Status = notconnect` / `err-disabled` → red
  - `Status = administratively down` → grey
  - Any `*Errors` > 0 → red
  - `PoE Power (W)` > 0 → green
  - `Stale (≥N d)` = TRUE → red

---

## ⚙️ Performance & concurrency
- Uses a `ThreadPoolExecutor` with `--workers` threads (default 10).
- An **event‑driven progress bar** updates as jobs start/finish.
- Each device is independent; a failure on one does **not** stop others.

---

## 🐛 Logging, debug, and errors
- Add `--debug` to surface additional prints (e.g., enable mode attempts, jump host info, file counts).
- Per‑device errors are captured into the device’s summary row (and a minimal sheet may be created
  with the error text so the workbook always reflects all devices).

Common runtime issues & tips:
- **Authentication failures** → check Credential Manager entry or typed credentials.
- **SSH connectivity** → verify reachability from the workstation *or* via the jump host.
- **TextFSM templates missing** → parsing still proceeds, but some fields may be blank.
- **Channel/line rate limits** on older devices → consider lowering `--workers`.

---

## 🔧 Extending and customizing
- **Credentials**: adapt `Modules/credentials.py` to your environment (Linux keyring, Azure Key Vault, etc.).
- **Jump host**: tune `Modules/jump_manager.py` (keep‑alive, ciphers, auth methods) as needed.
- **Connection behaviour**: modify `Modules/netmiko_utils.py` for device types, timeouts, or SSH options.
- **Output columns**: adjust record construction in `main.py` (search for `detailed.append({...})`).
- **Conditional formatting**: tweak `_format_worksheet()` in `main.py`.

---

## 🔒 Security considerations
- Prefer secure stores over plaintext.
- Limit who can run the tool and who can read the generated Excel.
- When using a jump host, ensure strong authentication and proper network segmentation.

---

## 🧩 Compatibility
- Target devices: Cisco IOS/IOS‑XE access and distribution switches reachable via SSH.
- The tool relies on Netmiko; specify the right device type(s) inside `netmiko_utils.py`.
- TextFSM/NTC templates significantly improve interface parsing fidelity but are not strictly required.

---

## ✅ Examples
```bash
# Basic, with jump host
python -m main -d devices.txt -o audit.xlsx

# Direct (no bastion), 20 workers, stale disabled
python -m main --direct -w 20 --stale-days 0 -d devices.txt -o audit.xlsx

# Conservative concurrency, higher stale threshold, verbose
python -m main -w 4 --stale-days 90 --debug -d devices.txt -o siteA.xlsx
```

---

## 🧠 FAQs
**Q: Do I need NTC TextFSM templates?**  
*A:* They are recommended for better `show interfaces` parsing. Without them, the script still works and
uses its internal parser for `show interfaces status` and best‑effort logic elsewhere.

**Q: Where do credentials come from?**  
*A:* On Windows, from Credential Manager (default target `MyApp/ADM`). Otherwise, you are prompted interactively
or you can adapt `credentials.py` to your secret store.

**Q: How is `Mode` determined?**  
*A:* From `show interfaces status`: if VLAN column is `trunk`/`rspan` → `trunk`; if `routed` → `routed`; otherwise `access`.

**Q: How is a port considered *stale*?**  
*A:* Only for access ports and when `--stale-days > 0`. Connected ports are flagged stale only if `Last input ≥ N days`.
Disconnected ports require **both** no PoE draw and no LLDP/CDP neighbor to be flagged stale.

---

## 📝 License
MIT License

## 👤 Author
Christopher Davies