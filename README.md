# Microsoft Office Cloud Login / Online Content Control for macOS

This project provides a macOS AppleScript (and standalone application) to **enable or disable Microsoft Office cloud login and online content features**.

It is designed for privacy-focused, offline, or security-restricted environments.

---

## ✨ Features

- Interactive startup dialog:
  - **Disable Cloud Login Features**
  - **Enable Cloud Login Features**
  - **Cancel**
- Verifies Microsoft Office presence before execution
- Safe to run multiple times (idempotent)
- No error if settings are already applied
- Available as:
  - AppleScript (`.scpt`)
  - macOS Application (`.app`)

---

## 🧩 What Is Modified

The script controls the `UseOnlineContent` preference:

```bash
defaults write com.microsoft.Word UseOnlineContent -integer 0/1
defaults write com.microsoft.Excel UseOnlineContent -integer 0/1
defaults write com.microsoft.Powerpoint UseOnlineContent -integer 0/1
```
* **0** → Cloud login and online content **disabled**
* **1** → Cloud login and online content **enabled**

---

## Usage

### Running the Script
1. Open the `.scpt` file with **Script Editor**.
2. Click **Run**.
3. Choose an option: `disable`, `enable`, or `cancel`.

### Running as an Application
1. Double-click the `.app`.
2. Make a selection in the dialog.
3. Done.

---

## Requirements
* **OS**: macOS
* **Software**: Microsoft Office installed
* **Permissions**: No administrator rights required

---

## Notes
> [!IMPORTANT]
> * Changes apply to the **current user**.
> * **Restart** Office applications to ensure settings are fully applied.
> * This does **not** uninstall or block core Office functionality.

---

## Deployment
* **End-users**: Manual execution.
* **IT Teams**: MDM deployment (Jamf, Intune, Kandji, etc.).
* **Shared Environments**: Ideal for classroom or lab computers requiring restricted mode.

---

## License
**MIT License** – free to use, modify, and redistribute.
