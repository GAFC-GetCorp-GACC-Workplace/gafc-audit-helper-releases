# GAFC Audit Helper - Release Package

**Version:** 9.0
**Release Date:** 2025-11-28
**License:** Required (contact for activation)

---

## Package Contents

```
GAFC_Audit_Helper_Release/
|
+-- chuyen_dau_ki_v9.xlam          # Main Excel Add-in
|
+-- install.bat                    # DOUBLE-CLICK to install
+-- uninstall.bat                  # DOUBLE-CLICK to uninstall
+-- update.bat                     # DOUBLE-CLICK to update
|
+-- scripts/                       # PowerShell scripts (auto-run by BAT files)
|   +-- install_audit_helper.ps1
|   +-- uninstall_audit_helper.ps1
|   +-- update_audit_helper.ps1
|
+-- README.md                      # This file
+-- LICENSE_KEY.txt                # Your license key (add manually)
+-- CHANGELOG.md                   # Version history
```

---

## Quick Start (3 Steps)

### Step 1: Install

**DOUBLE-CLICK:** `install.bat`

Wait for completion message, then press any key.

### Step 2: Activate License

1. Open Microsoft Excel
2. InputBox will appear asking for license key
3. Enter your key: `GAFC-AUDIT-AN-XXXXXXXXXXXX`
4. Click OK

### Step 3: Start Using

Add-in is now loaded! Access functions via:
- Excel Ribbon tabs
- Custom macros

---

## License Activation

**Format:** `GAFC-AUDIT-AN-[12 characters]`

**Example:** `GAFC-AUDIT-AN-B6ECF5D6B097`

**Requirements:**
- Internet connection (first activation only)
- Valid license key from provider
- Hardware fingerprinting will be recorded

**Offline Grace Period:** 7 days after last successful validation

---

## Update to Latest Version

**DOUBLE-CLICK:** `update.bat`

Script will:
1. Check for new version
2. Download if available
3. Backup current version
4. Install new version

**Note:** Requires GitHub repository setup (edit URL in `scripts/update_audit_helper.ps1`)

---

## Uninstall

**DOUBLE-CLICK:** `uninstall.bat`

Or manually delete:
```
%APPDATA%\Microsoft\Excel\XLSTART\chuyen_dau_ki_v9.xlam
```

---

## Troubleshooting

### License Issues

**Problem:** File closes immediately or "License khong hop le"

**Solution:**
1. Delete: `%APPDATA%\GAFC\audit_tool_license.txt`
2. Reopen Excel
3. Enter valid license key

### Installation Issues

**Problem:** BAT file doesn't work

**Solution:**
1. Right-click `install.bat`
2. Select "Run as administrator"

### Macro Security

**Problem:** Macros disabled

**Solution:**
1. File → Options → Trust Center → Trust Center Settings
2. Macro Settings → Enable all macros
3. Restart Excel

---

## Technical Info

**System Requirements:**
- Windows 7 or later
- Microsoft Excel 2010 or later
- PowerShell 3.0+
- Internet (for activation)

**License Server:** https://license-gafc-server.vercel.app
**App ID:** audit_tool

---

## Support

For issues:
- Check CHANGELOG.md for known issues
- Contact license provider for activation help
- Report bugs with error details

---

**GAFC Audit Helper v9.0 - All Rights Reserved**
