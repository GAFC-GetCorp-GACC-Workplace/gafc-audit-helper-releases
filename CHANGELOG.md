# CHANGELOG - GAFC Audit Helper

## Version 9.0 - 2025-11-28

### Features
- License validation system with hardware fingerprinting
- Auto-install via BAT files (double-click installation)
- Auto-update mechanism
- Offline grace period (7 days)
- Ultra-clean code (no blank lines)

### Fixed
- VB6 form compatibility issue
- APP_ID mismatch (audit_tool vs audit-tool)
- Attribute lines corruption
- Vietnamese encoding issues (ongoing)

### Installation
- Double-click `install.bat` to install
- Double-click `uninstall.bat` to remove
- Double-click `update.bat` to update

### Technical
- 27 VBA modules
- License server: https://license-gafc-server.vercel.app
- Cache: 24 hours
- Offline grace: 7 days

---

## Version 8.x - Legacy
- Manual distribution without license system

---

For full documentation, see README.md
