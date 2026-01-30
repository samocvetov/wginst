# Auto-provisioning Windows with Winget

This repository contains a PowerShell script to automate the installation of essential software on a fresh Windows 10/11 installation using the Windows Package Manager (winget).

## Features
* **Auto-update**: Updates all existing packages to the latest versions.
* **Batch Install**: Installs 16+ popular apps (Browsers, Messengers, Utilities).
* **Silent Mode**: All installations are performed silently without popups.
* **Self-cleanup**: The script deletes itself after completion to keep your folders clean.

## Quick Start (One-Liner)

Open **Command Prompt (CMD)** or **PowerShell** as **Administrator** and paste the following command:

```cmd
curl -L [https://raw.githubusercontent.com/samocvetov/winget-install/main/install.ps1](https://raw.githubusercontent.com/samocvetov/winget-install/main/install.ps1) -o setup.ps1 && powershell -ExecutionPolicy Bypass -File setup.ps1
