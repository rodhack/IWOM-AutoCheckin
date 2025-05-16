# IWOM Daily Auto-Checkin Script (v3.1)

This script automates login and workday registration in the IWOM portal, intended for DXC employees with fixed working schedules.

## Features

- Auto login to [portal.bpocenter-dxc.com](httpsportal.bpocenter-dxc.com)
- Registers daily hours (0830–1830 Mon–Thu, 0800–1500 Fri)
- 2FA wait with 55s timeout
- Colored logging with file log history
- Auto-clean of old logs (keeps last 30)
- Configurable and ready for task scheduler

## Requirements

```bash
pip install selenium
pip install colorama
pip install webdriver-manager
