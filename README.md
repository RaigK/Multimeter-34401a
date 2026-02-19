# HP/Agilent 34401A Multimeter – GUI

A Python desktop application for controlling and logging measurements with the HP/Agilent 34401A digital multimeter.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)

---

## Features

- **Auto-connect** on startup using saved interface address (GPIB / RS-232 / USB-GPIB)
- **Simulation mode** when no device is connected
- **Real-time plot** with autoscaling and live statistics overlay
- **Live Excel saving** – every measurement is written to disk immediately (no data loss on crash)
- **Optional statistics block** and **embedded chart** in Excel export (toggled via checkboxes)
- **Clean Excel table** – 4 columns only: `#`, `Time (s)`, `<Function> (<Unit>)`, `Clock`
- **Filename prefix + suffix** (Date/Time or sequential number)
- **Language:** Deutsch / English
- **Theme:** Dark / Bright
- All settings persisted in `settings.json`

---

## Requirements

```
pip install pyvisa matplotlib openpyxl numpy
```

| Package | Purpose |
|---------|---------|
| `pyvisa` | VISA instrument communication |
| `matplotlib` | Real-time measurement plot |
| `openpyxl` | Excel file export |
| `numpy` | Statistics calculation |

> The application runs in **simulation mode** if `pyvisa` or a VISA backend is not installed.

---

## Usage

```bash
python multimeter_34401A.py
```

1. Select the VISA interface address (e.g. `GPIB0::22::INSTR`) or use `SIMULATION`
2. Click **Connect** (or let it auto-connect on startup)
3. Choose measurement function, range and resolution
4. Press **▶ Start** to begin logging
5. Excel file is created automatically at measurement start and updated after every sample

---

## Excel Output

| Column | Content |
|--------|---------|
| `#` | Sample index |
| `Zeit (s)` / `Time (s)` | Elapsed time in seconds |
| `DC Spannung (V)` etc. | Measured value with unit |
| `Uhrzeit` / `Clock` | Absolute timestamp (HH:MM:SS.mmm) |

Optional additions (controlled by checkboxes in the sidebar):
- **Statistics** – mean, std. dev., min, max, peak-peak appended below data
- **Chart** – line chart on a separate sheet showing value vs. time

---

## Settings

All settings are stored in `settings.json` next to the script and loaded automatically on startup. The gear icon (⚙) in the main window opens the settings dialog for language and theme selection.

---

## Supported Instruments

Tested with:
- HP 34401A
- Agilent 34401A
- Compatible instruments that respond to `*IDN?` with `34401` or `34410`

---

## License

MIT License – see [LICENSE](LICENSE)
