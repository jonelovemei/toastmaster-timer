# Toastmaster Timer — Web Edition

A free, browser-based version of the Toastmaster Timer. Python runs **directly in the
browser** via [PyScript](https://pyscript.net/) (Pyodide / WebAssembly), so there is
**no server, no install, and no cost** — anyone just opens a URL. This sidesteps the
corporate firewall limitation of the desktop `.exe`.

## Files

| File | Purpose |
|------|---------|
| `index.html` | Page shell + browser-side helpers (Web Audio buzzer, Web Bluetooth, file download, PDF print) |
| `app.py` | All app logic in Python (timer, color phases, roster, bilingual UI, report export) |
| `styles.css` | Styling |
| `pyscript.toml` | PyScript config (loads `openpyxl` + the report template) |
| `TimeReport_Template.xlsx` | Report template — exports are built on top of it |

All five files must stay together in the same folder.

## Run locally (for testing)

PyScript must be served over HTTP (not opened as a `file://`):

```bash
cd web
python3 -m http.server 8000
# open http://localhost:8000/index.html in Chrome or Edge
```

## Deploy to GitHub Pages (free, permanent, HTTPS)

1. Create a GitHub repo (e.g. `toastmaster-timer`).
2. Put these files in the repo root (or in a `/docs` folder).
3. Push to GitHub:
   ```bash
   git init
   git add index.html app.py styles.css pyscript.toml TimeReport_Template.xlsx README.md
   git commit -m "Toastmaster Timer web edition"
   git branch -M main
   git remote add origin https://github.com/<you>/toastmaster-timer.git
   git push -u origin main
   ```
4. In the repo: **Settings → Pages**
   - **Source:** Deploy from a branch
   - **Branch:** `main`, folder `/ (root)` (or `/docs` if you used that folder)
   - Save.
5. After ~1 minute your app is live at:
   `https://<you>.github.io/toastmaster-timer/`

Share that URL — anyone with Chrome or Edge can use it. The report template is bundled,
so exports work for everyone with no extra setup.

> **Browser note:** Use **Chrome or Edge** (desktop). Web Bluetooth is only available in
> Chromium-based browsers and requires HTTPS — GitHub Pages provides HTTPS automatically.

## Features

- **Login page** (always English) with a **language dropdown** (English / 中文). The whole
  app UI then shows in the chosen language.
- **Timer** with green / yellow / red / **Max Red** phases, **Fast Forward** (press & hold),
  Start / Pause / Reset.
- **Buzzer** via Web Audio (1 beep on yellow, 2 on red, 3 on max red) — works on every OS,
  toggle with the speaker button.
- **Roster:** add / select / remove / clear, **Log End Time**, **Next Speaker**.
- **Reports** (XLSX / PDF / TXT / CSV), all built on the three-section template
  (Table Topic / Prepared Speech / Speech Evaluator) with auto **Qualified** (Yes/No) and
  automatic overflow rows when a section has more than 10 people.
- **Bluetooth LED control** via Web Bluetooth (see below).

## Bluetooth (BLE) hardware setup

The browser can only talk to **BLE** modules, not classic Bluetooth SPP. So the timer LED
controller must use a BLE module — **HM-10** or **HM-19** (replacing the old HC-05/06).

**Signal bytes sent to the module** (unchanged from the desktop version):

| Phase | Byte |
|-------|------|
| Green | `0x01` |
| Yellow | `0x02` |
| Red | `0x03` |
| Max Red | `0x04` |
| Reset | `0x00` |

**GATT profile used by the web app** (the HM-10/HM-19 default):

- Service UUID: `0xFFE0`
- Characteristic UUID: `0xFFE1` (write)

**On the microcontroller side:** wire the HM-10/HM-19 TX/RX to your MCU UART (default
9600 baud). The MCU reads one byte from the BLE UART and drives the LEDs:
`0x01`→green, `0x02`→yellow, `0x03`→red, `0x04`→max red (e.g. blinking red), `0x00`→off.
This is the same byte protocol the desktop app used, so existing firmware logic carries over —
only the radio module changes from HC-05/06 to HM-10/19.

**Using it in the app:** click **Bluetooth / 蓝牙**, pick your HM-10/HM-19 device in the
browser's pairing dialog. Once connected the button shows a check mark, and the app sends the
phase byte automatically on every color change. Click again to disconnect.
