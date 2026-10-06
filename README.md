# WhatsApp Invite Sender

Playwright helper that pre-fills personalized wedding invite messages in WhatsApp Web. You press Send yourself.

## Placeholders

| Placeholder | Source |
|-------------|--------|
| `{{NAME}}` | Name column |
| `{{ROOM_NUMBER}}` | Room Number column |
| `{{ROOMMATES}}` | Other guests with same room (names only, bulleted) |
| `{{MOBILE_NUMBER}}` | Mobile number used for the chat |

Everything else in the template is static text.

## Expected Excel columns

`SR No`, `Name`, `Pariwar`, `Room Number`, `Mobile Number`, `Day`, `Shikh`

Only **SR No**, **Name**, **Room Number**, and **Mobile Number** are used. Mobile should include country code (e.g. `919876543210`). 10-digit numbers are treated as Indian (`91…`).

Use `--from-sr` / `--to-sr` to split work across machines (e.g. one 1–300, another 301–600).

## Playwright drafts (Python) — Windows / macOS / Linux

WhatsApp Web has **no Drafts folder**. This helper opens one WhatsApp window, pre-fills each chat’s message, and **never presses Send**. You send later by hand from those chats.

Same commands on every OS (use `python` or `python3` if that is what your system provides):

```bash
python setup_env.py
python run_drafts.py --file ./your-guests.xlsx --from-sr 1 --to-sr 300
```

| Flag | Meaning |
|------|---------|
| `--file` | Guest spreadsheet (required) |
| `--template` | Template file (default `template-default.txt`) |
| `--from-sr` / `--to-sr` | SR No range for splitting work |
| `--delay` | Average seconds between guests, randomized ±50% (default `8`) |
| `--batch-size` / `--batch-pause` | Longer break (default 180s, randomized) after every N drafts (default `25`) |
| `--max-per-run` | Max drafts per run (default `300`); re-run later to continue |
| `--no-resume` | Ignore progress and redo |

- `setup_env.py` creates `.venv`, installs packages, and downloads Chromium
- Login QR is saved in `.wa-profile/` (local to that machine; gitignored)
- Progress is saved under `playwright/.progress/`
- After drafts are created, stay in that WhatsApp window, open chats with draft text, and press **Send** yourself

### Ship to another computer

Copy this whole project folder (or clone the repo). On the new machine:

1. Install Python 3.10+ from https://www.python.org (tick **Add Python to PATH** on Windows)
2. Open Terminal / Command Prompt / PowerShell in the project folder
3. Run `python setup_env.py` then `python run_drafts.py --file ...`

Do **not** copy `.venv` or `.wa-profile` between machines — each OS/device should set those up locally and scan QR once.

Ban risk is lower than auto-send, but automation can still be flagged — use sensible delays and prefer a number you can afford to risk if doing hundreds at once.

### Invalid numbers

Numbers WhatsApp rejects are not counted as failures and don't count toward the 5-in-a-row stop. They are saved in `playwright/.progress/invalid-<guest file>.json` (by mobile, with SR and name) and skipped on later runs. Delete an entry from that file to retry it.

### Report and verification

At the end of every run the script:

1. **Verifies drafts** (read-only): scrolls the chat list and looks for rows showing "Draft". Each chat drafted in this run is marked `verified` (the chat title has the phone number), `likely` (the contact name matches) or `not_found`. A contact saved under a different name can show as `not_found` even if the draft exists.
2. **Writes a CSV report** to `playwright/.progress/report-<guest file>.csv` with one row per guest: status (`drafted`, `invalid`, `failed`, `pending`, `no_mobile`), verification, reason, shared-with SRs and time.

Guests marked `drafted` in the progress file are never drafted again on later runs.
