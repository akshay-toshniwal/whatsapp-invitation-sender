#!/usr/bin/env python3
"""
Create WhatsApp Web drafts only — never clicks Send.

Cross-platform (Windows / macOS / Linux):

  python setup_env.py
  python run_drafts.py --file ./guests.xlsx --from-sr 1 --to-sr 300

Or call this file directly after setup:

  python playwright/create_drafts.py --file ./guests.xlsx --from-sr 1 --to-sr 300
"""

from __future__ import annotations

import argparse
import csv
import json
import random
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
# Per-machine dirs (do not copy between Windows/macOS/Linux)
PROFILE_DIR = ROOT / ".wa-profile"
PROGRESS_DIR = ROOT / "playwright" / ".progress"

HEADER_MAP = {
    "name": "name",
    "roomnumber": "room",
    "roomno": "room",
    "room": "room",
    "mobilenumber": "mobile",
    "mobile": "mobile",
    "phone": "mobile",
    "phonenumber": "mobile",
    "srno": "sr",
    "sr": "sr",
    "sno": "sr",
    "serialno": "sr",
    "serialnumber": "sr",
}

INVALID_RE = re.compile(
    r"phone number shared via url is invalid|isn.t on WhatsApp|is not on WhatsApp|not registered on WhatsApp",
    re.I,
)

COMBINE_SEPARATOR = "\n\n--------------------\n\n"

REQUIRED = ("sr", "name", "room", "mobile")


def normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def parse_sr(value) -> int | None:
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None


def clean_phone(raw) -> str:
    digits = re.sub(r"\D", "", str(raw or ""))
    if len(digits) == 10:
        digits = "91" + digits
    return digits


def read_sheet(path: Path) -> list[dict]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        import csv

        with path.open(newline="", encoding="utf-8-sig") as f:
            return list(csv.DictReader(f))

    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h or "").strip() for h in rows[0]]
    out = []
    for row in rows[1:]:
        item = {}
        for i, header in enumerate(headers):
            if not header:
                continue
            item[header] = "" if row[i] is None else row[i]
        out.append(item)
    return out


def pick_columns(row: dict) -> dict:
    mapped = {}
    for raw_key, value in row.items():
        key = HEADER_MAP.get(normalize_header(raw_key))
        if key:
            mapped[key] = "" if value is None else str(value).strip()
    return mapped


def validate_columns(rows: list[dict]) -> None:
    if not rows:
        raise SystemExit("Spreadsheet is empty.")
    found = set()
    for header in rows[0].keys():
        key = HEADER_MAP.get(normalize_header(header))
        if key:
            found.add(key)
    missing = [name for name in REQUIRED if name not in found]
    if missing:
        raise SystemExit(
            f"Missing required column(s): {', '.join(missing)}. "
            f"Found: {', '.join(rows[0].keys())}"
        )


def build_guests(rows: list[dict]) -> list[dict]:
    validate_columns(rows)
    partial = []
    for i, row in enumerate(rows):
        cols = pick_columns(row)
        if not (cols.get("name") or cols.get("mobile")):
            continue
        sr_num = parse_sr(cols.get("sr"))
        partial.append(
            {
                "sr": str(sr_num if sr_num is not None else cols.get("sr") or i + 1),
                "sr_num": sr_num if sr_num is not None else i + 1,
                "name": cols.get("name") or "",
                "room": cols.get("room") or "",
                "mobile": clean_phone(cols.get("mobile")),
                "roommates": [],
            }
        )

    by_room: dict[str, list[str]] = {}
    for guest in partial:
        room = guest["room"].strip()
        if not room:
            continue
        by_room.setdefault(room, []).append(guest["name"])

    for guest in partial:
        room = guest["room"].strip()
        names = [n for n in by_room.get(room, []) if n and n != guest["name"]]
        # unique, keep order
        seen = set()
        unique = []
        for name in names:
            if name not in seen:
                seen.add(name)
                unique.append(name)
        guest["roommates"] = unique

    partial.sort(key=lambda g: g["sr_num"])
    return partial


def filter_by_sr(guests: list[dict], from_sr: int | None, to_sr: int | None) -> list[dict]:
    out = []
    for guest in guests:
        if from_sr is not None and guest["sr_num"] < from_sr:
            continue
        if to_sr is not None and guest["sr_num"] > to_sr:
            continue
        out.append(guest)
    return out


def roommate_list_text(names: list[str]) -> str:
    if not names:
        return "- कोई रूम सहयोगी नहीं"
    return "\n".join(f"- {name}" for name in names)


def build_message(template: str, guest: dict) -> str:
    values = {
        "NAME": guest["name"],
        "ROOM_NUMBER": guest["room"],
        "ROOMMATES": roommate_list_text(guest["roommates"]),
        "MOBILE_NUMBER": guest["mobile"],
    }
    text = template
    for key, value in values.items():
        text = text.replace(f"{{{{{key}}}}}", value)
        text = text.replace(f"<{key}>", value)
    text = text.replace("<ROOM NUMBER>", values["ROOM_NUMBER"])
    text = text.replace("[LIST OF ROOMATE]", values["ROOMMATES"])
    text = text.replace("[LIST OF ROOMMATE]", values["ROOMMATES"])
    text = text.replace("[LIST OF ROOMMATES]", values["ROOMMATES"])
    return text


def guest_key(guest: dict) -> str:
    return f"{guest['sr']}|{guest['mobile']}|{guest['name']}".lower()


def load_progress_file(path: Path, default: dict) -> dict:
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return default


def load_progress(path: Path) -> dict:
    return load_progress_file(path, {"drafted": {}})


def load_invalid(path: Path) -> dict:
    return load_progress_file(path, {"invalid": {}})


def save_progress(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data["updatedAt"] = datetime.now(timezone.utc).isoformat()
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


HONORIFICS = {"shri", "shriman", "sau", "smt", "shrimati", "kum", "ku", "chi", "dr", "mr", "mrs", "ms"}

SCRAPE_CHAT_LIST_JS = """() => {
    const out = [];
    for (const row of document.querySelectorAll('#pane-side [role="listitem"], #pane-side [role="row"]')) {
        const lines = (row.innerText || '').split('\\n').map(l => l.trim()).filter(Boolean);
        if (!lines.length) continue;
        out.push({title: lines[0], draft: lines.some(l => /^draft\\b/i.test(l))});
    }
    return out;
}"""


def name_tokens(text: str) -> set[str]:
    words = re.findall(r"[^\W\d_]{3,}", str(text).lower())
    return {w for w in words if w not in HONORIFICS}


def tokens_match(a: str, b: str) -> bool:
    return a == b or (min(len(a), len(b)) >= 4 and (a.startswith(b) or b.startswith(a)))


def match_draft(mobile: str, members: list[dict], titles: list[str]) -> str:
    """verified = chat title carries the phone number; likely = names match;
    not_found = no draft row matched."""
    tail = mobile[-10:]
    for title in titles:
        digits = re.sub(r"\D", "", title)
        if len(digits) >= 10 and digits.endswith(tail):
            return "verified"
    guest_tokens = set().union(*(name_tokens(m["name"]) for m in members))
    for title in titles:
        t_tokens = name_tokens(title)
        hits = {t for t in t_tokens if any(tokens_match(t, g) for g in guest_tokens)}
        if len(hits) >= 2:
            return "likely"
    return "not_found"


def collect_draft_titles(page) -> list[str]:
    """Scroll the chat list and return titles of rows that show a Draft."""
    pane = page.locator("#pane-side")
    pane.wait_for(state="visible", timeout=30_000)
    titles: dict[str, bool] = {}
    pane.evaluate("el => { el.scrollTop = 0; }")
    page.wait_for_timeout(1000)
    for _ in range(400):
        for row in page.evaluate(SCRAPE_CHAT_LIST_JS):
            if row["draft"]:
                titles[row["title"]] = True
        before = pane.evaluate("el => el.scrollTop")
        pane.evaluate("el => { el.scrollTop += Math.max(200, el.clientHeight * 0.8); }")
        page.wait_for_timeout(400)
        if pane.evaluate("el => el.scrollTop") == before:
            break
    return list(titles)


def verify_drafts(page, drafted: list[tuple[str, list[dict]]]) -> dict[str, str]:
    print("\nVerifying drafts in the chat list (read-only)…")
    try:
        titles = collect_draft_titles(page)
    except Exception as exc:  # noqa: BLE001
        print(f"Verification skipped: {exc}")
        return {}
    print(f"Chats showing a draft: {len(titles)} · drafts created this run: {len(drafted)}")
    results = {mobile: match_draft(mobile, members, titles) for mobile, members in drafted}
    counts = {k: sum(1 for v in results.values() if v == k) for k in ("verified", "likely", "not_found")}
    print(f"  verified (number match): {counts['verified']} · likely (name match): {counts['likely']} · not found: {counts['not_found']}")
    for mobile, members in drafted:
        if results[mobile] == "not_found":
            print(f"  not found: {' + '.join('SR ' + str(m['sr']) + ' ' + m['name'] for m in members)} ({mobile})")
    if counts["not_found"]:
        print("  (a contact saved under a different name can show as not found even if the draft exists)")
    return results


def write_report(path: Path, guests: list[dict], progress: dict, invalid: dict) -> None:
    by_mobile: dict[str, list[dict]] = {}
    for g in guests:
        if g["mobile"]:
            by_mobile.setdefault(g["mobile"], []).append(g)
    drafted = progress.get("drafted", {})
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["SR", "Name", "Room", "Mobile", "Status", "Verified", "Reason", "Shared with SR", "Updated at"])
        for g in guests:
            entry = drafted.get(guest_key(g), {})
            reason = entry.get("reason", "")
            if not g["mobile"]:
                status = "no_mobile"
            elif g["mobile"] in invalid.get("invalid", {}):
                status, reason = "invalid", "number not on WhatsApp"
                entry = invalid["invalid"][g["mobile"]]
            else:
                status = entry.get("status", "pending")
            shared = ", ".join(str(o["sr"]) for o in by_mobile.get(g["mobile"], []) if o is not g)
            w.writerow([g["sr"], g["name"], g["room"], g["mobile"], status,
                        entry.get("verified", ""), reason, shared, entry.get("at", "")])


def wait_for_login(page) -> None:
    print("Waiting for WhatsApp Web login (scan QR if shown)…")
    page.wait_for_selector(
        '[contenteditable="true"][data-tab="3"], [data-testid="chat-list"], #side',
        timeout=0,
    )
    print("Logged in.")


def navigate_in_app(page, url: str, first: bool) -> bool:
    """Open a chat URL. After the first load, click a link so WhatsApp Web
    routes in-app instead of reloading (a reload discards unsent drafts).
    Returns False if a full reload was detected."""
    if first:
        page.goto(url, wait_until="domcontentloaded", timeout=120_000)
        page.evaluate("window.__waDraftMarker = true")
        return True
    page.evaluate(
        """(url) => {
            window.__waDraftMarker = true;
            const a = document.createElement('a');
            a.href = url;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            a.remove();
        }""",
        url,
    )
    page.wait_for_timeout(1500)
    try:
        return bool(page.evaluate("window.__waDraftMarker === true"))
    except Exception:
        return False


def open_draft(page, guest: dict, message: str, first: bool) -> tuple[bool, str | None, bool]:
    from urllib.parse import quote

    url = (
        f"https://web.whatsapp.com/send?phone={guest['mobile']}"
        f"&text={quote(message)}"
    )
    in_app = navigate_in_app(page, url, first)

    # Wait for whichever appears first: the compose box, or WhatsApp's
    # "isn't on WhatsApp" / "invalid" dialog.
    compose = page.locator('div[contenteditable="true"][data-tab="10"], footer div[contenteditable="true"]')
    invalid = page.get_by_text(INVALID_RE)
    box = None
    deadline = time.monotonic() + 90
    while time.monotonic() < deadline:
        try:
            if invalid.first.is_visible():
                try:  # dismiss the dialog so the next chat can open
                    page.get_by_role("button", name=re.compile(r"^ok$", re.I)).click(timeout=3000)
                except Exception:
                    pass
                return False, "invalid_phone", in_app
            if compose.last.is_visible():
                box = compose.last
                break
        except Exception:
            pass
        page.wait_for_timeout(500)
    if box is None:
        raise TimeoutError("chat did not open (no compose box or invalid-number dialog)")

    current = ""
    try:
        current = (box.inner_text() or "").strip()
    except Exception:
        current = ""

    if not current:
        box.click()
        page.keyboard.insert_text(message)

    # Critical: do NOT click Send.
    return True, None, in_app


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create WhatsApp Web message drafts (never sends)."
    )
    parser.add_argument("--file", required=True, help="Guest spreadsheet (.xlsx/.xls/.csv)")
    parser.add_argument(
        "--template",
        default=str(ROOT / "template-default.txt"),
        help="Message template file",
    )
    parser.add_argument("--from-sr", type=int, default=None, dest="from_sr", help="Start SR No")
    parser.add_argument("--to-sr", type=int, default=None, dest="to_sr", help="End SR No")
    parser.add_argument(
        "--delay", type=float, default=8.0,
        help="Average seconds between guests; randomized +/-50%% (default 8)",
    )
    parser.add_argument(
        "--batch-size", type=int, default=25, dest="batch_size",
        help="Pause for a longer break after this many drafts (default 25, 0 = off)",
    )
    parser.add_argument(
        "--batch-pause", type=float, default=180.0, dest="batch_pause",
        help="Average seconds for the batch break, randomized +/-30%% (default 180)",
    )
    parser.add_argument(
        "--max-per-run", type=int, default=300, dest="max_per_run",
        help="Stop after this many drafts in one run (default 300)",
    )
    parser.add_argument(
        "--no-resume",
        action="store_true",
        help="Ignore saved progress and redo all",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    file_path = Path(args.file).expanduser().resolve()
    template_path = Path(args.template).expanduser().resolve()

    if not file_path.exists():
        print(f"File not found: {file_path}", file=sys.stderr)
        return 1
    if not template_path.exists():
        print(f"Template not found: {template_path}", file=sys.stderr)
        return 1

    template = template_path.read_text(encoding="utf-8")
    guests = filter_by_sr(build_guests(read_sheet(file_path)), args.from_sr, args.to_sr)
    if not guests:
        print("No guests in the selected SR range.", file=sys.stderr)
        return 1

    progress_path = PROGRESS_DIR / f"drafts-{file_path.name}.json"
    progress = {"drafted": {}} if args.no_resume else load_progress(progress_path)
    invalid_path = PROGRESS_DIR / f"invalid-{file_path.name}.json"
    invalid = load_invalid(invalid_path)

    # One chat per phone number: guests sharing a number are combined into
    # a single draft, separated by a divider line.
    groups: dict[str, list[dict]] = {}
    for guest in guests:
        if guest["mobile"]:
            groups.setdefault(guest["mobile"], []).append(guest)

    def is_drafted(g: dict) -> bool:
        return progress.get("drafted", {}).get(guest_key(g), {}).get("status") == "drafted"

    pending = []  # list of (mobile, members)
    done_guests = invalid_guests = merged_guests = 0
    for mobile, members in groups.items():
        merged_guests += len(members) - 1
        if mobile in invalid.get("invalid", {}):
            invalid_guests += len(members)
        elif all(is_drafted(m) for m in members):
            done_guests += len(members)
        else:
            pending.append((mobile, members))
            if len(members) > 1:
                names = ", ".join(f"SR {m['sr']} {m['name']}" for m in members)
                print(f"Combining into one draft ({mobile}): {names}")
    no_mobile = sum(1 for g in guests if not g["mobile"])
    total_pending = len(pending)
    if args.max_per_run > 0:
        pending = pending[: args.max_per_run]

    print(
        f"Guests in range: {len(guests)} · already drafted: {done_guests} · "
        f"invalid numbers (skipped): {invalid_guests} · no mobile: {no_mobile} · "
        f"merged into shared chats: {merged_guests} · "
        f"chats to draft: {total_pending}"
        + (f" (this run: {len(pending)})" if len(pending) != total_pending else "")
    )
    print("NOTE: WhatsApp has no Drafts inbox. Each chat keeps its own unsent text.")
    print("This script will NOT press Send.\n")

    PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1280, "height": 900},
            args=["--disable-blink-features=AutomationControlled"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto("https://web.whatsapp.com/", wait_until="domcontentloaded", timeout=120_000)
        wait_for_login(page)

        ok_count = 0
        fail_count = 0
        drafted_groups: list[tuple[str, list[dict]]] = []
        invalid_count = 0
        consecutive_fail = 0
        warned_reload = False

        for i, (mobile, members) in enumerate(pending):
            guest = members[0]
            keys = [guest_key(m) for m in members]
            message = COMBINE_SEPARATOR.join(build_message(template, m) for m in members)
            label = " + ".join(f"SR {m['sr']} {m['name']}" for m in members)
            print(
                f"[{i + 1}/{len(pending)}] {label} "
                f"({mobile}) … ",
                end="",
                flush=True,
            )
            try:
                ok, reason, in_app = open_draft(page, guest, message, first=(i == 0))
                if not in_app and not warned_reload:
                    warned_reload = True
                    print("\n  WARNING: WhatsApp reloaded; earlier drafts may be lost. Test with a few guests first.")
                if not ok and reason == "invalid_phone":
                    invalid.setdefault("invalid", {})[mobile] = {
                        "sr": guest["sr"],
                        "name": " + ".join(m["name"] for m in members),
                        "at": datetime.now(timezone.utc).isoformat(),
                    }
                    save_progress(invalid_path, invalid)
                    invalid_count += 1
                    print("invalid number, ignored")
                elif ok:
                    for key in keys:
                        progress.setdefault("drafted", {})[key] = {
                            "status": "drafted",
                            "at": datetime.now(timezone.utc).isoformat(),
                        }
                    drafted_groups.append((mobile, members))
                    ok_count += 1
                    consecutive_fail = 0
                    print("draft OK")
                else:
                    for key in keys:
                        progress.setdefault("drafted", {})[key] = {
                            "status": "failed",
                            "reason": reason,
                            "at": datetime.now(timezone.utc).isoformat(),
                        }
                    fail_count += 1
                    consecutive_fail += 1
                    print(f"FAILED ({reason})")
                save_progress(progress_path, progress)
            except Exception as exc:  # noqa: BLE001 - keep going through the list
                for key in keys:
                    progress.setdefault("drafted", {})[key] = {
                        "status": "failed",
                        "reason": str(exc),
                        "at": datetime.now(timezone.utc).isoformat(),
                    }
                save_progress(progress_path, progress)
                fail_count += 1
                consecutive_fail += 1
                print(f"FAILED ({exc})")

            if consecutive_fail >= 5:
                print("\nStopping: 5 failures in a row. Check the browser/login, then re-run to resume.")
                break

            if i < len(pending) - 1:
                if args.batch_size > 0 and (i + 1) % args.batch_size == 0:
                    pause = args.batch_pause * random.uniform(0.7, 1.3)
                    print(f"  Batch break: {pause:.0f}s…")
                    time.sleep(pause)
                else:
                    time.sleep(args.delay * random.uniform(0.5, 1.5))

        if drafted_groups:
            results = verify_drafts(page, drafted_groups)
            for mobile, members in drafted_groups:
                if mobile in results:
                    for m in members:
                        progress["drafted"][guest_key(m)]["verified"] = results[mobile]
            save_progress(progress_path, progress)

        report_path = PROGRESS_DIR / f"report-{file_path.stem}.csv"
        write_report(report_path, guests, progress, invalid)

        print(f"\nDone. Drafted: {ok_count} · Invalid numbers: {invalid_count} · Failed: {fail_count}")
        if invalid_count:
            print(f"Invalid numbers file: {invalid_path}")
        print(f"Progress file: {progress_path}")
        print(f"Report (CSV): {report_path}")
        print(
            """
Next steps (manual send — safer):
  1. Stay in this WhatsApp Web window
  2. Open chats that still show draft text
  3. Press Send yourself, one by one
  4. Re-run the same command to resume failures / remaining SRs
"""
        )
        input("Browser is open for sending. Press Enter here when you are done…")
        context.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
