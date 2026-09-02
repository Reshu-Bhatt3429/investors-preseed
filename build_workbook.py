#!/usr/bin/env python3
"""Build the pre-seed investor outreach workbook, CSV files, and README.

Input : raw/*.json  (read-only, never modified)
Output: investors_preseed.xlsx, csv/*.csv, README.md

Rerunnable: it overwrites the outputs every time. It never writes into raw/.
Requires openpyxl (tested with 3.1.5).
"""

from __future__ import annotations

import csv
import json
import math
import re
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "raw"
XLSX = ROOT / "investors_preseed.xlsx"
CSV_DIR = ROOT / "csv"
README = ROOT / "README.md"

VERIFIED_DATE = "2026-09-02"
EXCEL_CELL_LIMIT = 32767
TRUNC_AT = 32000

SOURCES = [
    # (json file, category label, sheet name, csv name)
    ("accelerators_and_funds.json", "Accelerator/Fund", "Accelerators & Funds", "accelerators_and_funds.csv"),
    ("angels_twitter.json", "Angel", "Angels (Twitter)", "angels_twitter.csv"),
    ("vcs_and_programs.json", "VC/Program", "VCs & Programs", "vcs_and_programs.csv"),
    ("healthcare_investors_and_programs.json", "Healthcare", "Healthcare Focus", "healthcare_focus.csv"),
    ("open_channels_reddit_directories.json", "Open Channel", "Open Channels & Reddit", "open_channels_reddit_directories.csv"),
]

# --------------------------------------------------------------------------
# 1. Flatten helpers
# --------------------------------------------------------------------------


def flat(value) -> str:
    """Turn any JSON value into a '; '-joined string. Keep the word 'unknown'."""
    if value is None:
        return "unknown"
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, list):
        parts = [flat(v) for v in value]
        return "; ".join(p for p in parts if p)
    if isinstance(value, dict):
        return "; ".join(f"{k}: {flat(v)}" for k, v in value.items())
    return str(value)


def join(*parts: str) -> str:
    """Join non-empty, non-'unknown' parts with '; '. Return 'unknown' if none."""
    keep, seen = [], set()
    for p in parts:
        p = (p or "").strip()
        if not p or p.lower() in {"unknown", "none", "n/a", ""}:
            continue
        if p in seen:
            continue
        seen.add(p)
        keep.append(p)
    return "; ".join(keep) if keep else "unknown"


def cap(text: str) -> str:
    if len(text) > TRUNC_AT:
        return text[:TRUNC_AT] + " ... [truncated]"
    return text


# --------------------------------------------------------------------------
# 2. Status flag rules
# --------------------------------------------------------------------------

# CLOSED/PAUSED: a status word in the notes or the stage text.
# 'closed-source' is excluded because it describes software, not a program.
CLOSED_RE = re.compile(
    r"(?:\bclosed\b(?!\s*-?\s*(?:source|round|deal|beta|door))"
    r"|\bpaused\b|\binactive\b|\bno longer\b|\bnot active\b|\bstopped\b"
    r"|not currently investing|no longer investing|program complete"
    r"|winding down|\bdefunct\b|shut down|discontinued)",
    re.I,
)

# NOT IDEA STAGE: explicit phrases only. Phrases such as
# 'idea stage not explicitly stated' are NOT a rejection, so they are excluded.
NOT_IDEA_RE = re.compile(
    r"""(
      idea[-\s]stage\s*:\s*no\b
    | idea\s*stage\s+is\s+(?:a\s+)?(?:poor\s+fit|unlikely|not\s+listed)
    | treat\s+idea\s+stage\s+as\s+unlikely
    | \bnot\s+idea[-\s]stage\b
    | idea[-\s]stage[^.;]{0,40}?is\s+not\s+a\s+fit
    | idea[-\s]stage\s+(?:is\s+)?excluded
    | does\s+not\s+invest\s+at\s+(?:the\s+)?idea
    | (?<!pre-)(?<!pre\s)\bseed\s+only\b
    | \bno\s+pre-?seed\b(?!\s+(?:statement|or\b))
    | no\s+pre-?seed\s+stage\s+mentioned
    | requires?\s+(?:mrr|revenue|traction|paying\s+customers|a\s+live\s+product)
    | \bmrr\s+required\b | require\s+traction | \bin\s+MRR\b
    | early\s+revenue\s+stage
    | repeat\s+founders\s+only | second-?time\s+founders\s+only
    | post-?revenue\s+only | growth[-\s]stage\s+only | later[-\s]stage\s+only
    )""",
    re.I | re.X,
)

STATUS_OPEN = "OPEN"
STATUS_CLOSED = "CLOSED/PAUSED"
STATUS_NOT_IDEA = "NOT IDEA STAGE"


def status_flag(notes: str, stage: str) -> tuple[str, str]:
    """Return (flag, reason). Reason quotes the matched phrases."""
    text = f"{notes} || {stage}"
    hits = sorted({m.group(0).strip().lower() for m in CLOSED_RE.finditer(text)})
    if hits:
        return STATUS_CLOSED, "notes/stage contain: " + ", ".join(f"'{h}'" for h in hits)
    hits = sorted({m.group(0).strip().lower() for m in NOT_IDEA_RE.finditer(text)})
    if hits:
        return STATUS_NOT_IDEA, "notes/stage contain: " + ", ".join(f"'{h}'" for h in hits)
    return STATUS_OPEN, ""


# --------------------------------------------------------------------------
# 3. Idea-stage, contact, and broad-fit tests for START HERE
# --------------------------------------------------------------------------

IDEA_POS_RE = re.compile(
    r"(idea[-\s]stage\s*:\s*(?:yes|partial)|\bidea[-\s]stage\b|\bidea stage\b|\bpre-?seed\b"
    r"|\bpre-?product\b|\bpre-?revenue\b|\bpre-?PMF\b|no product|earliest stage"
    r"|first check|first money in|day one|too early|napkin|before product market fit)",
    re.I,
)

TIMING_RE = re.compile(
    r"[^.]*\b(?:deadline|applications? (?:are )?(?:open|close[sd]?)|apply by|before "
    r"[A-Z][a-z]{2,8} \d{1,2}|cohort|batch \d|rolling basis|rolling applications|"
    r"applications open|next cohort|closes on)\b[^.]*\.",
    re.I,
)

DM_OPEN_RE = re.compile(r"(dms?\s+(?:are\s+)?open|open\s+dms|dm\s+open|public\s+dm|twitter_dm|cold\s+dm)", re.I)
URL_RE = re.compile(r"https?://\S+")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

FOCUS_WEIGHTS = [
    # Broad mandates rank first because this workbook should work for any startup idea.
    (re.compile(r"sector[- ]agnostic|generalist|any (?:stage or )?sector|all sectors|across industries|multi-sector|horizontal", re.I), 8),
    # Healthcare is an intentional plus, spanning software, devices, and life sciences.
    (re.compile(r"health(?:care|tech)?|digital health|medical|medtech|biotech|life sciences?|therapeutic|diagnostic|clinical|wellness|bio\b", re.I), 7),
    # Reward other major startup categories evenly so no single vertical dominates.
    (re.compile(r"climate|clean ?tech|energy|sustainab|agri(?:culture|tech)?|food ?tech", re.I), 3),
    (re.compile(r"deep ?tech|hard ?tech|hardware|robotic|manufactur|space|aerospace|materials", re.I), 3),
    (re.compile(r"enterprise|SaaS|developer tools?|infrastructure|cybersecurity|future of work", re.I), 3),
    (re.compile(r"consumer|marketplace|commerce|education|edtech|creator|media|gaming", re.I), 3),
    (re.compile(r"\bAI\b|machine learning|\bML\b|LLM|artificial intelligence|agentic", re.I), 2),
]

CHANNEL_BONUS = [
    (re.compile(r"application form|apply|public_form|open application|typeform|airtable", re.I), 3),
    (EMAIL_RE, 3),
    (DM_OPEN_RE, 2),
    (re.compile(r"cold email|we read (?:all|every)|read everything|no warm intro needed|no intro needed", re.I), 3),
    (re.compile(r"warm intro only|warm intro required|intro[- ]only", re.I), -4),
]

# Push rows down when their own notes explicitly say the opportunity is a poor fit.
MISFIT_RE = re.compile(
    r"does not fit|is not a fit|not public goods|only open (?:tooling|source)"
    r"|public goods only|do not (?:spend|waste)",
    re.I,
)

# The founder must pay to pitch, or the pitch is behind a paid platform.
PAYWALL_RE = re.compile(
    r"pay[-\s]?to[-\s]?pitch|application fee|non-refundable|paid intro"
    r"|\$\d[\d,.]*\s*(?:fee|/\s*month|per month|a month)|% of the amount raised"
    r"|tuition|membership fee|subscription",
    re.I,
)

# The program only takes a founder group that this workbook cannot confirm.
ELIGIBILITY_RE = re.compile(
    r"only in women-led|women-led only|women founders only|female founders only"
    r"|students only|must be (?:a )?(?:student|enrolled)|currently enrolled"
    r"|under-?represented (?:founders )?only|HBCU only|must be under \d\d",
    re.I,
)


NEGATIVE_SENTENCE_RE = re.compile(
    r"does not|do not|no prediction|not appear|is not|are not|none found|not verified"
    r"|unverified|could not|no evidence|not stated|no public|never|avoid|skip",
    re.I,
)


def first_sentence(notes: str, keywords: re.Pattern | None = None, limit: int = 240) -> str:
    """Pick one sentence from the notes. Prefer a positive sentence that names a target sector."""
    text = " ".join(notes.split())
    if not text or text.lower() == "unknown":
        return "unknown"
    sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+(?=[A-Z0-9'\"(])", text) if s.strip()]
    if not sentences:
        sentences = [text]
    sentences = [re.sub(r"^(?:Description|Reachable investors|Typical outcome|Notes)\s*:\s*", "", s)
                 for s in sentences]
    sentences = [s for s in sentences if s]
    if not sentences:
        sentences = [text]
    chosen = sentences[0]
    if keywords is not None:
        best = None
        for s in sentences:
            if not keywords.search(s) or len(s) > limit:
                continue
            if NEGATIVE_SENTENCE_RE.search(s):
                continue              # never sell a row on a negative sentence
            best = s
            break
        if best:
            chosen = best
    if len(chosen) > limit:
        chosen = chosen[: limit - 3].rsplit(" ", 1)[0] + "..."
    return chosen


FIT_RE = re.compile(
    r"sector[- ]agnostic|generalist|any sector|all sectors|health|medical|medtech|biotech|life science"
    r"|digital health|clinical|wellness|climate|energy|sustainab|agri|food|deep ?tech|hard ?tech"
    r"|hardware|robotic|space|materials|enterprise|SaaS|developer|infrastructure|cybersecurity"
    r"|consumer|marketplace|commerce|education|edtech|creator|media|gaming|\bAI\b|machine learning",
    re.I,
)


# --------------------------------------------------------------------------
# 4. Load and normalise
# --------------------------------------------------------------------------


def load_raw() -> dict[str, list[dict]]:
    data = {}
    for fname, *_ in SOURCES:
        with (RAW / fname).open(encoding="utf-8") as fh:
            data[fname] = json.load(fh)
    return data


def native_rows(fname: str, entries: list[dict]) -> tuple[list[str], list[list[str]]]:
    """Native columns for the per-source sheet, in the order the JSON declares them."""
    if fname == "vcs_and_programs.json":
        # The source combines broad and specialized investors. Put the entries
        # most useful to a cross-sector founder at the top of the native sheet.
        entries = sorted(
            entries,
            key=lambda e: (
                -sum(
                    weight
                    for rx, weight in FOCUS_WEIGHTS
                    if rx.search(" ".join(str(e.get(k, "")) for k in ("sectors", "stage", "notes", "type")))
                ),
                str(e.get("name", "")).lower(),
            ),
        )
    headers: list[str] = []
    for e in entries:
        for k in e:
            if k not in headers:
                headers.append(k)
    pretty = [h.replace("_", " ").strip().capitalize() for h in headers]
    rows = [[cap(flat(e.get(h, "unknown"))) for h in headers] for e in entries]
    return pretty, rows


MASTER_HEADERS = [
    "Category", "Name", "Type/What they do", "Stage", "Check size / terms", "Sectors",
    "Geography", "Website/URL", "Email or form", "Twitter", "Partner Twitters", "LinkedIn",
    "Other channels", "Best channel",
    "Outreach rule / cold pitch policy / how to apply", "Timing / deadline",
    "Status flag", "Status reason", "Also in", "Notes", "Sources", "Verified date",
]


def to_master(fname: str, category: str, e: dict) -> dict:
    g = lambda k: flat(e.get(k, "unknown"))  # noqa: E731
    if category == "Accelerator/Fund":
        row = dict(
            type=g("type"), stage=g("stage"), check=g("check_size_terms"), sectors=g("sectors"),
            geo=g("geography"), url=g("website"), email=join(g("email"), g("application_url")),
            twitter=g("twitter"), partners=g("partner_twitters"), linkedin=g("linkedin"),
            other="unknown", best=g("best_channel"), rule=g("application_url"),
            timing=g("application_timing"), notes=g("notes"),
        )
    elif category == "Angel":
        row = dict(
            type=g("what_they_do"), stage=g("stage"), check=g("check_size"), sectors=g("sectors"),
            geo=g("geography"), url=join(g("linkedin")), email=g("email_or_form"),
            twitter=g("twitter"), partners="unknown", linkedin=g("linkedin"),
            other=g("other_channels"), best=g("best_channel"), rule=g("outreach_rule"),
            timing="unknown", notes=g("notes"),
        )
    elif category in {"VC/Program", "Healthcare"}:
        row = dict(
            type=g("type"), stage=g("stage"), check=g("check_size"), sectors=g("sectors"),
            geo=g("geography"), url=g("website"), email=g("pitch_email_or_form"),
            twitter=g("twitter"), partners=g("partner_twitters"), linkedin=g("linkedin"),
            other="unknown", best=g("best_channel"), rule=g("cold_pitch_policy"),
            timing="unknown", notes=g("notes"),
        )
    else:  # Open Channel
        notes = " | ".join(
            f"{label}: {val}"
            for label, val in [
                ("Description", g("description")),
                ("Reachable investors", g("reachable_investors")),
                ("Typical outcome", g("typical_outcome")),
                ("Notes", g("notes")),
            ]
            if val and val.lower() != "unknown"
        ) or "unknown"
        contacts = g("contacts")
        row = dict(
            type=g("channel_type"), stage="unknown", check=g("cost"), sectors="unknown",
            geo=g("geography"), url=g("url"), email=contacts,
            twitter="unknown", partners="unknown", linkedin="unknown",
            other="unknown", best=g("channel_type"), rule=g("how_to_apply_or_post"),
            timing="unknown", notes=notes,
        )

    # The status flag reads the raw `notes` and `stage` text of the entry, not
    # the merged description block, so that words such as "a closed round" in a
    # description do not trigger a false CLOSED flag.
    flag, reason = status_flag(g("notes"), row["stage"])
    return {
        "Category": category,
        "Name": g("name"),
        "Type/What they do": row["type"],
        "Stage": row["stage"],
        "Check size / terms": row["check"],
        "Sectors": row["sectors"],
        "Geography": row["geo"],
        "Website/URL": row["url"],
        "Email or form": row["email"],
        "Twitter": row["twitter"],
        "Partner Twitters": row["partners"],
        "LinkedIn": row["linkedin"],
        "Other channels": row["other"],
        "Best channel": row["best"],
        "Outreach rule / cold pitch policy / how to apply": row["rule"],
        "Timing / deadline": row["timing"],
        "Status flag": flag,
        "Status reason": reason,
        "Also in": "",  # filled later
        "Notes": row["notes"],
        "Sources": g("sources"),
        "Verified date": g("verified_date") or VERIFIED_DATE,
        "_file": fname,
    }


# --------------------------------------------------------------------------
# 5. START HERE selection
# --------------------------------------------------------------------------


def norm_name(name: str) -> str:
    """Loose name key: lower case, drop a trailing bracket, drop punctuation."""
    n = re.sub(r"\([^)]*\)", " ", name.lower())
    n = re.sub(r"[^a-z0-9]+", " ", n)
    return " ".join(n.split())


def contact_of(row: dict) -> str:
    """Best reachable contact, or '' when no public email, form, apply URL, or open DM exists.

    A bare company homepage does not count as a contact route. For an open
    channel the `url` IS the place where you post or apply, so it does count.
    """
    fields = ["Email or form", "Outreach rule / cold pitch policy / how to apply", "Best channel"]
    if row.get("Category") == "Open Channel":
        fields.append("Website/URL")
    for field in fields:
        v = row.get(field, "")
        if not v or v.lower() == "unknown":
            continue
        m = EMAIL_RE.search(v)
        if m:
            return m.group(0)
        m = URL_RE.search(v)
        if m:
            return m.group(0).rstrip(".,;)")
    dm_blob = " ".join(str(row.get(f, "")) for f in
                       ("Best channel", "Outreach rule / cold pitch policy / how to apply", "Notes"))
    if DM_OPEN_RE.search(dm_blob):
        tw = row.get("Twitter", "")
        return f"Twitter DM {tw}" if tw and tw.lower() != "unknown" else "Twitter DM (DMs open)"
    return ""


def fit_score(row: dict) -> int:
    blob = " ".join(str(row.get(f, "")) for f in
                    ("Sectors", "Type/What they do", "Notes", "Name", "Stage"))
    score = sum(w for rx, w in FOCUS_WEIGHTS if rx.search(blob))
    chan = " ".join(str(row.get(f, "")) for f in
                    ("Best channel", "Email or form", "Outreach rule / cold pitch policy / how to apply"))
    score += sum(w for rx, w in CHANNEL_BONUS if rx.search(chan))
    stage_blob = f"{row.get('Stage','')} {row.get('Notes','')}"
    if re.search(r"idea[-\s]stage\s*:\s*yes|no idea is too early|pre-?product|pre-?PMF|before product market fit", stage_blob, re.I):
        score += 4
    elif re.search(r"\bpre-?seed\b", stage_blob, re.I):
        score += 2
    if MISFIT_RE.search(str(row.get("Notes", ""))):
        score -= 8
    money_blob = " ".join(str(row.get(f, "")) for f in
                          ("Check size / terms", "Notes",
                           "Outreach rule / cold pitch policy / how to apply"))
    if PAYWALL_RE.search(money_blob):
        score -= 5
    if ELIGIBILITY_RE.search(money_blob):
        score -= 5
    if row.get("Category") == "Open Channel":
        chan_blob = " ".join(str(row.get(f, "")) for f in
                             ("Check size / terms", "Outreach rule / cold pitch policy / how to apply", "Notes"))
        if re.search(r"\bfree\b", chan_blob, re.I):
            score += 1
        if re.search(r"open pitching|pitching is the norm|open application|open to pitch"
                     r"|rolling (?:applications|basis)|apply any ?time|anyone can apply"
                     r"|no warm intro|cold (?:pitch|outreach) (?:is )?(?:allowed|welcome|ok)",
                     chan_blob, re.I):
            score += 4
    return score


def build_start_here(master: list[dict], limit: int = 40,
                     per_cat_cap: int = 15, per_cat_min: int = 4) -> list[dict]:
    # A name flagged CLOSED/PAUSED or NOT IDEA STAGE in ANY source file is
    # dropped everywhere. The strictest source wins.
    bad_names = {r["Name"].strip().lower() for r in master if r["Status flag"] != STATUS_OPEN}

    pool = []
    for r in master:
        if r["Status flag"] != STATUS_OPEN or r["Name"].strip().lower() in bad_names:
            continue
        idea_blob = f"{r['Stage']} {r['Notes']} {r['Outreach rule / cold pitch policy / how to apply']}"
        if not IDEA_POS_RE.search(idea_blob):
            continue
        contact = contact_of(r)
        if not contact:
            continue
        pool.append((fit_score(r), r, contact))

    pool.sort(key=lambda t: (-t[0], t[1]["Name"].lower()))

    # Keep one row per name, and one row per contact route: the highest-scoring one.
    seen_names, seen_contacts, deduped = set(), set(), []
    for item in pool:
        name_key = norm_name(item[1]["Name"])
        contact_key = item[2].strip().lower().rstrip("/")
        if name_key in seen_names or contact_key in seen_contacts:
            continue
        seen_names.add(name_key)
        seen_contacts.add(contact_key)
        deduped.append(item)
    pool = deduped

    picked, counts = [], Counter()
    chosen: set[int] = set()

    # Pass 1: guarantee every category is represented.
    for cat in ("Accelerator/Fund", "Angel", "VC/Program", "Healthcare", "Open Channel"):
        for item in pool:
            if counts[cat] >= per_cat_min:
                break
            if item[1]["Category"] != cat or id(item[1]) in chosen:
                continue
            counts[cat] += 1
            chosen.add(id(item[1]))
            picked.append(item)

    # Pass 2: fill by score, with a per-category cap.
    for item in pool:
        if len(picked) >= limit:
            break
        if id(item[1]) in chosen or counts[item[1]["Category"]] >= per_cat_cap:
            continue
        counts[item[1]["Category"]] += 1
        chosen.add(id(item[1]))
        picked.append(item)

    # Pass 3: top up if the caps left the list short.
    for item in pool:
        if len(picked) >= limit:
            break
        if id(item[1]) not in chosen:
            chosen.add(id(item[1]))
            picked.append(item)

    picked = picked[:limit]
    picked.sort(key=lambda t: (-t[0], t[1]["Name"].lower()))

    # Timing: use the row's own field, else the same company's row in another
    # source file, else the first sentence in the notes that names a deadline.
    timing_by_name: dict[str, str] = {}
    for r in master:
        t = str(r.get("Timing / deadline", "")).strip()
        if t and t.lower() != "unknown":
            timing_by_name.setdefault(norm_name(r["Name"]), t)

    out = []
    for i, (score, r, contact) in enumerate(picked, start=1):
        src = URL_RE.search(r["Sources"])
        timing = str(r.get("Timing / deadline", "")).strip()
        if not timing or timing.lower() == "unknown":
            timing = timing_by_name.get(norm_name(r["Name"]), "")
        if not timing:
            blob = " ".join([str(r.get("Best channel", "")), str(r.get("Notes", "")),
                             str(r.get("Outreach rule / cold pitch policy / how to apply", ""))])
            blob = URL_RE.sub(" ", " ".join(blob.split()))   # a URL dot is not a full stop
            m = TIMING_RE.search(blob)
            timing = " ".join(m.group(0).split())[:240] if m else "unknown"
        out.append({
            "Rank": i,
            "Name": r["Name"],
            "Category": r["Category"],
            "Why it fits (sector/stage evidence)": first_sentence(
                join(r["Sectors"], r["Type/What they do"], r["Stage"], r["Notes"]), FIT_RE
            ),
            "Best channel": r["Best channel"],
            "Contact (email/form/URL)": contact,
            "Twitter": r["Twitter"],
            "Check size": r["Check size / terms"],
            "Deadline/timing if any": timing,
            "Source URL": src.group(0).rstrip(".,;)") if src else "unknown",
            "_score": score,
            "_status": r["Status flag"],
        })
    return out


START_HEADERS = ["Rank", "Name", "Category", "Why it fits (sector/stage evidence)", "Best channel",
                 "Contact (email/form/URL)", "Twitter", "Check size", "Deadline/timing if any", "Source URL"]


# --------------------------------------------------------------------------
# 6. Excel writing
# --------------------------------------------------------------------------

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
LINK_FONT = Font(color="0563C1", underline="single")
FILL_OPEN = PatternFill("solid", fgColor="C6EFCE")
FILL_CLOSED = PatternFill("solid", fgColor="D9D9D9")
FILL_NOT_IDEA = PatternFill("solid", fgColor="FFEB9C")
STATUS_FILLS = {STATUS_OPEN: FILL_OPEN, STATUS_CLOSED: FILL_CLOSED, STATUS_NOT_IDEA: FILL_NOT_IDEA}

WRAP_COLS = {"notes", "why it fits (sector/stage evidence)", "status reason", "description",
             "outreach rule / cold pitch policy / how to apply", "outreach rule", "how to apply or post",
             "cold pitch policy", "reachable investors", "typical outcome", "stage",
             "check size / terms", "sources", "best channel", "notes ", "what they do",
             "type/what they do", "application timing", "timing / deadline"}


def write_sheet(wb, title: str, headers: list[str], rows: list[list[str]],
                status_col: str | None = None) -> None:
    ws = wb.create_sheet(title=title)
    ws.append(headers)
    for c in ws[1]:
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for row in rows:
        ws.append([cap(str(v)) for v in row])

    lower = [h.strip().lower() for h in headers]
    status_idx = lower.index(status_col.lower()) if status_col and status_col.lower() in lower else None

    widths = {}
    for j, h in enumerate(headers, start=1):
        lengths = sorted([len(str(h))] + [len(str(row[j - 1])) if j - 1 < len(row) else 0
                                          for row in rows])
        p90 = lengths[min(len(lengths) - 1, int(len(lengths) * 0.9))]
        longest = lengths[-1]
        # Wrap a named notes column, or any column that holds long free text.
        wrap = lower[j - 1] in WRAP_COLS or longest > 80

        for i, row in enumerate(rows, start=2):
            val = str(row[j - 1]) if j - 1 < len(row) else ""
            cell = ws.cell(row=i, column=j)
            cell.alignment = Alignment(wrap_text=wrap, vertical="top")
            if val.startswith("http") and " " not in val and len(val) < 250:
                cell.hyperlink = val
                cell.font = LINK_FONT
            if status_idx is not None:
                fill = STATUS_FILLS.get(str(row[status_idx]))
                if fill is not None:
                    cell.fill = fill

        # Size on the 90th percentile, not the longest value, so one long cell
        # does not stretch the whole column. Wrapped columns stay narrower.
        if wrap:
            widths[j] = max(22, min(45, p90 + 2))
        else:
            widths[j] = max(12, min(60, p90 + 2))
    for j, w in widths.items():
        ws.column_dimensions[get_column_letter(j)].width = w

    # Keep wrapped records readable without allowing a long research note to
    # create an unusably tall row. Five visible lines is a practical scan view;
    # the full value remains available in the cell/formula bar.
    ws.row_dimensions[1].height = 32
    for i, row in enumerate(rows, start=2):
        visible_lines = 1
        for j, value in enumerate(row, start=1):
            text = str(value)
            width = max(8, widths.get(j, 12) - 2)
            line_count = sum(max(1, math.ceil(len(line) / width)) for line in text.splitlines() or [""])
            visible_lines = max(visible_lines, min(5, line_count))
        ws.row_dimensions[i].height = 15 * visible_lines

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, len(rows) + 1)}"


def write_csv(path: Path, headers: list[str], rows: list[list[str]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(headers)
        w.writerows(rows)


# --------------------------------------------------------------------------
# 7. README
# --------------------------------------------------------------------------

COLD_EMAIL = """Subject: {Company} - pre-seed: {one-line outcome}

Hi {Name}, I'm building {Company}, which helps {target user} solve {specific problem}.
So far, {one concrete signal: prototype, users, pilot, waitlist, or research result}.
We're raising a {amount} pre-seed round to {one use of funds}.
Your work with {relevant sector or portfolio company} is why I'm reaching out.
Deck attached. Would a 20-minute conversation next week make sense?
"""

TWITTER_DM = """Hi {Name} - I'm building {Company} for {target user}: {one-line outcome}.
{one concrete signal}. Raising a {amount} pre-seed round.
Your work in {relevant sector} looks aligned. Open to a short conversation?
"""


def build_readme(master, start, per_cat, per_status, gaps, sheet_counts, reddit_format, facts) -> str:
    lines = []
    a = lines.append
    a("# Idea-stage investor outreach workbook — all startup sectors")
    a("")
    a(f"Verified date: **{VERIFIED_DATE}**. Deadlines, application windows, and fund status change fast.")
    a("Check the source URL before you send anything.")
    a("")
    a("This repository is designed for founders across healthcare, biotech, climate, deep tech, SaaS,")
    a("consumer, AI, fintech, and other startup categories. Generalist opportunities rank highly, and")
    a("healthcare receives an additional boost without excluding any other kind of company.")
    a("")
    a("## What this folder holds")
    a("")
    a("| File | What it is |")
    a("| --- | --- |")
    a("| `investors_preseed.xlsx` | The full workbook. Eight sheets. Start on the `START HERE` sheet. |")
    a("| `csv/` | The same sheets as UTF-8 CSV files, for scripts and mail-merge. |")
    a("| `build_workbook.py` | The script that builds the workbook, the CSV files, and this README. Rerun it any time. |")
    a("| `raw/` | The source JSON research files. The build script only reads them. |")
    a("")
    a("## Sheet guide")
    a("")
    a("| Sheet | Rows | What it gives you |")
    a("| --- | --- | --- |")
    a(f"| START HERE | {sheet_counts['START HERE']} | The top 40 reachable targets for founders in any startup sector, with extra weight for generalist and healthcare fit. |")
    a(f"| Master | {sheet_counts['Master']} | Every entry from all five research files in one schema. Nothing is dropped. |")
    a(f"| Accelerators & Funds | {sheet_counts['Accelerators & Funds']} | The accelerator and fund file with its own columns. |")
    a(f"| Angels (Twitter) | {sheet_counts['Angels (Twitter)']} | The angel file with its own columns. |")
    a(f"| VCs & Programs | {sheet_counts['VCs & Programs']} | Broad and sector-specific VCs and startup programs; finance is one vertical among many. |")
    a(f"| Healthcare Focus | {sheet_counts['Healthcare Focus']} | Healthcare, medtech, digital health, and biotech investors and programs. |")
    a(f"| Open Channels & Reddit | {sheet_counts['Open Channels & Reddit']} | Subreddits, angel groups, grants, directories, and open application forms. |")
    a(f"| Closed - Skip | {sheet_counts['Closed - Skip']} | Every row you must not contact, with the reason. |")
    a("")
    a("## How the Status flag was derived")
    a("")
    a("The script reads the `notes` text and the `stage` text of each entry. It applies three rules in order.")
    a("")
    a("1. **CLOSED/PAUSED.** The text contains one of: `closed`, `paused`, `inactive`, `no longer`,")
    a("   `not active`, `stopped`, `not currently investing`, `no longer investing`, `program complete`,")
    a("   `winding down`, `defunct`, `shut down`, `discontinued`.")
    a("   The phrase `closed-source` is excluded, because it describes software and not a program.")
    a("2. **NOT IDEA STAGE.** The text contains an explicit rejection of the idea stage. Examples:")
    a("   `idea-stage: no`, `idea stage is a poor fit`, `not idea-stage`, `seed only` (but not `pre-seed only`),")
    a("   `no pre-seed`, `requires traction`, `requires revenue`, `in MRR`, `early revenue stage`,")
    a("   `repeat founders only`, `later-stage only`.")
    a("   Soft phrases such as `idea stage not explicitly stated` do **not** trigger this flag. They mean")
    a("   the research found no statement, not a refusal.")
    a("3. **OPEN.** Everything else.")
    a("")
    a("The `Status reason` column quotes the exact phrase that fired the rule, so you can check every call.")
    a("")
    a("The flag is a filter, not a fact. `OPEN` means no rejection phrase was found. It does not prove the")
    a("investor is active today.")
    a("")
    a("## How START HERE was selected")
    a("")
    a("The script scores every OPEN row and keeps the top 40. A row must pass three gates first.")
    a("")
    a("1. The status flag is OPEN. A name flagged CLOSED/PAUSED or NOT IDEA STAGE in **any** source file")
    a("   is dropped from every file. The strictest source wins.")
    a("2. The stage, notes, or how-to-apply text names the idea stage or the pre-seed stage. Example words:")
    a("   `idea-stage`, `pre-seed`, `pre-product`, `pre-revenue`, `earliest stage`, `first check`.")
    a("3. A public contact route exists: an email, an application form, an apply URL, or an open Twitter DM.")
    a("   A bare company homepage does not count. For an open channel the post URL does count.")
    a("")
    a("The score favors broad applicability and healthcare while keeping access and stage fit central:")
    a("")
    a("| Signal | Points |")
    a("| --- | --- |")
    a("| Sector-agnostic, generalist, any-sector, or multi-sector mandate | +8 |")
    a("| Healthcare, digital health, medtech, biotech, clinical, or life sciences | +7 |")
    a("| Climate/energy, deep tech/hardware, enterprise/SaaS, or consumer/education | +3 per group |")
    a("| AI or machine learning | +2 |")
    a("| Open application form, published email, or open DMs | +2 to +3 each |")
    a("| States idea stage / pre-product / pre-PMF | +4 |")
    a("| Warm intro only | -4 |")
    a("| The notes explicitly say the opportunity does not fit | -8 |")
    a("| Pay-to-pitch, application fee, or paid intro platform | -5 |")
    a("| Eligibility is limited to a group this workbook cannot confirm | -5 |")
    a("")
    a("Rules on top of the score: one row per name, one row per contact route, a maximum of 15 rows per")
    a("category, and a minimum of 4 rows per category, so all five source files appear in the list.")
    a("")
    a("## Counts")
    a("")
    a("### Per category")
    a("")
    a("| Category | Rows |")
    a("| --- | --- |")
    for k, v in per_cat:
        a(f"| {k} | {v} |")
    a(f"| **Total** | **{len(master)}** |")
    a("")
    a("### Per status")
    a("")
    a("| Status flag | Rows |")
    a("| --- | --- |")
    for k, v in per_status:
        a(f"| {k} | {v} |")
    a("")
    a("### Status inside each category")
    a("")
    a("| Category | OPEN | NOT IDEA STAGE | CLOSED/PAUSED |")
    a("| --- | --- | --- | --- |")
    grid = defaultdict(Counter)
    for r in master:
        grid[r["Category"]][r["Status flag"]] += 1
    for k, _ in per_cat:
        c = grid[k]
        a(f"| {k} | {c[STATUS_OPEN]} | {c[STATUS_NOT_IDEA]} | {c[STATUS_CLOSED]} |")
    a("")
    a("## Top 15 from START HERE")
    a("")
    a("| # | Name | Category | Why it fits | Best channel | Contact |")
    a("| --- | --- | --- | --- | --- | --- |")
    for r in start[:15]:
        def cell(x, n=150):
            x = " ".join(str(x).split()).replace("|", "/")
            return x[: n - 3] + "..." if len(x) > n else x
        a(f"| {r['Rank']} | {cell(r['Name'], 45)} | {r['Category']} | {cell(r['Why it fits (sector/stage evidence)'])} "
          f"| {cell(r['Best channel'], 70)} | {cell(r['Contact (email/form/URL)'], 70)} |")
    a("")
    a("## Outreach playbook")
    a("")
    a("Work in this order. Send the applications first, then the cold emails, then the DMs.")
    a("Log every send in a sheet with the date, so you do not contact the same person twice.")
    a("")
    a("### Cold email template (6 lines)")
    a("")
    a("```text")
    a(COLD_EMAIL.strip())
    a("```")
    a("")
    a("Rules: one metric, one ask, no attachment over 10 MB, no confidential data.")
    a("Send from your own domain. Put the deck in the first message; do not ask permission to send it.")
    a("")
    a("### Twitter DM template (3 lines)")
    a("")
    a("```text")
    a(TWITTER_DM.strip())
    a("```")
    a("")
    a("Rules: DM only where the workbook says the DMs are open. Do not send a second DM if there is no reply.")
    a("Reply to their posts for two weeks first; a cold DM after that gets a better read rate.")
    a("")
    a("### r/angelinvestors post format")
    a("")
    if reddit_format:
        a("Taken from the `how_to_apply_or_post` field of the `r/angelinvestors` entry in the open-channels file:")
        a("")
        a("```text")
        for line in reddit_format:
            a(line)
        a("```")
    else:
        a("The open-channels file holds no r/angelinvestors entry, so no post format is available.")
    a("")
    a("Warning: a public fundraising post can be a general solicitation. The open-channels file states:")
    a("'General solicitation on a public forum can have securities-law consequences (Reg D 506(c) in the")
    a("US)'. The r/startups template states: 'Do NOT solicit funds publicly--this may be illegal for you")
    a("to do so'. Read the rule and take legal advice before you post. This workbook is not legal advice.")
    a("")
    a("## Known gaps")
    a("")
    a("These fields hold the word `unknown` most often. Fill them yourself before you use the row.")
    a("")
    a(f"| Field | `unknown` rows | Share of {len(master)} |")
    a("| --- | --- | --- |")
    for field, n in gaps:
        a(f"| {field} | {n} | {n / len(master) * 100:.0f}% |")
    a("")
    a("Other gaps you must accept:")
    a("")
    a(f"- {facts['health_rows']} of {len(master)} rows mention healthcare, healthtech, medtech, biotech, medical,")
    a(f"  clinical, or life sciences; {facts['health_open']} of those are OPEN.")
    a(f"- {facts['generalist_rows']} rows explicitly describe a generalist, sector-agnostic, any-sector, or")
    a("  multi-sector mandate. Many other rows cover several named sectors without using those exact labels.")
    a(f"- {facts['warm']} rows give a warm intro as the best channel. A warm intro is not in this workbook. You must")
    a("  find the path yourself.")
    a("- Check sizes come from third-party aggregators for many funds. Treat them as a range, not a quote.")
    a("- The open-channel rows have no `stage` field. Their status comes from the notes text only.")
    a(f"- {facts['also_in']} rows share a name with a row in another source file. The `Also in` column marks them.")
    a("  The two rows can disagree, because they were researched from different sources.")
    a("")
    a("## Rebuild")
    a("")
    a("```bash")
    a("python3 build_workbook.py")
    a("```")
    a("")
    a(f"Verified date: {VERIFIED_DATE}. Deadlines change. Confirm each one on the source URL before you apply.")
    a("")
    return "\n".join(lines)


def reddit_post_format(raw: dict) -> list[str]:
    """Pull the r/angelinvestors posting steps out of the open-channels file."""
    for e in raw["open_channels_reddit_directories.json"]:
        if e.get("name", "").strip().lower() == "r/angelinvestors":
            text = flat(e.get("how_to_apply_or_post", ""))
            steps = re.split(r"(?=\b\d\)\s)", text)
            out = []
            for s in steps:
                s = " ".join(s.split())
                if s:
                    out.append(s)
            return out
    return []


# --------------------------------------------------------------------------
# 8. Main
# --------------------------------------------------------------------------


def main() -> None:
    CSV_DIR.mkdir(parents=True, exist_ok=True)
    raw = load_raw()

    master: list[dict] = []
    for fname, category, _sheet, _csvname in SOURCES:
        for e in raw[fname]:
            master.append(to_master(fname, category, e))

    # 'Also in': same exact name present in more than one source file.
    by_name = defaultdict(list)
    for r in master:
        by_name[r["Name"].strip().lower()].append(r)
    for _name, rows in by_name.items():
        if len(rows) > 1:
            for r in rows:
                others = sorted({o["Category"] for o in rows if o is not r})
                r["Also in"] = "duplicate name also in: " + ", ".join(others)

    start = build_start_here(master)

    wb = Workbook()
    wb.remove(wb.active)

    # START HERE
    start_rows = [[r[h] for h in START_HEADERS] for r in start]
    write_sheet(wb, "START HERE", START_HEADERS, start_rows)
    # every START HERE row is OPEN, so paint the whole block green
    ws = wb["START HERE"]
    for row in ws.iter_rows(min_row=2, max_row=len(start_rows) + 1):
        for c in row:
            c.fill = FILL_OPEN

    # Master
    master_rows = [[r[h] for h in MASTER_HEADERS] for r in master]
    write_sheet(wb, "Master", MASTER_HEADERS, master_rows, status_col="Status flag")

    # Native sheets
    native = {}
    for fname, _cat, sheet, csvname in SOURCES:
        headers, rows = native_rows(fname, raw[fname])
        write_sheet(wb, sheet, headers, rows)
        native[sheet] = (headers, rows, csvname)

    # Closed - Skip
    skip_headers = ["Category", "Name", "Status flag", "Reason", "Stage", "Notes", "Website/URL", "Sources"]
    skip_rows = [[r["Category"], r["Name"], r["Status flag"], r["Status reason"], r["Stage"],
                  r["Notes"], r["Website/URL"], r["Sources"]]
                 for r in master if r["Status flag"] != STATUS_OPEN]
    write_sheet(wb, "Closed - Skip", skip_headers, skip_rows, status_col="Status flag")

    wb.save(XLSX)

    # CSV files
    write_csv(CSV_DIR / "master.csv", MASTER_HEADERS, master_rows)
    write_csv(CSV_DIR / "start_here.csv", START_HEADERS, start_rows)
    for sheet, (headers, rows, csvname) in native.items():
        write_csv(CSV_DIR / csvname, headers, rows)
    write_csv(CSV_DIR / "closed_skip.csv", skip_headers, skip_rows)

    # README
    per_cat = [(c, n) for c, n in Counter(r["Category"] for r in master).most_common()]
    order = [STATUS_OPEN, STATUS_NOT_IDEA, STATUS_CLOSED]
    sc = Counter(r["Status flag"] for r in master)
    per_status = [(k, sc[k]) for k in order]

    gap_fields = [h for h in MASTER_HEADERS if h not in ("Category", "Name", "Status flag", "Status reason", "Also in", "Verified date")]
    gaps = sorted(
        ((f, sum(1 for r in master if str(r[f]).strip().lower() == "unknown")) for f in gap_fields),
        key=lambda t: -t[1],
    )
    gaps = [g for g in gaps if g[1] > 0][:12]

    health_rx = re.compile(
        r"health(?:care|tech)?|digital health|medical|medtech|biotech|life sciences?|clinical|wellness|bio\b",
        re.I,
    )
    health_rows = [r for r in master if health_rx.search(" ".join(str(v) for v in r.values()))]
    generalist_rx = re.compile(
        r"generalist|sector[- ]agnostic|any (?:stage or )?sector|all sectors|multi-sector|horizontal",
        re.I,
    )
    facts = {
        "health_rows": len(health_rows),
        "health_open": sum(1 for r in health_rows if r["Status flag"] == STATUS_OPEN),
        "generalist_rows": sum(
            1 for r in master if generalist_rx.search(" ".join(str(v) for v in r.values()))
        ),
        "warm": sum(1 for r in master if re.search(r"warm[_ ]intro", str(r["Best channel"]), re.I)),
        "also_in": sum(1 for r in master if r["Also in"]),
    }

    sheet_counts = {
        "START HERE": len(start_rows),
        "Master": len(master_rows),
        "Closed - Skip": len(skip_rows),
    }
    for sheet, (_h, rows, _c) in native.items():
        sheet_counts[sheet] = len(rows)

    README.write_text(
        build_readme(master, start, per_cat, per_status, gaps, sheet_counts,
                     reddit_post_format(raw), facts),
        encoding="utf-8",
    )

    print(f"wrote {XLSX}")
    print(f"wrote {README}")
    for p in sorted(CSV_DIR.glob("*.csv")):
        print(f"wrote {p}")
    print()
    print("sheet row counts (data rows, header excluded):")
    for k, v in sheet_counts.items():
        print(f"  {k}: {v}")
    print()
    print("status counts:", dict(per_status))
    print("category counts:", dict(per_cat))


if __name__ == "__main__":
    main()
