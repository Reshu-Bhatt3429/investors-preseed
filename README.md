# Idea-stage investor outreach workbook — all startup sectors

Verified date: **2026-09-02**. Deadlines, application windows, and fund status change fast.
Check the source URL before you send anything.

This repository is designed for founders across healthcare, biotech, climate, deep tech, SaaS,
consumer, AI, fintech, and other startup categories. Generalist opportunities rank highly, and
healthcare receives an additional boost without excluding any other kind of company.

## What this folder holds

| File | What it is |
| --- | --- |
| `investors_preseed.xlsx` | The full workbook. Eight sheets. Start on the `START HERE` sheet. |
| `csv/` | The same sheets as UTF-8 CSV files, for scripts and mail-merge. |
| `build_workbook.py` | The script that builds the workbook, the CSV files, and this README. Rerun it any time. |
| `raw/` | The source JSON research files. The build script only reads them. |

## Sheet guide

| Sheet | Rows | What it gives you |
| --- | --- | --- |
| START HERE | 40 | The top 40 reachable targets for founders in any startup sector, with extra weight for generalist and healthcare fit. |
| Master | 366 | Every entry from all five research files in one schema. Nothing is dropped. |
| Accelerators & Funds | 52 | The accelerator and fund file with its own columns. |
| Angels (Twitter) | 62 | The angel file with its own columns. |
| VCs & Programs | 96 | Broad and sector-specific VCs and startup programs; finance is one vertical among many. |
| Healthcare Focus | 11 | Healthcare, medtech, digital health, and biotech investors and programs. |
| Open Channels & Reddit | 145 | Subreddits, angel groups, grants, directories, and open application forms. |
| Closed - Skip | 34 | Every row you must not contact, with the reason. |

## How the Status flag was derived

The script reads the `notes` text and the `stage` text of each entry. It applies three rules in order.

1. **CLOSED/PAUSED.** The text contains one of: `closed`, `paused`, `inactive`, `no longer`,
   `not active`, `stopped`, `not currently investing`, `no longer investing`, `program complete`,
   `winding down`, `defunct`, `shut down`, `discontinued`.
   The phrase `closed-source` is excluded, because it describes software and not a program.
2. **NOT IDEA STAGE.** The text contains an explicit rejection of the idea stage. Examples:
   `idea-stage: no`, `idea stage is a poor fit`, `not idea-stage`, `seed only` (but not `pre-seed only`),
   `no pre-seed`, `requires traction`, `requires revenue`, `in MRR`, `early revenue stage`,
   `repeat founders only`, `later-stage only`.
   Soft phrases such as `idea stage not explicitly stated` do **not** trigger this flag. They mean
   the research found no statement, not a refusal.
3. **OPEN.** Everything else.

The `Status reason` column quotes the exact phrase that fired the rule, so you can check every call.

The flag is a filter, not a fact. `OPEN` means no rejection phrase was found. It does not prove the
investor is active today.

## How START HERE was selected

The script scores every OPEN row and keeps the top 40. A row must pass three gates first.

1. The status flag is OPEN. A name flagged CLOSED/PAUSED or NOT IDEA STAGE in **any** source file
   is dropped from every file. The strictest source wins.
2. The stage, notes, or how-to-apply text names the idea stage or the pre-seed stage. Example words:
   `idea-stage`, `pre-seed`, `pre-product`, `pre-revenue`, `earliest stage`, `first check`.
3. A public contact route exists: an email, an application form, an apply URL, or an open Twitter DM.
   A bare company homepage does not count. For an open channel the post URL does count.

The score favors broad applicability and healthcare while keeping access and stage fit central:

| Signal | Points |
| --- | --- |
| Sector-agnostic, generalist, any-sector, or multi-sector mandate | +8 |
| Healthcare, digital health, medtech, biotech, clinical, or life sciences | +7 |
| Climate/energy, deep tech/hardware, enterprise/SaaS, or consumer/education | +3 per group |
| AI or machine learning | +2 |
| Open application form, published email, or open DMs | +2 to +3 each |
| States idea stage / pre-product / pre-PMF | +4 |
| Warm intro only | -4 |
| The notes explicitly say the opportunity does not fit | -8 |
| Pay-to-pitch, application fee, or paid intro platform | -5 |
| Eligibility is limited to a group this workbook cannot confirm | -5 |

Rules on top of the score: one row per name, one row per contact route, a maximum of 15 rows per
category, and a minimum of 4 rows per category, so all five source files appear in the list.

## Counts

### Per category

| Category | Rows |
| --- | --- |
| Open Channel | 145 |
| VC/Program | 96 |
| Angel | 62 |
| Accelerator/Fund | 52 |
| Healthcare | 11 |
| **Total** | **366** |

### Per status

| Status flag | Rows |
| --- | --- |
| OPEN | 332 |
| NOT IDEA STAGE | 15 |
| CLOSED/PAUSED | 19 |

### Status inside each category

| Category | OPEN | NOT IDEA STAGE | CLOSED/PAUSED |
| --- | --- | --- | --- |
| Open Channel | 134 | 3 | 8 |
| VC/Program | 87 | 7 | 2 |
| Angel | 57 | 1 | 4 |
| Accelerator/Fund | 44 | 4 | 4 |
| Healthcare | 10 | 0 | 1 |

## Top 15 from START HERE

| # | Name | Category | Why it fits | Best channel | Contact |
| --- | --- | --- | --- | --- | --- |
| 1 | Precursor Ventures | Accelerator/Fund | Generalist software and hardware (portfolio: fintech, healthcare, SaaS, consumer); fund; Idea-stage: yes. | application form; hello@precursorvc.com is published for general in... | hello@precursorvc.com |
| 2 | Contrary (Contrary Capital) | Accelerator/Fund | generalist: biotech, fintech, defense, robotics, aerospace, prediction markets, education, healthcare; fund; Idea-stage: yes. | application form (cold applications accepted per secondary sources)... | stanford@vp.contrary.com |
| 3 | Pear VC (PearX) | Accelerator/Fund | Sector-agnostic; AI, enterprise, biotech, consumer verticals with partner expertise; accelerator / fund; Idea-stage: yes ('Any founder starting a c... | application form | https://pear.vc/pearx |
| 4 | Seedcamp | Accelerator/Fund | Sector-agnostic software and technology (AI, fintech, healthtech, deep tech, infrastructure); fund; Idea-stage: yes (official FAQ: 'From the moment... | application form ('Pitch Us'); cold applications accepted | https://seedcamp.com/out/f7-startup-details-core |
| 5 | Surge by Peak XV | Accelerator/Fund | Sector-agnostic technology (AI, SaaS, consumer, fintech, deep tech); accelerator / fund; Idea-stage: yes (application form accepts 'Idea/Prototype'... | application form; cold outreach accepted (FAQ invites founders to f... | contactus@peakxv.com |
| 6 | Elizabeth Yin | Angel | software generalist; SaaS; B2B; fintech; digital health; AI; Co-founder and General Partner, Hustle Fund (pre-seed fund for software startups); ex-... | public_form | https://hustlefund.typeform.com/to/UGTnIt |
| 7 | Sahil Lavingia | Angel | generalist; SaaS; creator tools; future of work; crypto; Founder and CEO of Gumroad; angel investor via SHL Capital rolling fund; Pre-seed and seed... | twitter_dm | https://shl.vc/apply |
| 8 | Tim Draper | Angel | AI, robotics & automation; crypto & blockchain; healthcare & biotech; aerospace & space; unconventional ideas; Founder and Managing Partner, Draper... | public_form | hello@meetthedrapers.com |
| 9 | Cyan Banister | Angel | Portfolio: Uber, SpaceX, Notion, plus 145 investments on record. | email | hi@longjourney.vc |
| 10 | Garry Tan | Angel | generalist; AI; consumer; fintech; crypto; enterprise; deep tech; President and CEO of Y Combinator; co-founder of Initialized Capital (now board p... | public_form | garry@initialized.com |
| 11 | Ryan Hoover | Angel | generalist; consumer; B2B; fintech; AI; health; Founder of Product Hunt; investor at Weekend Fund (early-stage fund with 350+ operator LPs); Early-... | email | team@weekend.fund |
| 12 | Everywhere Ventures | VC/Program | Future of work, fintech, health, climate, supply chain, deep tech, consumer, gov/legal; VC fund (founder-LP collective); Pre-seed accepted. | Submit the Airtable 'Join our portfolio' form linked from the homep... | https://airtable.com/appeY02X0EMCQQ19G/shrjpNqWOj0SJP1ix |
| 13 | Homebrew | VC/Program | Generalist, mission-driven; portfolio heavy in fintech, AI agents, crypto tooling, infrastructure; VC firm (partners' own capital, evergreen); Pre-... | Cold email a partner directly with a short deck; Hunter Walk public... | hunter@homebrew.co |
| 14 | 1517 Fund | Accelerator/Fund | Software, hardware with data plays, deep tech/sci-fi tech, biotech; 'more about the who than the what'; fund / grant; Idea-stage: yes. $50,000 chec... | application form or cold email ('Write in! We read all submissions'... | info@1517fund.com |
| 15 | a16z Speedrun | Accelerator/Fund | Sector-agnostic (team roots in games, AI, consumer tech); accelerator / fund; Idea-stage: yes (solo founders accepted; they aim to be the first che... | application form; sr-team@a16z.com is published for questions | sr-team@a16z.com |

## Outreach playbook

Work in this order. Send the applications first, then the cold emails, then the DMs.
Log every send in a sheet with the date, so you do not contact the same person twice.

### Cold email template (6 lines)

```text
Subject: {Company} - pre-seed: {one-line outcome}

Hi {Name}, I'm building {Company}, which helps {target user} solve {specific problem}.
So far, {one concrete signal: prototype, users, pilot, waitlist, or research result}.
We're raising a {amount} pre-seed round to {one use of funds}.
Your work with {relevant sector or portfolio company} is why I'm reaching out.
Deck attached. Would a 20-minute conversation next week make sense?
```

Rules: one metric, one ask, no attachment over 10 MB, no confidential data.
Send from your own domain. Put the deck in the first message; do not ask permission to send it.

### Twitter DM template (3 lines)

```text
Hi {Name} - I'm building {Company} for {target user}: {one-line outcome}.
{one concrete signal}. Raising a {amount} pre-seed round.
Your work in {relevant sector} looks aligned. Open to a short conversation?
```

Rules: DM only where the workbook says the DMs are open. Do not send a second DM if there is no reply.
Reply to their posts for two weeks first; a cold DM after that gets a better read rate.

### r/angelinvestors post format

Taken from the `how_to_apply_or_post` field of the `r/angelinvestors` entry in the open-channels file:

```text
1) Read the pinned 'Founders if you are posting, please read this first', which links to the wiki page /r/angelinvestors/wiki/seeking-investment-requirements. The wiki (snapshot 2025-04-23) sets expectations by category: SaaS 'Demonstrate initial user adoption or revenue (e.g., 6+ months of MRR or user growth). Pre-revenue is acceptable if the market opportunity is compelling'; hardware 'at least a functional prototype'; deep tech 'Team Expertise' and milestones; consumer 'Market Validation'. Deck sections: Problem/Solution, Market Opportunity, Product/Traction, Business Model, Team, Use of Funds.
2) Use post flair: 'Seeking Investment', 'Investor Interest', 'Investor Discussion', 'Startup Update', 'Lessons Learned'.
3) Write a complete pitch; AutoModerator removes 'incomplete pitches' (a founder reported four removals on 2026-08-31 before the post stayed up). Include problem, product, traction, team, ask, and terms (example that stayed up: '$500K SAFE at $5M cap').
4) Use the AutoModerator 'Weekly Expertise & Resource Exchange' (every Wednesday) for non-monetary asks; its template: OFFERING: Expertise, Capacity, Evidence; SEEKING: 'The Bottleneck: (e.g., Need a warm intro to a Fintech-focused seed fund)', 'The Context: (e.g., Post-revenue SAAS, $20k MRR, expanding to UK)', 'The Trade'.
```

Warning: a public fundraising post can be a general solicitation. The open-channels file states:
'General solicitation on a public forum can have securities-law consequences (Reg D 506(c) in the
US)'. The r/startups template states: 'Do NOT solicit funds publicly--this may be illegal for you
to do so'. Read the rule and take legal advice before you post. This workbook is not legal advice.

## Known gaps

These fields hold the word `unknown` most often. Fill them yourself before you use the row.

| Field | `unknown` rows | Share of 366 |
| --- | --- | --- |
| Timing / deadline | 314 | 86% |
| Other channels | 304 | 83% |
| Partner Twitters | 276 | 75% |
| LinkedIn | 221 | 60% |
| Twitter | 189 | 52% |
| Sectors | 146 | 40% |
| Stage | 145 | 40% |
| Email or form | 116 | 32% |
| Check size / terms | 40 | 11% |
| Best channel | 18 | 5% |
| Outreach rule / cold pitch policy / how to apply | 12 | 3% |
| Geography | 7 | 2% |

Other gaps you must accept:

- 87 of 366 rows mention healthcare, healthtech, medtech, biotech, medical,
  clinical, or life sciences; 79 of those are OPEN.
- 74 rows explicitly describe a generalist, sector-agnostic, any-sector, or
  multi-sector mandate. Many other rows cover several named sectors without using those exact labels.
- 37 rows give a warm intro as the best channel. A warm intro is not in this workbook. You must
  find the path yourself.
- Check sizes come from third-party aggregators for many funds. Treat them as a range, not a quote.
- The open-channel rows have no `stage` field. Their status comes from the notes text only.
- 51 rows share a name with a row in another source file. The `Also in` column marks them.
  The two rows can disagree, because they were researched from different sources.

## Rebuild

```bash
python3 build_workbook.py
```

Verified date: 2026-09-02. Deadlines change. Confirm each one on the source URL before you apply.
