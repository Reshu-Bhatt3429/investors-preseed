# Pre-seed investor outreach workbook

Verified date: **2026-09-02**. Deadlines, application windows, and fund status change fast.
Check the source URL before you send anything.

## What this folder holds

| File | What it is |
| --- | --- |
| `investors_preseed.xlsx` | The full workbook. Seven sheets. Start on the `START HERE` sheet. |
| `csv/` | The same sheets as UTF-8 CSV files, for scripts and mail-merge. |
| `build_workbook.py` | The script that builds the workbook, the CSV files, and this README. Rerun it any time. |
| `raw/` | The source JSON research files. The build script only reads them. |

## Sheet guide

| Sheet | Rows | What it gives you |
| --- | --- | --- |
| START HERE | 40 | The top 40 targets for an idea-stage founder in trading bots, prediction markets, quant, and AI agents. Each row has a live contact route. |
| Master | 355 | Every entry from all four research files in one schema. Nothing is dropped. |
| Accelerators & Funds | 52 | The accelerator and fund file with its own columns. |
| Angels (Twitter) | 62 | The angel file with its own columns. |
| Fintech-Crypto-AI VCs | 96 | The sector VC file with its own columns. |
| Open Channels & Reddit | 145 | Subreddits, angel groups, grants, directories, and open application forms. |
| Closed - Skip | 33 | Every row you must not contact, with the reason. |

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

The score adds points for sector fit and subtracts points for friction:

| Signal | Points |
| --- | --- |
| Text names a prediction market, Polymarket, or Kalshi | +10 |
| Text names trading, quant, market making, or capital markets | +6 |
| Text names AI agents or agentic products | +5 |
| Text names fintech, or crypto / web3 / DeFi | +4 each |
| Text names AI or machine learning | +3 |
| Open application form, published email, or open DMs | +2 to +3 each |
| States idea stage / pre-product / pre-PMF | +4 |
| Warm intro only | -4 |
| The notes say the use case does not fit (public-goods grants) | -8 |
| Pay-to-pitch, application fee, or paid intro platform | -5 |
| Eligibility is limited to a group this workbook cannot confirm | -5 |

Rules on top of the score: one row per name, one row per contact route, a maximum of 15 rows per
category, and a minimum of 4 rows per category, so all four source files appear in the list.

## Counts

### Per category

| Category | Rows |
| --- | --- |
| Open Channel | 145 |
| Sector VC | 96 |
| Angel | 62 |
| Accelerator/Fund | 52 |
| **Total** | **355** |

### Per status

| Status flag | Rows |
| --- | --- |
| OPEN | 322 |
| NOT IDEA STAGE | 15 |
| CLOSED/PAUSED | 18 |

### Status inside each category

| Category | OPEN | NOT IDEA STAGE | CLOSED/PAUSED |
| --- | --- | --- | --- |
| Open Channel | 134 | 3 | 8 |
| Sector VC | 87 | 7 | 2 |
| Angel | 57 | 1 | 4 |
| Accelerator/Fund | 44 | 4 | 4 |

## Top 15 from START HERE

| # | Name | Category | Why it fits | Best channel | Contact |
| --- | --- | --- | --- | --- | --- |
| 1 | Y Combinator | Sector VC | Directly relevant YC fintech companies: Arbital (perps, memes, stocks, market-making), Spectre Intelligence ('We create AI traders'), Prodigy Resea... | Apply online at apply.ycombinator.com; on-time deadline shown as 'N... | https://apply.ycombinator.com/home |
| 2 | Base Ecosystem Fund / Base Batches / Base ... | Sector VC | Strong fit: explicit 'Trading' and 'AI Agents' portfolio categories; Batch 003 accepted teams include Onsight (social prediction markets), Opal (pr... | Apply to Base Batches 004 before Sep 9, 2026 (written plus video su... | https://www.base.org/ecosystem-fund/apply |
| 3 | Variant Fund | Sector VC | Relevant portfolio (official page): Melee (Pre-Seed 2025, onchain markets, alpha.melee.markets), MetaDAO (token 2025), True Markets (Series A 2024)... | Use the contact form at https://variant.fund/contact/ or email info... | https://variant.fund/contact/ |
| 4 | Contrary | Sector VC | Contrary Research published a deep dive 'Are prediction markets doomed to fail' (https://research.contrary.com/deep-dive/are-prediction-markets-doo... | Use the Applications link (applications.contrary.com) from the site... | https://applications.contrary.com |
| 5 | Paradigm | Sector VC | Strong fit signals for a trading / prediction-market builder: portfolio page lists Kalshi ('Prediction market platform'), Hyperliquid ('Financial t... | Email info@paradigm.xyz or the city inbox (sf@ / nyc@) with a short... | info@paradigm.xyz |
| 6 | Homebrew | Sector VC | Strong fit for AI-agent + fintech pitch. | Cold email a partner directly with a short deck; Hunter Walk public... | hunter@homebrew.co |
| 7 | a16z crypto (incl. CSX Crypto Startup Acce... | Sector VC | CSX page (https://a16zcrypto.com/csx/ and /accelerator/) describes curriculum (product-market fit, business models, tokenization, go-to-market, fun... | Watch https://a16zcrypto.com/csx/ and @a16zcrypto for the next CSX ... | https://a16zcrypto.com/csx/ |
| 8 | Balaji Srinivasan | Angel | Investments: Alchemy, Benchling, Digital Ocean, OpenSea, Perplexity, Polymarket, Replit, Starkware, Varda, Zora, plus protocols. | twitter_dm | Twitter DM @balajis |
| 9 | Outlier Ventures | Sector VC | The 'Crypto x Agentic AI' track fits AI-agent trading bots. | Submit the online application form; the Base Camp page says it is a... | https://outlierventures.io/apply/form/ |
| 10 | Tusk Venture Partners | Sector VC | Regulated-market thesis fits prediction markets (Polymarket/CFTC issues) very well. | Email info@tusk.vc with a deck that addresses regulatory risk; warm... | info@tusk.vc |
| 11 | Chapter One | Sector VC | Portfolio: Compound Finance, Ondo Finance, Helius, ether.fi, Erebor bank, Mercury, Supabase, Superhuman, Together AI. | Email hi@chapterone.com or DM @jmj on X; site lists both under cont... | hi@chapterone.com |
| 12 | Multicoin Capital | Sector VC | Thesis-driven; explicit MEV and DeFi segments fit trading-infrastructure pitches. | Use the contact form and pick 'Pitch an idea for investment'; a war... | hello@multicoin.capital |
| 13 | Better Tomorrow Ventures (BTV) | Sector VC | Leads pre-seed rounds. | Warm intro from a fintech founder or a public DM to Sheel Mohnot (@... | Twitter DM @btv_vc |
| 14 | Electric Capital | Sector VC | Portfolio has several trading venues and AI-agent finance startups, so an AI trading-agent pitch is on-thesis; $1M minimum check means they expect ... | Email info@electriccapital.com; the contact page invites founders d... | info@electriccapital.com |
| 15 | Hack VC | Sector VC | Runs a quantitative research division with a proprietary trading desk and hack.summit() developer community (third-party). | Email contact@hack.vc with a deck, or DM a partner on X; the site h... | contact@hack.vc |

## Outreach playbook

Work in this order. Send the applications first, then the cold emails, then the DMs.
Log every send in a sheet with the date, so you do not contact the same person twice.

### Cold email template (6 lines)

```text
Subject: {Company} - pre-seed, automated trading + prediction markets

Hi {Name}, I build {Company}: automated trading bots and prediction-market
(Polymarket) agents. I run the system on live capital today.
Result so far: {one metric, for example "N trades, X% net return since <date>"}.
I raise a {amount} pre-seed on a SAFE to add {one use of funds}.
You back {their sector: fintech / crypto / AI agents}, so I contact you first.
Deck attached. Can I send you the live dashboard link?
```

Rules: one metric, one ask, no attachment over 10 MB, no confidential data.
Send from your own domain. Put the deck in the first message; do not ask permission to send it.

### Twitter DM template (3 lines)

```text
Hi {Name} - I build automated trading bots + Polymarket agents, live on real capital.
{one metric}. Raising a {amount} pre-seed on a SAFE.
Can I send a 10-slide deck?
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

| Field | `unknown` rows | Share of 355 |
| --- | --- | --- |
| Timing / deadline | 303 | 85% |
| Other channels | 293 | 83% |
| Partner Twitters | 266 | 75% |
| LinkedIn | 211 | 59% |
| Twitter | 178 | 50% |
| Sectors | 146 | 41% |
| Stage | 145 | 41% |
| Email or form | 116 | 33% |
| Check size / terms | 38 | 11% |
| Best channel | 18 | 5% |
| Outreach rule / cold pitch policy / how to apply | 12 | 3% |
| Geography | 7 | 2% |

Other gaps you must accept:

- Only 22 of 355 rows name a prediction market, Polymarket, or Kalshi anywhere in
  their text, and 18 of those are OPEN. Sector fit is a score from the text, not a promise.
- 36 rows give a warm intro as the best channel. A warm intro is not in this workbook. You must
  find the path yourself.
- Check sizes come from third-party aggregators for many funds. Treat them as a range, not a quote.
- The open-channel rows have no `stage` field. Their status comes from the notes text only.
- 51 rows share a name with a row in another source file. The `Also in` column marks them.
  The two rows can disagree, because they were researched from different sources.

## Rebuild

```bash
python3 /home/kreshnith/investors-preseed/build_workbook.py
```

Verified date: 2026-09-02. Deadlines change. Confirm each one on the source URL before you apply.
