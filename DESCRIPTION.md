# SheetsDemon

**A local Excel tutor and formula auditor that works on the workbook you actually have open.**

---

## One-line (GitHub "About" field)

> Local AI Excel tutor and formula auditor — parses formulas properly instead of guessing, reads your live workbook over AppleScript, and works with no API key.

---

## Short description

SheetsDemon is a desktop companion for learning spreadsheets. It sits beside
Microsoft Excel, reads the workbook you have open — including unsaved edits and
live calculated results — and tells you what is wrong with your formulas and
why.

It was built for an introductory university spreadsheets course, which shaped
two decisions that run through the whole project. First, it tutors as readily as
it answers: ask for a hint and it gives you a graded nudge rather than the
solution. Second, it refuses to guess. Every claim about a formula comes from a
real tokenizer, not a regular expression, because a confidently wrong answer
costs a student marks they cannot get back.

---

## Longer description

Most spreadsheet "AI helpers" read your formula as a string and improvise.
That fails in the exact places beginners need help: a `SUM` range that stops one
row short still returns a number, a relative reference still fills without
complaint, and nothing looks broken until the marks come back.

SheetsDemon takes the opposite approach. It ships a formula engine that
tokenizes Excel syntax properly, converts to R1C1, simulates what a fill would
do, and checks ranges against the data actually sitting in the sheet. The audits
are built to be trustworthy rather than thorough: they were tuned against real
graded coursework until they produced zero false positives on correct work,
because an auditor that cries wolf is worse than no auditor.

**What it checks.** Hardcoded constants buried inside formulas, reference types
that will break when filled, aggregate ranges that miss rows or swallow totals,
lookup syntax, number-format consistency, and header/grid structure. Findings
come with the specific cell, the offending formula, and the corrected version —
and `fix it` applies the ones it can prove.

**What it builds.** Native Excel charts, real Excel Tables with filter buttons
and number formats, fresh workbooks, mock datasets, and practice sheets that
vary a concept without handing over the answer.

**How it talks to Excel.** On macOS it drives the running Excel application over
AppleScript, so it sees unsaved work and live calculated values — something a
file-based tool reading the `.xlsx` fundamentally cannot do. It can write
formulas straight into the sheet and you watch them recalculate.

**It works with no API key at all.** Without a provider configured, the chat box
routes on keywords and answers from the engine directly: audits, fixes, hints,
charts, tables and new spreadsheets all function offline, with no signup and no
network. Add a key and a language model takes over the routing and the prose.
Eight providers are supported behind one adapter — Anthropic, Gemini, Groq,
Mistral, Cerebras, OpenRouter, NVIDIA NIM, and local Ollama, which needs no key
or account and is auto-detected when running.

---

## At a glance

| | |
|---|---|
| Interface | Three-panel browser command centre + CLI (`./sd`) |
| Engine | ~9,000 lines of Python, no runtime services |
| Model tools | 22, spanning audit, tutor, build and live-Excel actions |
| Providers | 8, including a zero-key local option |
| Tests | 411 assertions, plus 90 live-Excel checks |
| Requires | Python 3.10+; live Excel features need macOS + Microsoft Excel |

---

## What it is not

It is not a spreadsheet, and it does not replace Excel. It is not a cloud
service — nothing leaves your machine unless you configure a hosted model, and
keys live outside the repository in `~/.config/sheets-doctor/env`. It is not a
general-purpose agent: the tools reach your workspace folder and the workbook
you have open, and nothing else.
