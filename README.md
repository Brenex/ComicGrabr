# ComicGrabr

ComicGrabr is a Python automation tool that synchronizes a comic pull list from [League of Comic Geeks](https://leagueofcomicgeeks.com/), searches for released issues through the AirDC++ Web API, validates candidate files, queues valid downloads, and automatically retries unresolved comics until they are found.

The script is designed to run unattended, including from cron or Unraid User Scripts.

## Table of Contents

- [Features](#features)
- [How ComicGrabr Works](#how-comicgrabr-works)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [State Files](#state-files)
- [Usage](#usage)
- [Automatic Retry Queue](#automatic-retry-queue)
- [Comic Matching and Validation](#comic-matching-and-validation)
- [Scheduling](#scheduling)
- [Logging](#logging)
- [Discord Notifications](#discord-notifications)
- [Troubleshooting and Notes](#troubleshooting-and-notes)

## Features

- **League of Comic Geeks integration**
  - Logs into your LCG account.
  - Downloads the current pull-list Excel export.
  - Rebuilds `pull_list.json` with current and upcoming releases.

- **Persistent retry queue**
  - Failed searches and failed download attempts are stored separately in `retry_list.json`.
  - Retry entries survive normal LCG refreshes even after their original release dates have passed.
  - Pending retries are processed automatically on normal runs, including non-Wednesdays.
  - Successfully queued comics and comics AirDC++ reports as already present are removed from the retry queue.
  - When there is nothing left to retry, `retry_list.json` remains as an empty JSON list: `[]`.

- **AirDC++ search and download**
  - Searches AirDC++ hubs for release-day comics and pending retries.
  - Prefers CBZ files slightly over otherwise equivalent CBR results.
  - Queues validated files through the AirDC++ Web API.
  - Detects the AirDC++ "file already exists" condition and treats it as a resolved comic rather than a failure.

- **Strict candidate validation**
  - Requires an exact issue-number match.
  - Validates publication/cover year.
  - Compares normalized series titles before accepting a result.
  - Rejects unrelated titles, wrong issue numbers, old volumes, and malformed filenames instead of simply taking the first search result.
  - Supports a tightly constrained late-year cover-date rollover for October-December releases.

- **Safe unattended execution**
  - Persistent files, the `.env` file, temporary LCG export, and logs are resolved relative to the directory containing `comicgrabr.py`.
  - This prevents cron, Unraid User Scripts, systemd, or another launcher from accidentally writing state files into a different working directory.
  - Discord requests have a timeout so notification failures cannot block the comic-processing workflow indefinitely.

- **Discord notifications**
  - Script start/completion.
  - Successfully queued downloads.
  - Comics already present in AirDC++.
  - Search/download/API errors.
  - Upcoming releases for the next Wednesday.
  - Dry-run notifications are clearly marked.

- **Logging**
  - Console logging.
  - Timestamped log files in `logs/`.
  - Configurable log level.
  - Automatic cleanup of logs older than 56 days.

## How ComicGrabr Works

A normal run follows this workflow:

```text
Refresh LCG pull list
        |
        v
Update pull_list.json
        |
        +---- Wednesday? ---- yes ----> Search today's releases
        |
        +-----------------------------> Load retry_list.json
                                         |
                                         v
                                  Retry unresolved comics
                                         |
                 +-----------------------+-----------------------+
                 |                       |                       |
                 v                       v                       v
             Queued               Already exists             Failed
                 |                       |                       |
                 +------ remove from retry_list.json            |
                                                                 v
                                                      add/update retry entry
```

The key design choice is that the current LCG schedule and unresolved work are stored separately:

- `pull_list.json` is disposable current/upcoming schedule data rebuilt from LCG.
- `retry_list.json` is persistent unresolved state and is not rebuilt from LCG.

This prevents a comic that was missed on release day from disappearing simply because its release date is now in the past.

## Prerequisites

Before running ComicGrabr, ensure you have:

- **Python 3**
- **AirDC++ with its Web API enabled**
- **A League of Comic Geeks account**
- **An LCG pull list**
- Network access from the machine running ComicGrabr to the AirDC++ Web API

Optional:

- **Discord webhook** for notifications

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/Brenex/ComicGrabr.git
   cd ComicGrabr
   ```

2. Create a virtual environment:

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

   If installing manually:

   ```bash
   pip install requests pandas xlrd beautifulsoup4 discord-webhook python-dotenv
   ```

4. Create a `.env` file in the same directory as `comicgrabr.py`.

## Configuration

ComicGrabr loads `.env` from the directory containing the script rather than from the process's current working directory.

Example:

```dotenv
# League of Comic Geeks credentials
LCG_USERNAME="your_lcg_username"
LCG_PASSWORD="your_lcg_password"

# AirDC++ Web API
AIRDCPP_API_URL="http://127.0.0.1:5600/api/v1/"
AIRDCPP_USERNAME="your_airdcpp_api_username"
AIRDCPP_PASSWORD="your_airdcpp_api_password"

# Optional Discord notifications
DISCORD_WEBHOOK_URL="https://discord.com/api/webhooks/YOUR_WEBHOOK_ID/YOUR_WEBHOOK_TOKEN"
```

### AirDC++ URL

The URL should include the API path and trailing slash:

```text
http://127.0.0.1:5600/api/v1/
```

If AirDC++ is running elsewhere on your network, use the appropriate IP address or hostname instead.

## State Files

ComicGrabr keeps its runtime state beside `comicgrabr.py`, regardless of the working directory from which it is launched.

| File | Purpose |
|---|---|
| `.env` | Credentials and API configuration |
| `pull_list.json` | Current/upcoming releases imported from LCG |
| `retry_list.json` | Persistent unresolved comics awaiting another search/download attempt |
| `league_of_comic_geeks_pulls.xls` | Temporary LCG export; normally deleted after processing |
| `logs/` | Timestamped execution logs |

### `pull_list.json`

This file is rebuilt from LCG and intentionally contains only current/future releases. Past releases may disappear from it during synchronization.

Example:

```json
[
    {
        "comic_name": "X-Men 42",
        "release_date": "2026-12-09"
    }
]
```

### `retry_list.json`

This file contains only unresolved comics and persists independently of the LCG pull list.

Example:

```json
[
    {
        "comic_name": "X-Men 41",
        "release_date": "2026-11-11",
        "retry_count": 1,
        "first_failed_at": "2026-11-11T14:00:30",
        "last_attempt_at": "2026-11-11T14:00:30",
        "last_failure_reason": "no validated comic file found on AirDC++"
    }
]
```

When all retries are resolved:

```json
[]
```

Retry writes use an atomic temporary-file replacement to reduce the chance of leaving a partially written JSON database if execution is interrupted.

## Usage

Basic execution:

```bash
python comicgrabr.py
```

### Normal Run

```bash
python comicgrabr.py
```

Normal behavior:

1. Refresh the LCG pull list.
2. Update `pull_list.json`.
3. If today is Wednesday, search today's scheduled releases.
4. On every day of the week, process every unresolved entry in `retry_list.json`.
5. Add failed releases to the retry queue.
6. Remove retry entries when successfully queued or already present.

If LCG cannot be refreshed, ComicGrabr can still continue using its existing local data so pending retries are not blocked by an LCG outage.

### Retry Missing Comics Only

```bash
python comicgrabr.py --retry-missing
```

This:

- skips the LCG refresh;
- does not rebuild `pull_list.json`;
- processes only unresolved comics currently stored in `retry_list.json`.

The old option remains available as a compatibility alias:

```bash
python comicgrabr.py --search-past-releases
```

It now has the same behavior as `--retry-missing`. It no longer means "search every historical item in `pull_list.json`."

### Use an Existing LCG Excel Export

```bash
python comicgrabr.py --excel-file /path/to/pulls.xls
```

or:

```bash
python comicgrabr.py -f /path/to/pulls.xls
```

The supplied file is used instead of downloading a fresh LCG export. ComicGrabr updates `pull_list.json` and then continues normal processing; it does **not** delete a user-supplied Excel file.

If the supplied path does not exist, ComicGrabr attempts to download a fresh LCG export instead.

### Dry Run

```bash
python comicgrabr.py --dry-run
```

Dry-run mode searches and validates normally but:

- does not queue downloads;
- does not modify `retry_list.json`;
- marks Discord notifications as dry-run messages.

This is useful for validating search behavior and candidate matching without changing state.

### Log Level

```bash
python comicgrabr.py --log-level DEBUG
```

Supported values:

- `DEBUG`
- `INFO`
- `WARNING`
- `ERROR`
- `CRITICAL`

Default:

```text
INFO
```

Options may be combined where appropriate:

```bash
python comicgrabr.py --retry-missing --dry-run --log-level DEBUG
```

## Automatic Retry Queue

A comic is added to `retry_list.json` when it cannot be successfully resolved, such as when:

- no valid AirDC++ candidate is found;
- AirDC++ authentication/search is unavailable;
- a valid candidate is found but the download cannot be queued.

The retry entry keeps the comic's **original release date**. ComicGrabr never moves the publication date forward just to schedule another attempt.

A retry entry records:

- comic name;
- original release date;
- retry count;
- first failure time;
- most recent attempt time;
- most recent failure reason.

On subsequent normal runs, pending retry entries are processed automatically, regardless of the day of the week.

A retry entry is removed when:

- the file is successfully queued; or
- AirDC++ reports that the file already exists on disk or in its queue.

The work list is deduplicated by comic name and release date. If a comic is both today's release and already present in the retry queue, ComicGrabr processes it only once.

## Comic Matching and Validation

ComicGrabr does not blindly select the first CBZ/CBR returned by AirDC++.

Candidate filenames are parsed and checked against the requested comic.

### Issue Number

The issue number must match exactly after normalization.

Examples:

```text
019 -> 19
001 -> 1
1A  -> 1a
```

Searching for issue 19 will therefore reject issue 18, 20, 119, and unrelated results.

### Series Title

Titles are normalized before comparison. Normalization includes:

- case folding;
- punctuation cleanup;
- dash normalization;
- treating `&` and `and` equivalently.

For normal same-year results, the normalized title must meet the configured similarity threshold.

### Year Validation

Normally the filename year must equal the comic's release year.

For releases dated **October through December**, ComicGrabr also searches the following year and may accept a filename with `release_year + 1`.

This exception is intentionally strict. A next-year candidate is accepted only when:

1. the original release is in October, November, or December;
2. the issue number matches exactly;
3. the normalized series title matches exactly;
4. the candidate year is exactly one year after the actual release year.

For example, a comic released in November 2026 may legitimately be distributed as:

```text
X-Men 041 (2027) (Digital).cbz
```

But the rollover exception will not accept:

- the wrong issue number;
- a different X-Men series;
- a September release with the following year;
- a file two or more years ahead.

### CBZ Preference

When multiple otherwise valid candidates are available, CBZ receives a small preference over CBR.

## Scheduling

ComicGrabr works best when run **daily**, because pending retry entries are processed on every normal run.

Example cron entry, daily at 03:00:

```cron
0 3 * * * /path/to/ComicGrabr/.venv/bin/python /path/to/ComicGrabr/comicgrabr.py >> /path/to/ComicGrabr/cron_output.log 2>&1
```

On Wednesday, the same run also processes that day's new releases.

On other days, ComicGrabr still:

- refreshes LCG;
- updates the upcoming pull list;
- retries unresolved comics;
- reports the next Wednesday's releases.

Because ComicGrabr resolves its state paths relative to `comicgrabr.py`, the scheduler's working directory does not determine where `pull_list.json`, `retry_list.json`, `.env`, or the temporary Excel file are stored.

## Logging

ComicGrabr logs to both standard output and timestamped files under:

```text
logs/
```

Example:

```text
logs/comic_grabber_bot_2026-11-11_14-00-17.log
```

Log behavior:

- console and file verbosity follow `--log-level`;
- old ComicGrabr log files are cleaned up automatically;
- default retention is **56 days**.

For detailed candidate rejection reasons, use:

```bash
python comicgrabr.py --log-level DEBUG
```

Candidate validation decisions are also logged at INFO level where relevant, including issue/year mismatches and selected validated matches.

## Discord Notifications

When `DISCORD_WEBHOOK_URL` is configured, ComicGrabr sends notifications for events such as:

- script start and completion;
- queued downloads;
- already-existing comics;
- API/authentication/search failures;
- retry-related failures;
- next Wednesday's releases.

Discord delivery is non-critical to the download workflow. Notification HTTP calls use a finite timeout, so a Discord outage or TLS/network problem should not leave ComicGrabr blocked indefinitely.

If `DISCORD_WEBHOOK_URL` is empty, notifications are skipped.

## Troubleshooting and Notes

### A comic was not found on release day

Check:

```bash
cat retry_list.json
```

If ComicGrabr could not resolve the release, it should appear there and will be retried automatically on subsequent normal runs.

You can also force a retry-only run:

```bash
python comicgrabr.py --retry-missing
```

### `retry_list.json` is empty

An empty file containing:

```json
[]
```

is normal and means there are currently no unresolved comics.

### State files appear to be written somewhere unexpected

Current versions resolve persistent paths relative to the location of `comicgrabr.py`, not the shell's working directory.

For example, if the script is:

```text
/mnt/user/scripts/ComicGrabr/comicgrabr.py
```

then its state files will be:

```text
/mnt/user/scripts/ComicGrabr/pull_list.json
/mnt/user/scripts/ComicGrabr/retry_list.json
/mnt/user/scripts/ComicGrabr/.env
```

This is particularly important for Unraid User Scripts and cron, which may invoke programs from a temporary or unrelated working directory.

### AirDC++ returns many unrelated files

This is expected for broad hub searches. ComicGrabr validates each CBZ/CBR candidate before accepting it.

Review the logs for messages such as:

```text
Issue mismatch
Year mismatch
Title mismatch
Unable to parse candidate filename
```

These indicate that the validator is intentionally rejecting false positives.

### A November/December comic has the following year in its filename

This may be a legitimate publisher cover-date convention. ComicGrabr automatically attempts the following year for October-December releases but only accepts it with an exact normalized title and issue-number match.

### AirDC++ API problems

Verify:

- AirDC++ is running;
- the Web API is enabled;
- `AIRDCPP_API_URL` is reachable from the ComicGrabr host;
- the username/password are correct;
- the API URL contains `/api/v1/`.

### LCG login problems

Verify `LCG_USERNAME` and `LCG_PASSWORD` in `.env`.

LCG website changes may occasionally require updates to login or CSRF-token handling.

### Excel parsing problems

ComicGrabr currently expects the League of Comic Geeks `.xls` export and uses `xlrd` to read the `Pulls` sheet.

The expected columns are:

- `Comic`
- `Release`

### Duplicate downloads

If AirDC++ reports that the selected file already exists on disk or in its queue, ComicGrabr treats the comic as resolved. If the comic was in `retry_list.json`, its retry entry is removed.

## License

See the repository for licensing information.
