# GMC daily report: setup guide

This folder holds a small program that makes a **daily Google Merchant Center report** and emails
it. Each time it runs, it:

1. reads every account under one parent Merchant Center account from Google's Merchant API (about
   150 requests, roughly a minute),
2. adds that day's figures to a history file (`data\gmc_cache.sqlite`), so later reports can compare
   with earlier days,
3. builds an Excel report with six sheets (account overview, sudden drops, top issues, a priority
   list, every disapproved product, and approval rate over time),
4. saves the report into a **OneDrive folder**, where a **Power Automate flow** emails it.

Windows Task Scheduler starts it once a day. Nothing else needs to be running on the PC.

**Every run sends an email**, including a run you start by hand. The only run that sends nothing
is `--check`.

You never edit the Python files. Everything you change is in **`settings.env`**, and the steps
below give the exact line number.

---

## What you need

- A **Windows PC** that is **switched on and signed in** at the time you schedule. OneDrive only
  uploads while you are signed in.
- **Internet access** to `pypi.org` (to install libraries, once) and to
  `merchantapi.googleapis.com` (every day).
- **OneDrive** (work or school) signed in on this PC, and **Power Automate** and **Outlook** on the
  same Microsoft 365 account.
- This folder: either a **zip file** you were given, or access to the **GitHub repository** it is
  kept in.
- The **Google service account key file** (a `.json` file) and the **service account's email
  address** (it ends in `.iam.gserviceaccount.com`).
- The **history file** (step 6). The key and the history file are never on GitHub: you get them
  privately from the person who manages the report.
- Access to **Google Merchant Center** for the parent account, or someone who has it (step 1).

---

## Step 1: Give the service account access, and note the account ID

The program signs in to Google as the service account. That account must be a user on the
**parent** Merchant Center account (the one all the accounts sit under).

1. Open <https://merchants.google.com> and switch to the **parent** account.
2. Go to **Settings → Access and services**. If the service account's email address is already in
   the list with **Admin** access, skip to point 4.
3. Otherwise: **Add user** → paste the service account's email address → choose **Admin** → save.
   You need Admin rights on the parent account to do this. Access to the parent account covers
   every account under it, so you don't add it account by account.
4. Note the **parent account ID**: the number shown next to the account name at the top of
   Merchant Center (it also appears in the browser's address bar after `a=`). Numbers only. You
   need it in step 8.

> **Six accounts are not under the parent account.** The program also reads these, because they
> are on its account list. Access to the parent account does **not** cover them: the service
> account needs its own access to each one.
>
> | Account | Account ID |
> | :--- | :--- |
> | LG España | 130445164 |
> | LG Electronics France | 462116364 |
> | LG Electronics Belgium FR | 5815616819 |
> | LG Electronics Belgium NL | 5815014784 |
> | LG Chile | 527424218 |
> | LG Electronics Netherlands | 5438723367 |
>
> If you were given the same key that already runs this report, these accesses are already in
> place and there is nothing to do. With a different service account, switch to each of these
> accounts in Merchant Center and repeat points 2 and 3 there.
>
> An account without access doesn't stop the run, but it has no figures. After the first run
> (step 11), open the report's **Overview** sheet and check that none of the six shows the status
> **Not in last sync** with blank cells. If one does, the log (`data\logs\gmc.log`) shows an error
> for its account ID: give that account access, then the next run picks it up.

## Step 2: Install Python 3.14

1. Download Python **3.14** for Windows from <https://www.python.org/downloads/windows/>.
2. Run the installer and keep the defaults. That also installs the `py` command used below.
3. Check it: open **PowerShell** and type `py -3.14 --version`. It should print `Python 3.14.x`.

## Step 3: Get the folder (outside OneDrive)

**If you were given a zip file:**

1. Right-click the zip → **Extract All…** → choose a folder such as `C:\GMC\`.
2. You now have `C:\GMC\gmc-report\` with `run_report.py` inside.

**If it's on GitHub:**

1. Sign in to GitHub, open the repository, click the green **Code** button → **Download ZIP**.
2. Right-click the downloaded zip → **Extract All…** → choose `C:\GMC\`.
3. The folder inside is named after the repository, for example `gmc-report-standalone-main`.
   **Rename it to `gmc-report`**, so you have `C:\GMC\gmc-report\` with `run_report.py` inside.

> **Do not put this folder inside OneDrive or a shared drive.** It will hold the Google key and the
> history file (client data), and inside OneDrive they would be uploaded to the cloud. Only the
> *report* goes to OneDrive (step 7).

The rest of this guide calls this folder **the runner folder** and assumes `C:\GMC\gmc-report`. If
yours is elsewhere, use your path instead.

## Step 4: Install the libraries

In PowerShell:

```
cd C:\GMC\gmc-report
py -3.14 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-standalone.txt
```

The first command enters the folder. The second makes a private copy of Python just for this
folder (`.venv`). The third installs the exact library versions the program was tested with. It
ends with `Successfully installed …`.

## Step 5: Put the key file in place

1. Copy the service account key file into the runner folder's `secrets` folder.
2. Rename it to exactly **`sa-key.json`**, so the full path is `C:\GMC\gmc-report\secrets\sa-key.json`.
3. You can delete `secrets\PUT-sa-key.json-HERE.txt`.

> The key can read every account. Move it directly (USB stick, or a private folder
> only you can open). **Never email it or post it in a chat.**

## Step 6: Put the history file in place

The report compares today with earlier days (the "Spike Radar" and "Approval Rate Over Time"
sheets). That history comes from the **history file** you were given with this zip
(`gmc_cache-copy-<date>.sqlite`). **This step is required:** the setup check in step 9 fails
without it. If you didn't receive the file, ask for it before going on.

1. Copy it into the runner folder's `data` folder.
2. Rename it to exactly **`gmc_cache.sqlite`**, so the full path is
   `C:\GMC\gmc-report\data\gmc_cache.sqlite`.
3. You can delete `data\PUT-gmc_cache.sqlite-HERE.txt`.

From then on, every run adds that day's figures to this file.

> The history file holds client product data. Handle it like the key.

## Step 7: Make the OneDrive folder for the reports

1. In File Explorer, open **OneDrive** and create a new folder, for example **`GMC Reports`**.
2. Use a folder that **only this PC** saves reports into. If another computer already saves this
   report into a folder that a flow emails from, make a different folder, or the team gets two
   emails a day.
3. Open the new folder, click the address bar and copy the full path, for example
   `C:\Users\you\OneDrive - Company\GMC Reports`.

## Step 8: Fill in `settings.env`

1. In the runner folder, copy `settings.env.template` and rename the copy to **`settings.env`**.
   (In File Explorer, turn on **View → Show → File name extensions** first, so the file doesn't
   end up named `settings.env.txt`.)
2. Open `settings.env` in **Notepad** and change these lines. Type straight after the `=`, with no
   quotes:
   - **line 7** (`PARENT_MCA_ID=`): the parent account ID from step 1.
   - **line 12** (`REPORT_DROP_DIR=`): the OneDrive folder path from step 7.
3. Leave the other lines as they are unless you have a reason:
   - **line 16** (`FETCH_TRIES=`) and **line 17** (`FETCH_RETRY_MINUTES=`): if Google fails, how
     many tries per run and the minutes between them (3 tries, 30 minutes apart).
   - **line 19** (`SAVE_TRIES=`) and **line 20** (`SAVE_RETRY_MINUTES=`): if saving the Excel file
     fails, how many tries and the minutes between them. These retries make no new Google requests.
   - **line 24** (`GCP_SA_KEY_PATH=`), **line 25** (`DB_PATH=`), **line 26** (`LOG_DIR=`) and
     **line 27** (`STORE_DIRECTORY_PATH=`): only if you keep the key, the history file, the log or
     the account list somewhere other than the runner folder. Empty means the runner folder.
4. Save the file.

## Step 9: Check the setup (sends nothing)

```
cd C:\GMC\gmc-report
.venv\Scripts\python.exe run_report.py --check
```

Each line starts with `[OK     ]`, `[WARNING]` or `[PROBLEM]`. The last line should read
**`All set.`** It checks Python, the key, the account ID, the report folder, the history file, the
account list, signing in to Google, and whether this network can reach the Merchant API. It reads no
account data and saves nothing.

- `[PROBLEM]`: fix that item and run `--check` again.
- "Google sign-in failed": the key file is wrong or has been deleted in Google Cloud. Ask whoever
  gave you the key for a new one.
- "Reaching the Merchant API failed": this network blocks `merchantapi.googleapis.com`. Ask IT to
  allow it.
- `[PROBLEM]` "no history file", or "the history file … is empty": do step 6 (replace the empty
  file with the one you were given).
- `[WARNING]` about OneDrive or the folder name: re-read steps 3 and 7.

`--check` doesn't test whether the service account can see the accounts. The first real run
(step 11) does. If that run fails with "permission denied" or "403" in the log, re-check step 1.

## Step 10: Set up the email (Power Automate)

The flow watches the OneDrive folder and emails each new report. Build it once at
<https://make.powerautomate.com>, signed in with the same Microsoft account as OneDrive:

1. **Create → Automated cloud flow.** Name it, for example `GMC report email`.
2. **Trigger:** **OneDrive for Business → When a file is created**. For **Folder**, **click the
   folder icon and pick the folder from step 7. Don't type the path:** a typed path makes every
   run fail with `Invalid fileId`.
3. **Condition** (so only finished reports are sent, never a half-written file):
   - `x-ms-file-name` **contains** `gmc-report-`
   - **And** `x-ms-file-name` **ends with** `.xlsx`

   Use `x-ms-file-name`, **not** `x-ms-file-name-encoded`. The encoded one never matches.
4. Under **True**, add **Office 365 Outlook → Send an email (V2)**:
   - **To**, **Subject** and **Body**: your choice.
   - **Attachments**: switch to the per-field view and add one attachment. Set **Name** to
     `x-ms-file-name` and **Content** to the trigger's **File content**. Don't add a
     "Get file content" step; the trigger already has the file.
5. **Optional: a summary in the email body.** The runner folder has `email-summary.ts`, a script
   that reads the report and writes a short summary (overall figures, weak accounts, biggest drops,
   key issues):
   1. **Before saving it, edit the top of the file** (`GREETING`, `SIGN_OFF` and the
      `SIGNATURE_LINES` list) to your own wording and signature.
   2. Open Excel on the web → **Automate → New script**, replace everything with the edited file,
      and save it as `email-summary`.
   3. In the flow, before the email step, add **Excel Online (Business) → Run script**: Location
      **OneDrive for Business**, Document library **OneDrive**, File = the trigger's `x-ms-file-id`,
      Script = `email-summary`.
   4. In **Send an email (V2)**, set **Body** to the script's **result**.
6. **Save** the flow and make sure it's turned **on**.

## Step 11: First run by hand (this sends a real email)

```
cd C:\GMC\gmc-report
.venv\Scripts\python.exe run_report.py
```

It takes about a minute. The last log line reads `standalone run done: … report …`. Then check:

1. the new `gmc-report-….xlsx` is in the OneDrive folder, and also at onedrive.com (proof it
   uploaded);
2. Power Automate → your flow → **Run history** shows a green run;
3. the email arrives with the report attached.

## Step 12: Schedule it (Task Scheduler)

1. Start menu → **Task Scheduler** → **Create Task…** (not "Create Basic Task").
2. **General** tab:
   - Name: `GMC report`
   - Choose **Run only when user is logged on**. OneDrive only uploads while you are signed in, so
     a run while signed out would leave the report on this PC.
3. **Triggers** tab → **New…** → **Daily**, at the time you choose → OK.
4. **Actions** tab → **New…** → **Start a program**:
   - Program/script: `C:\GMC\gmc-report\.venv\Scripts\pythonw.exe`
     (`pythonw` runs without a window, so nobody closes it by accident)
   - Add arguments: `run_report.py`
   - Start in: `C:\GMC\gmc-report` (no quotes, no `\` at the end)
5. **Conditions** tab: on a laptop, untick **Start the task only if the computer is on AC power**.
6. **Settings** tab:
   - tick **Run task as soon as possible after a scheduled start is missed** (the PC was asleep or
     off);
   - tick **Stop the task if it runs longer than** → **3 hours** (a run with every retry takes
     about 1½ hours at most);
   - **If the task is already running** → **Do not start a new instance**.
   - Leave **If the task fails, restart every…** off. It doesn't react to a failed run of a
     script. The retries are built into `run_report.py` (lines 16–20 of `settings.env`).
7. OK.

To test the schedule: right-click the task → **Run**. **This sends an email.**

## Every day: did it work?

Task Scheduler → the task → **Last Run Result**:

| Shown | Meaning | What to do |
| :--- | :--- | :--- |
| `(0x0)` | Done: report saved | Nothing |
| `(0x1)` | Google failed on every try, so there's no report and nothing was saved | Usually a Google or internet outage. The next scheduled run tries again. If it repeats, run `--check` |
| `(0x2)` | Figures saved, but the **Excel file was not saved**, so no email | Fix the folder (OneDrive running? folder still there?), then run `.venv\Scripts\python.exe run_report.py --resave`. That rebuilds the same report with no Google requests |
| `(0x3)` | Not set up (key, account ID or folder missing) | Run `--check` and fix what it lists |
| `(0x4)` | Another run was still going, so this one didn't start | Nothing, unless it keeps happening |
| `(0x5)` | Something unexpected broke | Look at the end of the log (below) |
| `(0x41301)` | It is running right now | Wait |

**The log** is `data\logs\gmc.log` in the runner folder. It records every Google request, every
retry and every saved report. Useful lines: `standalone run done`, `fetch failed`,
`report save failed`.

When an email doesn't arrive but the result is `(0x0)`: check that the file reached
onedrive.com, then check the flow's Run history. Nothing in this folder can see whether an email
was delivered.

## Updating to a new version

When there is a new version (a new zip, or an update on GitHub):

1. Delete the **`app`** folder in the runner folder.
2. Get the new version as in step 3, but extract it somewhere else (not over the runner folder).
   Open the extracted folder (`gmc-report`, or the `…-main` folder from GitHub) and copy
   **everything in it** into the runner folder. Choose **Replace** when asked.
   This **keeps** `data\`, `secrets\`, `settings.env` and `.venv\`, because the zip has none of them.
3. Run `.venv\Scripts\python.exe -m pip install -r requirements-standalone.txt` again (it's quick
   when nothing changed).
4. Compare `settings.env` with the new `settings.env.template`. If the template has a new line, add
   it.
5. Run `--check`.

`VERSION.txt` says which version the folder holds.

## Rules that matter

- **Never delete `data\gmc_cache.sqlite`.** It's the only copy of this PC's history. Now and then,
  copy it somewhere safe and private.
- **Only this PC saves into the folder on line 12.** If two computers save into folders that send
  email, the team gets two emails a day.
- **Every run sends an email**, including Task Scheduler's **Run** button and a hand run. Only
  `--check` sends nothing, and `--resave` sends a report only if it wasn't saved before.
- **Keep the runner folder out of OneDrive and shared drives** (key and client data).
- **Never upload the key, the history file or `settings.env` to GitHub**, or send them by email or
  chat.
