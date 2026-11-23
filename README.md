# GMC daily report (standalone)

Makes the daily Google Merchant Center report on a Windows PC: it reads every account from the
Merchant API, builds a six-sheet Excel report and saves it into a OneDrive folder, where a Power
Automate flow emails it. Windows Task Scheduler starts it once a day.

**To set it up, follow [INSTRUCTIONS.md](INSTRUCTIONS.md).**

Two files are deliberately **not** in this repository: the Google key file and the history file
(client data). Get them privately from the person who manages the report, and **never upload
them here**. `.gitignore` keeps them, `settings.env` and the logs out of git.
