# Operation Daily Planning portal

A static web portal (GitHub Pages) for the Operation Daily Report.
It keeps the **last 7 report days**; older days are deleted automatically.

## Folder layout

| Path | What it is |
|---|---|
| `index.html` | The portal (Overview, Fleet, FTTx-Ops, Box-Ops, date buttons). |
| `data/` | Converted report days (`YYYY-MM-DD.json`), `days.json`, `last_upload.json`. Written by the robot – don't edit by hand. |
| `uploads/` | Put the daily Excel here. The robot converts it and removes it. Files with errors move to `uploads/failed/`. |
| `template/Operation_Daily_Report_Template.xlsx` | The Excel the team fills every day. |
| `scripts/convert.py` | Excel → data converter (checks the file, keeps 7 days). |
| `.github/workflows/portal.yml` | Runs the converter and publishes the site after every upload. |

## One-time setup (about 5 minutes)

1. Create a new repository on GitHub (e.g. `ops-portal`) and upload **all** files and folders from this package, including the hidden `.github` folder.
   *Tip: GitHub's web upload skips folders starting with a dot. If `.github/workflows/portal.yml` is missing after upload, create it in GitHub with **Add file → Create new file**, name it `.github/workflows/portal.yml` and paste the file's content.*
2. Repository **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. Repository **Settings → Actions → General → Workflow permissions: Read and write permissions** → Save.
4. Open the **Actions** tab → **Update portal** → **Run workflow** once.
5. The portal address appears in Settings → Pages (e.g. `https://<your-account>.github.io/ops-portal/`). Link it from your main portal.

Private repository? GitHub Pages for private repositories needs a paid GitHub plan (Pro, Team or Enterprise). On Enterprise you can also limit who can view the site.

## Daily use (team)

1. Fill `template/Operation_Daily_Report_Template.xlsx` (or yesterday's copy). Set the date on **Report info**.
2. Optional: on the portal click **✓ Check Excel** to see any mistakes before uploading.
3. Click **⬆ Upload daily Excel** on the portal (opens GitHub's upload page for `uploads/`), drop the file, click **Commit changes**.
4. After 1–2 minutes the portal shows the new day. The header shows the result of the last upload; if it failed, it lists what to fix.

Uploading a file for a date that already exists replaces that day. Uploading the 8th day removes the oldest.
Team members need write access to the repository to upload.

## Settings

In `index.html`, `PORTAL_CONFIG.githubRepo` can be set to `"owner/repo"` if the site is not on a `github.io` address (for example a custom domain), so the Upload button knows where to go.
To keep more or fewer days, change `KEEP_DAYS` in `scripts/convert.py` and the `7` in the portal's check messages.
