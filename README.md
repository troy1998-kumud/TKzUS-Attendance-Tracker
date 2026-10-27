# TkzUS Attendance

Clock in, clock out and time off for TkzUS staff, with an HR view. It runs as a Google Apps Script web app and stores everything in a Google Sheet.

- Staff open one web link and sign in with their **Employee ID and PIN**. They don't need a Google or Claude account.
- HR opens the same link, taps **Open the HR view** and enters the **HR PIN**.
- All times are **India Standard Time**. A shift that ends after midnight counts toward the day it started.

> GitHub stores the source code. The website itself runs on Google Apps Script, not on GitHub Pages.

## Folder layout

```
src/
  Code.gs            Server code: sheet setup, sign in, clock in/out, leave rules, HR data
  Index.html         The whole web page (staff and HR screens, TkzUS logo embedded)
  appsscript.json    Apps Script settings: IST time zone, web app runs as owner, open to anyone
preview/
  build_preview.py   Builds a local preview with a fake Google Sheet, for trying changes
  gasmock.js         The fake Google services used by the preview
.clasp.json.example  Template for pushing with clasp (optional)
.gitignore
README.md
```

## Company rules built in

| Rule | Value | Where to change it |
|---|---|---|
| Work week | Monday to Friday, weekends off | `Code.gs` (`dow` checks) |
| Default shift | 6:00 PM to 3:00 AM IST | Settings tab |
| Kumud Srivastava, Nikhil Katta | 4:00 PM to 1:00 AM IST | Employees tab, Shift start/end |
| Allowed difference | 45 minutes late or early, no flag | Settings tab |
| Short day | Under 8 hours from clock in to clock out (9-hour shift includes a 1-hour break) | Settings tab |
| Overtime | None | — |
| Paid leave | 14 days a year | Settings tab, or per person in Employees |
| Sick leave | 5 days a year | Settings tab, or per person in Employees |
| Leave year | January 1 to December 31, no carry-over | `balances()` in `Code.gs` |

## Setup, the simple way (copy and paste)

1. Go to **sheets.new** and name the sheet **TkzUS Attendance**.
2. Open **Extensions → Apps Script**.
3. Replace everything in `Code.gs` with `src/Code.gs` from this repo.
4. Click **+ → HTML**, name it exactly **Index**, and paste in `src/Index.html`.
5. (Optional) Open **Project Settings**, tick **Show "appsscript.json" manifest file**, and paste in `src/appsscript.json`.
6. Save. Pick **setup** from the function list and click **Run**. Allow access when Google asks (**Advanced → Go to … (unsafe) → Allow**; it says unsafe only because Google hasn't reviewed your own script).
7. Click **Deploy → New deployment → Web app**. Set Execute as **Me** and Who has access **Anyone**, then click **Deploy**. Copy the **Web app URL** and share it with staff.

`setup` creates four tabs:

- **Employees**: the 10 staff, their shifts, and a random 4-digit **PIN** each. Give each person only their own PIN.
- **Settings**: the 6-digit **HR PIN** and the company rules above.
- **Attendance**: one row per person per shift, written by the app.
- **Leave**: every leave request. HR approves in the app, or by changing the Status cell.

Running `setup` again is safe. It never overwrites tabs that already have data.

## Setup with GitHub and clasp (optional)

Use this if you want to push from your computer instead of copying and pasting.

```bash
npm install -g @google/clasp
clasp login
# Create the Sheet first (step 1 above), open Extensions → Apps Script,
# then copy the Script ID from Project Settings.
cp .clasp.json.example .clasp.json   # paste your Script ID into it
clasp push                           # uploads src/ to your script
```

Then run `setup` and deploy as in steps 6 and 7. The script must be created from inside the Sheet (bound to it), because `setup` uses that Sheet.

`.clasp.json` holds your script ID and is git-ignored. Don't commit it.

## Updating the live site

After changing code, use **Deploy → Manage deployments → pencil → Version: New version → Deploy**. This keeps the same link. **New deployment** would create a new link.

## Trying changes locally

```bash
python3 preview/build_preview.py
# open preview/preview.html in a browser
```

The preview uses a fake Google Sheet that lives in the browser tab and resets when you reload. The browser console prints the staff PINs and the HR PIN.

## Day to day

- **Add a person**: add a row in Employees with ID, name, PIN, and Active set to **Yes**. Leave shift and leave days blank to use the defaults.
- **Remove a person**: set Active to **No**. Their history stays in the sheet.
- **Reset a PIN**: edit the PIN cell.
- **Fix a wrong punch**: edit the row in the Attendance tab.
- **Export**: use **File → Download** in the Sheet.

## Security notes

- PINs are checked on the server, so staff can only clock in as themselves and can't see anyone else's data.
- Anyone with the link can see the sign-in page. Keep PINs private and change any PIN that leaks.
- The HR PIN opens everyone's data. Only tick "Remember on this device" on your own computer.
- Don't commit real PINs to GitHub. They live only in your Google Sheet, not in this code.
