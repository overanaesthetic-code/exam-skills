# Connect IELTS Ladder to Google Sheets

Takes about 5 minutes. Until you finish, the course works exactly as before (progress stays on the student's device).

## 1. Make the sheet
1. Go to sheets.new and name it, e.g. **IELTS Ladder · progress**.
2. Menu **Extensions → Apps Script**.
3. Delete everything in the editor, paste the whole of `Code.gs` from this folder, press **Save** (the disk icon).

## 2. Publish it as a web app
1. Top right: **Deploy → New deployment**.
2. Gear icon next to "Select type" → **Web app**.
3. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone
4. **Deploy**. Google asks for permission → choose your account → *Advanced* → *Go to … (unsafe)* → **Allow**.
   (It says "unsafe" only because the script is yours and not reviewed by Google. It only touches this one sheet.)
5. Copy the **Web app URL**. It ends in `/exec`.

## 3. Turn it on in the course
Open `config.json` in the repo and paste the URL:

```json
"sync": { "url": "https://script.google.com/macros/s/…/exec" }
```

Commit. In a minute or two GitHub Pages updates and syncing starts.

## What you'll see
The tabs appear after the first student syncs:

| Tab | What's in it |
|---|---|
| **Students** | One row per code: label, placement, level per skill, sections done, average %, last active |
| **Results** | Every scored section and drill, with unit, tier and score |
| **Essays** | Every writing task with the full text, word count and level |
| **Speaking** | Every speaking task: level, words per minute, fillers |
| `_state` (hidden) | Full copy of each learner's progress, used to restore it on a new device |

- Progress is sent about 5 seconds after each change, and when the student leaves the page.
- A student who opens the course on a new phone with the **same code** gets their progress back.
- No names or emails travel. A row is the code plus the label you gave that code.

## If you change Code.gs later
**Deploy → Manage deployments → pencil icon → Version: New version → Deploy.** The URL stays the same.

## Good to know
- Anyone who has a code can restore that code's progress. That's the same person who can open the course, so it matches how access already works.
- Don't share the sheet itself; it's private to your Google account.
