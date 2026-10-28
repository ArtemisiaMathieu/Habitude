# Habitide— daily habits & tasks

A habit tracker and to-do list that earns you coins, multiplies them for streaks, and lets each kind of day have its own look. It runs entirely in your browser (or as a home-screen app), and your data lives in **your own** Google Sheet. Anyone who opens the GitHub Pages link gets their own separate, empty copy — nobody can see anyone else's data.

---

## 1. Put the app on GitHub Pages (about 3 minutes)

1. On GitHub, create a new **public** repository, e.g. `habiterrium`.
2. Click **Add file → Upload files** and drag in everything from this folder **except** the `apps-script` folder (it doesn't hurt to include it). Keep the `icons` folder as a folder.
3. Commit.
4. Go to **Settings → Pages**. Under *Build and deployment*, choose **Deploy from a branch**, branch **main**, folder **/ (root)**, then **Save**.
5. After a minute your app is live at `https://YOUR-USERNAME.github.io/habiterrium/`.

The app works immediately — everything saves in that browser. Steps 2–3 add your Google Sheet.

## 2. Create your Google Sheet backend (about 5 minutes)

1. Create a new, blank Google Sheet (name it anything, e.g. *Habiterrium data*).
2. In the sheet: **Extensions → Apps Script**. Delete the sample code and paste in everything from `apps-script/Code.gs`.
3. On the line `const APP_KEY = 'CHANGE-ME-to-a-long-secret-phrase';` replace the text between the quotes with your own long secret phrase. Save (💾).
4. In the function dropdown at the top choose **setup** and click **Run**. Approve the permissions (Google will warn that the app isn't verified — that's because it's *your* script; choose *Advanced → Go to project*). This creates the tabs and a small **Dashboard** tab.
5. Click **Deploy → New deployment**. Click the gear ⚙ → **Web app**.
   - *Execute as:* **Me**
   - *Who has access:* **Anyone**
   - **Deploy**, then copy the **Web app URL** (ends in `/exec`).

> "Anyone" only means the URL can be reached without a Google login. Every request must include your APP_KEY, or the script refuses it.

## 3. Connect the app

Open your app → **Settings → Google Sheets sync** → paste the Web app URL and your APP_KEY → **Test & sync now**. The pill at the top right turns to **Synced**.

Do the same on each device (phone, laptop) and they'll all share the same sheet. The newest change wins.

> If you ever edit `Code.gs` later, use **Deploy → Manage deployments → ✏️ → Version: New version** so the URL stays the same.

## 4. Add it to your phone's home screen

- **iPhone:** open the link in Safari → Share → **Add to Home Screen**.
- **Android:** open in Chrome → ⋮ → **Add to Home screen** / **Install app**.

It opens full-screen like a regular app and still works offline (changes sync when you're back online).

---

## How it works

**Day templates** — one per kind of day (Work day, Day off, Luteal phase…). Each has its own habits in your chosen order, an optional tasks list, optional category headings, and its own skin. Each day you open the app, you pick which template you're living. You can go back to any past day and change it.

**Habits** — each has:
- **Coins** 1–5
- **Priority** — 3 *Must* (can't be skipped: missing it resets the streak *and* deducts its coins), 2 *Should* (missing it resets the streak), 1 *Optional* (skip freely, nothing lost)
- **Category** with a color (Health, Hygiene, Dog care…)
- One habit can belong to several templates; its streak follows it across all of them.

**Streaks & multiplier** — A streak only counts days whose template includes that habit; days without it are invisible to the streak (Mon + Thu habit done both days = unbroken streak). For every full week a streak runs, the bonus goes up by your chosen step (default +0.25×) until your chosen cap (1.25×–3×, default 3×). Set both in **Settings → Scoring**. Today never counts as "missed" until it's over.

**Tasks** — a running list with groups, due dates and priority; sort by priority, due date or group; filter by group; see open or completed. **Add to Today's List** asks how many coins (1–5) it's worth today. Ticking it off on the Day page earns the coins and closes the task. Quick-add on the Day page creates a task and puts it on that day at once.

**Skins** — Settings → Skins → **Create skin**: upload a JPG/PNG style reference, describe the feel you want, generate, fine-tune every color, the fonts and the corner roundness, then save and apply it to any templates (or make it the default).
- Without an API key, the skin is made on your device from the image's colors and keywords in your prompt (*dark, pastel, bold, elegant, playful, modern, cozy…*).
- With a Claude API key (Settings → AI skin designer), Claude designs it from the image and prompt. The key stays on that device only.
- **Download** a skin to share it as a small `.habitude-skin.json` file; others use **Upload** to add it to their collection. Uploaded skins are checked so they can only contain colors, fonts and an image.

**Progress** — coin totals, streaks and multipliers, category completion, a 30-day chart, and **Ask your data**: choose what to show (coins, completion %, habits done, tasks done, coins lost), group it by day/week/month/habit/category/template/task group, filter by date range, category or template, and see it as a chart or table.

**Exports** — full or custom (tick the data types + date range):
- **Visual report** — opens a printable page (save as PDF) or downloads it as an HTML file.
- **Data spreadsheet** — an `.xlsx` with one tab per data type (or CSV).

**Your Google Sheet** — tabs for every data type (Habits, Templates, Tasks, HabitLog, DayTasks…), plus read-only `DailySummary`, `HabitStats` and a `Dashboard` with formulas. Edit your data in the app; the sheet is the permanent record and is rewritten on each sync.

**Backups** — Settings → *Your data* → download a single backup file any time, or restore one.

## Files

| File | What it is |
|---|---|
| `index.html` | The page |
| `styles.css` | Layout and design (all colors come from the active skin) |
| `app.js` | Everything the app does |
| `sw.js`, `manifest.webmanifest`, `icons/` | Home-screen install and offline support |
| `apps-script/Code.gs` | Paste into your Google Sheet's Apps Script |
