# Setup guide

This guide takes you from zero to a working spreadsheet and Telegram bot. It is written for people who have never used Apps Script. *¿En español? Lee [SETUP.es.md](SETUP.es.md).*

**Only want the spreadsheet, without the bot?** Do Step 1 and Step 7 and stop there. The bot is optional.

**Want to try it first?** Upload one of the files in `demo/` (fake data) the same way as in Step 1 and look around.

## What you need

About 30 minutes, once, on a computer. After that everything happens from Telegram on your phone. No programming, no server, no cost.

- **A Google account** (a normal Gmail account works; some work or school accounts block the scripts this needs).
- **Telegram** on your phone.
- **A computer with a web browser** for the setup (Chrome works best). The Google Sheets phone app can't do these steps.
- **The two files** from this repository: the spreadsheet template (English `templates/Family_Budget.xlsx` or Spanish `templates/Presupuesto_Familiar.xlsx`) and the bot file `bot/FamilyBudgetBot.gs`. Download them from the latest [Release](../../releases) or with **Code ▸ Download ZIP**.

The steps below use the English names. In the Spanish template the same menus are in Spanish (for example *Presupuesto familiar* instead of *Family Budget*).

## Step 1: Put the spreadsheet in Google Drive

The spreadsheet must become a real Google Sheet, not an Excel file, or the next steps won't appear.

1. Go to [drive.google.com](https://drive.google.com), click **+ New ▸ File upload** and choose the spreadsheet template.
2. Double-click the uploaded file to open it.
3. Click **File ▸ Save as Google Sheets**. A new copy opens in a new tab.
4. Check the top of the new copy: next to the file name there must be **no green "XLSX" label**. From now on use only this copy; you can delete the Excel one.
5. Click **File ▸ Settings** and set:
   - **Locale**: your country (for example Spain or United Kingdom). Avoid "United States" if you write dates day first.
   - **Time zone**: yours.
6. Open the **Transactions** tab and delete row 2 (the grey example row).

### Shortcut: "Make a copy" link

If the project page offers a **Make a copy** link, click it while signed in to Google. You get your own Google Sheet with the bot code already inside, so you can skip Step 1 points 1–4 and Step 3 points 1–4. Continue with File ▸ Settings (Step 1 point 5), then Step 2.

## Step 2: Create your Telegram bot

Telegram's official helper, BotFather, makes the bot for you and gives you its password (the "token").

1. In Telegram, search for **@BotFather** (it has a blue check mark) and tap **Start**.
2. Send `/newbot`.
3. Type a **name** people will see, for example *Family Budget*.
4. Type a **username** that ends in `bot`, for example *smith\_family\_bot*. If it's taken, try another.
5. BotFather replies with a long **token** like `123456789:AAH...xyz`. Copy it and keep it somewhere safe. **Never share it**: whoever has it controls the bot.
6. Optional but recommended: send `/setjoingroups`, pick your bot, then **Disable**, so nobody can add it to a group.

## Step 3: Paste the bot code into the spreadsheet

The bot lives inside your spreadsheet, so open the code editor **from the spreadsheet itself**, not from Google Drive.

1. In your Google Sheet, click **Extensions ▸ Apps Script**. A new tab opens with a little code in it.
2. On your computer, open *FamilyBudgetBot.gs* with a text editor (on a Mac: right-click ▸ Open With ▸ TextEdit; on Windows: Notepad). Select everything and copy it.
3. Back in Apps Script, click inside the code, select everything and delete it. Paste the bot code.
4. Click the **save** icon (or press Ctrl+S / ⌘S).
5. Click the **gear icon** (Project Settings) on the left and set the same **Time zone** as in Step 1.
6. Optional: click "Untitled project" at the top and rename it, for example *Family Budget Bot*.

No Extensions menu? The file is still the Excel version: go back to Step 1, point 3.

## Step 4: Publish it

Publishing gives the bot a private web address that Telegram can send your taps to.

1. In Apps Script, click the blue **Deploy ▸ New deployment**.
2. Click the gear next to *Select type* and choose **Web app**.
3. Set **Execute as: Me** and **Who has access: Anyone**.
4. Click **Deploy**, then **Authorize access**, and pick your Google account.
5. If Google says *"Google hasn't verified this app"*: click **Advanced ▸ Go to … (unsafe)**. This is normal: it's your own script, not an app from a store.
6. Tick **Select all** and click **Continue**. The first permission should mention only the spreadsheets this script is installed in, not all of them.
7. Copy the **Web app URL**. It ends in `/exec`.

**Does "Anyone" open my spreadsheet to anyone?** No. It only lets Telegram reach the bot without a Google login. Your spreadsheet stays private, the address only works with a secret key the bot adds by itself, and the bot ignores everyone who isn't on your list.

## Step 5: Connect the bot and turn on automation

1. Go back to the spreadsheet tab and **reload the page**.
2. Wait a few seconds: a new **Family Budget** menu appears at the end of the menu bar, after *Help*.
3. Click **Family Budget ▸ 1. Connect Telegram bot…**. Paste the **token** from Step 2, then the **URL** from Step 4.
4. You should see *"Connected to @your\_bot"*.
5. Click **Family Budget ▸ 2. Turn on daily automation**. From now on, every morning around 7:00 the bot adds your recurring payments and sends bill reminders.
6. Optional, if you pay in more than one currency: **Family Budget ▸ 3. Set up currencies, subcategories & date pickers**.

## Step 6: Let yourself in and invite your family

The bot only answers people on your list; everyone else gets silence.

**You (the first time)**

1. In Telegram, open your bot and send `/start`. It replies with your **Telegram ID**, a number like *123456789*.
2. In the spreadsheet, open the **Settings** tab. Find **Allowed Telegram user IDs** and paste your number in the yellow cell to its right.
3. Send `/start` again. The menu buttons appear.

**Everyone else**

1. In the spreadsheet: **Family Budget ▸ 4. Secret word for new users…** and type a word, for example *mango42*.
2. Tell that word privately to the person you want to add. They open the bot and send it.
3. You get a Telegram message with their name and two buttons: **Full access** (can add entries) or **Read-only** (can only look). One tap adds them.
4. They send `/start` and they're in.

Each person picks their language with `/english` or `/espanol`. To remove someone, delete their number from Settings.

### The access list: who can use the bot

The bot obeys only the numbers in two cells of the **Settings** tab. Anyone not listed gets no answer at all, not even an error.

| Settings cell | Who goes there | What they can do |
| --- | --- | --- |
| **Allowed Telegram user IDs** | You and anyone who adds entries | Everything: add, undo, mark bills paid, add new people, export (if turned on) |
| **Read-only Telegram user IDs** | Grandparents, an accountant… | Search, Recent, This month, Upcoming and reminders |

- Type only the numbers, separated by commas: `123456789, 987654321`. The number goes in the yellow cell to the right of the label.
- Changes work immediately; there is nothing to publish again.
- While the first list is empty, the bot shows its ID to anyone who writes, so you can set it up. Fill it in right away.

**Lock the two cells** (if other people can edit the spreadsheet)

1. In **Settings**, select the two yellow cells next to the ID labels.
2. Click **Data ▸ Protect sheets and ranges**, then **Set permissions**.
3. Choose **Only you** and click **Done**.

Now only you can change who has access. The one-tap "add" buttons in Telegram keep working, because the bot acts as you.

## Step 7: Fill in your basics

Yellow cells are the ones you type in; everything else calculates by itself. Plan about 20 minutes.

| Tab | What to do |
| --- | --- |
| **Settings** | Change *Child 1*, *Adult 1*… to real names (column *For (person)*). Adjust the categories: income first, expenses below; type new ones in empty yellow rows. |
| **Recurring** | One row per thing that repeats: salaries, rent, school, subscriptions, insurance. Fill in the amount, how often and the date of one real payment. Delete the rows you don't have. |
| **Bills & Taxes** | Upcoming deadlines whose amount changes (tax payments, school trips). The bot reminds you 3 days before. |

Leave the monthly budgets empty for the first 2–3 months. Then the **Year** tab shows what you really spend on average, and you set budgets a little below that.

Don't rename the tabs or the column titles, and don't insert or delete rows in Settings: clear a name instead.

## Using it every day

Log a payment the moment you make it: it takes about 10 seconds. Example: tap **➖ Expense ▸ Groceries ▸ Today**, type `45.20`, tap **Skip**, pick who it's for and who paid, then **✅ Save**.

| Button | What it does |
| --- | --- |
| ➖ Expense / ➕ Income / 🧾 Tax | Add an entry. After you pick the category, the bot shows the last entries in it, so you can see if it's already there. |
| 🔎 Search | Find past entries: type a word like *dentist*, optionally with a month or year (*dentist 2026*). Results show a total per person. |
| 🕘 Recent | The last 10 entries, with who each one is for and who logged it. Filter by category or by person. |
| 📊 This month | Income, expenses and budget used so far this month. |
| 📅 Upcoming | Bills and recurring payments in the next 30 days. |
| ↩️ Undo last | Delete the entry you just added. |
| 📤 Export | Optional, off by default. Sends a month or a year as a CSV file to the chat. See below. |

- **Dates:** tap Today, Yesterday or the calendar. If you type one, write the day first: `25/10`.
- **Other currencies:** type the currency after the amount, for example `42000 ars`. Set the monthly rate in the *Exchange rates* tab.
- **Charts and analysis:** open the spreadsheet. *Dashboard* shows the year, *Analysis* ranks your categories for any month.

### Export from the phone (optional)

Off by default. To turn it on, open **Settings** and set **Allow export from Telegram** to **Yes**. Only people in **Allowed Telegram user IDs** (full access) then see a **📤 Export** button; read-only people never do. They pick a month or a year and the bot sends a CSV file that opens in Excel or Google Sheets. Set it back to **No** and the button disappears at the next /start.

A CSV file can be forwarded anywhere, so only turn this on if everyone with full access should be able to take the data out of the sheet.

## Updating to a new version

When a new version of `bot/FamilyBudgetBot.gs` is published in this repository, your data, settings and family list stay as they are. Only the code changes.

1. Spreadsheet ▸ **Extensions ▸ Apps Script**.
2. Select all the code, delete it, paste the new file, and save.
3. **Deploy ▸ Manage deployments**, click the **pencil**, set *Version* to **New version**, then **Deploy**. Don't use *New deployment*: that would change the web address.
4. Everyone sends `/start` once in Telegram to get any new buttons.

## If something goes wrong

| What you see | What to do |
| --- | --- |
| No **Extensions** menu | The file is still Excel (green "XLSX" label). **File ▸ Save as Google Sheets** and use the new copy. On a phone there is no Extensions menu: use a computer. |
| Apps Script opened as a separate "Untitled project" | It was opened from Drive. Close it and open it again from the spreadsheet: **Extensions ▸ Apps Script**. |
| No **Family Budget** menu after reloading | In Apps Script, pick **onOpen** in the function list next to *Run*, click **Run**, allow it, and go back to the spreadsheet. |
| The bot doesn't answer at all | **Family Budget ▸ Check bot status**. If you see "302" or pending messages: publish a new version (*Updating*, step 3), then run **1. Connect Telegram bot** again with the same token and URL. |
| The bot answers with ⚠️ and a message | Something in the sheet is off (often a renamed tab or column). The message says what. |
| *"The data you entered violates the data validation rules"* | **Family Budget ▸ Repair the dropdowns of Transactions**. |
| Months or numbers in the wrong language | **File ▸ Settings ▸ Locale**: pick your country. |
| A person can't get in | Check their number is in **Settings**, or send them the secret word again. |

Anything broken in the sheet can be brought back with **File ▸ Version history**.

## Privacy and safety in plain words

Your data stays in your own Google Drive; nothing is sent to anyone except your messages to and from Telegram.

- **Only your list gets in.** The bot answers only the Telegram IDs in Settings, only in private chats. To everyone else it stays silent.
- **The script can only open this one spreadsheet**, not the rest of your Drive.
- **Keep two things secret:** the BotFather token and the secret word. If either leaks, get a new token from BotFather (`/revoke`) and run **1. Connect Telegram bot** again, or change the secret word.
- **Read-only people** can look and search but can't add, change or delete anything.
- **No AI** reads your data: the bot is buttons and rules only.
- To share the spreadsheet itself with someone who should only look: **Share ▸ their email ▸ Viewer**.
