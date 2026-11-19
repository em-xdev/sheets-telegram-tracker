# Changelog

## 1.0.0 – first public version

**Spreadsheet (English and Spanish)**
- Dashboard, Analysis, Year and Search views calculated from one Transactions list.
- Year and month selectors; categories ranked from highest to lowest with share, budget used, previous period and same period last year.
- Income and expense categories, people ("For"), payers and payment methods editable in Settings, with spare rows.
- Recurring items with frequency, next date and monthly equivalent; Bills & Taxes with days left.
- Demo files with ten months of fake data.

**Telegram bot**
- Button menu: Expense, Tax, Income, Search, This month, Upcoming, Recent, Undo last.
- Last entries of a category shown while adding, to avoid duplicates; duplicate warning before saving.
- 🕘 Recent with filters by category and person, showing who each entry was for and who logged it.
- Search with totals per person.
- Several currencies with monthly exchange rates (`42000 ars`, `100 usd`), provisional-rate markers and automatic recalculation.
- Daily automation: recurring items posted on their date, bill reminders with a one-tap "Paid" button.
- Closed bot: only listed Telegram IDs, private chats only; full-access and read-only roles; secret word for new users with one-tap approval.
- Bot language per person (`/english`, `/espanol`); works with the English or the Spanish spreadsheet.
- Import of older monthly sheets (Category · Subcategory · Date · Detail · amounts in two currencies).
- Menu tools: repair dropdowns, recalculate conversions, check bot status.
- Optional 📤 Export: a month or a year as a CSV file sent to the chat – off by default (Settings ▸ "Allow export from Telegram"), full-access users only.
- Script limited to its own spreadsheet (`@OnlyCurrentDoc`); cells written by the bot are protected against formula injection.
