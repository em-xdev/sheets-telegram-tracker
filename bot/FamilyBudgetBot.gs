/** @OnlyCurrentDoc */
/******************************************************************************
 * FAMILY BUDGET BOT  –  Telegram ⇄ Google Sheets (Google Apps Script)
 * BOT DE PRESUPUESTO FAMILIAR – English / Español
 *
 * Menu-driven, no AI: you tap buttons (Expense / Tax / Income → category →
 * date → amount → description → who) and the bot writes a row into the
 * "Transactions" sheet. It also:
 *   • auto-adds the items in the "Recurring" sheet on their due date
 *   • reminds you about pending "Bills & Taxes" before they are due
 *   • searches past entries, shows "This month" and "Upcoming", undoes
 *   • speaks English or Spanish – each person chooses (/english, /espanol)
 *   • optional 📤 Export: a month or a year as a CSV file in the chat – off unless
 *     Settings ▸ "Allow export from Telegram" = Yes, and only for full-access users
 *   • works with the English workbook (Family_Budget) or the Spanish one
 *     (Presupuesto_Familiar) – it recognises the sheet and column names
 *
 * DATES: nobody has to type them – there are Today / Yesterday buttons and a
 * tap-to-pick calendar. If you do type one it is always read DAY FIRST
 * (25/10/2026), and the bot repeats it with the month name ("25 Oct 2026")
 * so a mix-up is obvious. Dates are written to the sheet in the sheet's own
 * time zone, so they never shift by a day.
 *
 * SETUP (once, ~15 minutes) – also on the "How to use" sheet:
 *   1. Upload the workbook to Google Drive → open it → File ▸ Save as Google
 *      Sheets. File ▸ Settings: Locale = your country (Spain / United Kingdom;
 *      avoid United States if you write dates day first) and your Time zone.
 *   2. In Telegram, talk to @BotFather → /newbot → copy the token.
 *   3. In the Google Sheet: Extensions ▸ Apps Script. Delete what's there,
 *      paste this whole file, Save. Project Settings (gear) ▸ same time zone.
 *   4. Deploy ▸ New deployment ▸ type "Web app".
 *        Execute as: Me   ·   Who has access: Anyone   → Deploy.
 *      Authorise when asked, then copy the Web app URL (ends with /exec).
 *   5. Reload the sheet. A "Family Budget" / "Presupuesto familiar" menu
 *      appears ▸ 1. Connect Telegram bot → paste the token, then the URL.
 *   6. Menu ▸ 2. Turn on daily automation (recurring + reminders).
 *   (optional) Menu ▸ 3. Set up currencies… – adds an "Exchange rates" sheet,
 *      Currency / Original amount / Subcategory columns and calendar pickers.
 *      Pesos: type "42000 ars" in the bot, or Currency = ARS + Original amount
 *      in the sheet – the € amount fills in by itself. Old data: put it in a tab
 *      called "Import" and use Menu ▸ Import transactions.
 *   7. Send /start to your bot. It replies with your Telegram ID: put it in
 *      Settings ▸ "Allowed Telegram user IDs" (comma-separate both parents).
 *      People who should only look go in "Read-only Telegram user IDs".
 *      Send /start again – the menu appears.
 *
 * "Anyone" access is required by Telegram, but the bot only accepts calls
 * that carry a secret key (added automatically in step 5) and only answers
 * the Telegram IDs you allow in Settings.
 *
 * If you change this code later: Deploy ▸ Manage deployments ▸ edit (pencil)
 * ▸ Version: New version ▸ Deploy. The URL stays the same.
 ******************************************************************************/

const STATE_TTL = 6 * 60 * 60;   // seconds a half-finished entry is remembered
const DUP_DAYS = 3;              // same category + amount this close in time = possible duplicate
const RECURRING_DUP_DAYS = 5;    // auto-post skips an item already logged by hand this close to its date
const SEARCH_MAX = 15;           // results listed per search

/* =========================================================================
 *  Spreadsheet vocabulary – the bot accepts the English OR Spanish workbook.
 *  Code uses the English names; this maps them to the Spanish workbook.
 * ========================================================================= */
const WB_ES = {
  // sheets
  'Transactions': 'Movimientos', 'Recurring': 'Recurrentes', 'Bills & Taxes': 'Facturas e impuestos', 'Settings': 'Configuración',
  // headers
  'Date': 'Fecha', 'Type': 'Tipo', 'Category': 'Categoría', 'Description': 'Descripción', 'Amount': 'Importe', 'For': 'Para',
  'Paid by': 'Pagado por', 'Payment method': 'Forma de pago', 'Tax deductible': 'Deducible', 'Notes': 'Notas', 'Source': 'Origen',
  'Name': 'Nombre', 'Frequency': 'Frecuencia', 'Every (months)': 'Cada (meses)', 'First date': 'Primera fecha',
  'End date (optional)': 'Fecha fin (opcional)', 'Active': 'Activo', 'Last auto-posted': 'Último registro automático',
  'Item': 'Concepto', 'Due date': 'Vencimiento', 'Expected amount': 'Importe previsto', 'Status': 'Estado', 'Reminded': 'Recordado',
  'Monthly budget / expected': 'Presupuesto / previsto mensual', 'For (person)': 'Para (persona)', 'Tax types': 'Tipos de impuesto',
  // values
  'Expense': 'Gasto', 'Income': 'Ingreso', 'Yes': 'Sí', 'No': 'No', 'Pending': 'Pendiente', 'Paid': 'Pagado',
  'Bot': 'Bot', 'Auto': 'Auto', 'Taxes': 'Impuestos',
  // setting labels
  'Allowed Telegram user IDs': 'IDs de Telegram con acceso total', 'Read-only Telegram user IDs': 'IDs de Telegram solo lectura',
  'Currency symbol': 'Símbolo de moneda', 'Remind bills this many days before': 'Avisar de facturas con estos días de antelación',
  'Auto-post recurring items': 'Registrar recurrentes automáticamente',
  'Allow export from Telegram': 'Permitir exportar desde Telegram',
  // currencies, subcategory, import
  'Exchange rates': 'Tipos de cambio', 'Currency': 'Moneda', 'Original amount': 'Importe original',
  'Subcategory': 'Subcategoría', 'Import': 'Importar', 'Month': 'Mes', 'Latest': 'Último',
  'Monthly equivalent': 'Equivalente mensual', 'Manual': 'Manual',
  'Show ID to unknown users': 'Mostrar ID a desconocidos'
};
// Notes markers written next to converted amounts (removed again when the real rate arrives)
const MARK = { en: { prov: '[provisional rate]', none: '[no exchange rate]' }, es: { prov: '[tipo provisional]', none: '[sin tipo de cambio]' } };
const WB_REV = {}; Object.keys(WB_ES).forEach(function (k) { WB_REV[WB_ES[k]] = k; });

/* =========================================================================
 *  Bot texts
 * ========================================================================= */
const TXT = {
  en: {
    menu_expense: '➖ Expense', menu_tax: '🧾 Tax', menu_income: '➕ Income', menu_search: '🔎 Search',
    menu_month: '📊 This month', menu_upcoming: '📅 Upcoming', menu_undo: '↩️ Undo last',
    private: '👋 This is a private family bot.\nYour Telegram ID is <code>{0}</code>.\nAsk the owner to add it in the spreadsheet – Settings ▸ "{1}" (can add and edit) or "{2}" (can only look) – then send /start again.',
    readonly_alert: '👀 You have read-only access.',
    hi: 'Hi {0}! What do you want to do?',
    help_full: '<b>{0}</b> – anything you paid\n<b>{1}</b> – a tax payment\n<b>{2}</b> – salary, benefits, other income\n<b>{3}</b> – find past entries\n<b>{4}</b> – totals so far\n<b>{5}</b> – bills and recurring items in the next 30 days\n<b>{6}</b> – remove your last entry\n<b>{7}</b> – last entries, by category or person',
    help_read: '👀 You have <b>read-only</b> access.\n\n<b>{0}</b> – find past entries\n<b>{1}</b> – totals so far\n<b>{2}</b> – what is due in the next 30 days\n<b>{3}</b> – last entries, by category or person',
    help_tip: 'Tip: just type a word (e.g. <i>dentist</i>) to search.',
    lang_hint: '🌐 Español: /espanol',
    lang_set: '✅ Language: English',
    cancelled: 'Cancelled.',
    search_prompt: '🔎 Type what to look for, e.g. <b>dentist</b>, <b>school</b>, <b>Coop</b>.\nAdd a year or month to narrow it: <b>dentist 2026</b>, <b>school 09/2026</b> or <b>school september</b>.',
    readonly_no_write: '👀 You have read-only access, so you can\'t add or delete entries.',
    tap_start: 'Tap a button below to start 👇',
    newcat_len: 'Please type a name between 2 and 40 characters (or /cancel).',
    newcat_added: '✅ New category added to Settings: <b>{0}</b>\nTip: set its monthly budget in Settings.',
    newcat_existing: 'Using the existing category <b>{0}</b>',
    newcat_wrongtype: '"{0}" already exists as a different type. Pick another name (or /cancel).',
    newcat_full: 'No free {0} rows left in Settings. Add a row there by hand (or /cancel).',
    newcat_prompt: '✏️ Type the name of the new {0} category:',
    w_expense: 'expense', w_income: 'income',
    date_bad: '🤔 I didn\'t get that date. Type it <b>day first</b>, like <b>25/10</b> or <b>25/10/2026</b> – or tap the calendar.',
    date_future: '📅 That date is in the future. Pick a date up to today (or /cancel).',
    amount_bad: '🤔 Please type just the amount, e.g. <b>45.20</b> or <b>1250</b>.',
    use_buttons: 'Please use the buttons above, or /cancel to start again.',
    expired: '⌛ This menu has expired. Tap a button below to start again.',
    title_expense: '➖ <b>New expense</b>', title_income: '➕ <b>New income</b>', title_tax: '🧾 <b>New tax payment</b>',
    choose_cat: 'Choose a category:', which_tax: 'Which tax?',
    category_line: 'Category: <b>{0}</b>', tax_line: 'Tax: <b>{0}</b>', new_cat_btn: '➕ New category…',
    when: '📅 When?\n<i>(or type it, day first: 25/10)</i>',
    today: 'Today', yesterday: 'Yesterday', two_days: '2 days ago', calendar: '📅 Calendar', back: '↩ Back',
    pick_day: '📅 Pick the day:', date_line: '📅 Date: <b>{0}</b>',
    how_much: '💶 How much? Type the amount (e.g. <b>45.20</b>)',
    desc_prompt: '📝 Short description (e.g. <i>supermarket</i>) or tap Skip',
    desc_prompt_tax: '📝 Add a note (e.g. <i>2nd instalment 2026</i>) or tap Skip',
    skip: 'Skip', desc_line: '📝 Description: <i>{0}</i>',
    who_for: '👤 Who is it for?', whose_income: '👤 Whose income?', who_paid: '💳 Who paid?',
    for_line: '👤 For: <b>{0}</b>', paid_line: '💳 Paid by: <b>{0}</b>',
    check_save: '<b>Check and save</b>', kind_expense: '➖ Expense', kind_income: '➕ Income', kind_tax: '🧾 Tax',
    dup_title: '⚠️ <b>Possible duplicate</b> – already in the sheet:',
    save: '✅ Save', save_anyway: '✅ Save anyway', cancel_btn: '❌ Cancel',
    saved: '✅ <b>Saved</b> – {0}, {1}, {2}',
    progress: '{0} in {1}: <b>{2}</b>', progress_budget: ' of {0} ({1}%)', over_budget: ' 🔴 over budget',
    progress_expected: ' of {0} expected',
    undo_nothing: 'Nothing to undo – I only remember your last entry from this week.',
    undo_ask: 'Delete this entry?', yes_delete: '🗑 Yes, delete', keep: 'Keep it', kept: 'OK, kept.', deleted: '🗑 Deleted.',
    undo_gone: '⚠️ That row is no longer there – delete it by hand if needed.',
    undo_changed: '⚠️ The sheet changed since then, so I didn\'t delete anything. Please remove it by hand.',
    so_far: '(so far)', income: 'Income', expenses: 'Expenses', net: 'Net',
    budget_used: '🎯 Budget used: <b>{0}%</b> of {1}', top: 'Top spending',
    nothing_due: '📅 Nothing due in the next 30 days.', next30: '📅 <b>Next 30 days</b>', in_suffix: ' (in)',
    rec_title: '🔁 <b>Recurring items</b>', rec_added: 'Added:', rec_skipped: 'Skipped – you had already logged them:',
    overdue: '⚠️ overdue {0} days', due_today: '❗ due TODAY', due_in1: 'due tomorrow', due_in: 'due in {0} days',
    expected: 'Expected: {0}', paid_btn: '✅ Paid {0} – log it',
    bill_changed: '⚠️ That bill changed in the sheet (or is already paid). Nothing logged.',
    bill_paid: '✅ <b>{0}</b> marked as paid and logged ({1}, today).\nIf the amount was different, fix it in the sheet.',
    search_empty: '🔎 Type a word to search, e.g. <b>dentist</b>.', nothing_found: '🔎 Nothing found for <b>{0}</b>.',
    entry: 'entry', entries: 'entries', spent: '➖ Spent: <b>{0}</b>', received: '➕ Received: <b>{0}</b>',
    more: '…and {0} older. Add a month or year to narrow it.', everything: 'everything',
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    mon: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    dows: ['Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su'],
    cmd_start: 'Show the menu', cmd_cancel: 'Cancel the current entry', cmd_lang: 'Cambiar a español',
    how_much_fx: '💶 How much? Type the amount (e.g. <b>45.20</b>)\nOther currency: add it after the number, e.g. <b>42000 ars</b>',
    no_rate: '⚠️ There is no {0} exchange rate yet. Add it in the "{1}" sheet (how many {0} = 1 €) and type the amount again.',
    unknown_cur: '🤔 I only know these currencies: {0}. Add a column in "{1}" for others.',
    conv: '≈ <b>{0}</b> (1 € = {1} {2} · {3}{4})', provisional: ', provisional rate',
    rec_norate: 'Not added – no exchange rate yet:',
    id_reply: '🔑 Your Telegram ID is <code>{0}</code>.\nThe family admin has been told and can add you.',
    id_own: '🔑 Your Telegram ID is <code>{0}</code>.',
    id_admin: '👤 <b>{0}</b>{1} asked to join the family bot.\nID: <code>{2}</code>',
    add_full: '➕ Full access', add_read: '👀 Read-only', add_ignore: 'Ignore',
    added: '✅ {0} added – {1}. They can send /start now.', ignored: 'Ignored.',
    full_access: 'full access', read_access: 'read-only',
    welcome: '✅ You have been added to the family bot. Send /start to begin.',
    menu_recent: '🕘 Recent',
    menu_export: '📤 Export',
    export_pick: '📤 <b>Export to a file</b>\nWhich period? You get a CSV file that opens in Excel or Google Sheets.',
    export_year: 'All {0}', export_none: '📤 No entries in {0}.', export_sent: '📤 {0}: file sent ({1} entries).',
    export_caption: '📤 <b>{0}</b> · {1} entries\nSpent {2} · Received {3}', export_file: 'budget',
    export_off: '📤 Export from Telegram is turned off. Whoever manages the spreadsheet can turn it on in {1} ▸ "{0}".',
    export_fail: '⚠️ The file could not be sent: {0}',
    help_export: '<b>{0}</b> – a month or a year as a file (CSV)',
    recent_title: '🕘 <b>Last {0} entries</b>{1}', recent_none: '🕘 Nothing here yet.',
    recent_all: '🕘 All', recent_cat: '📂 By category', recent_who: '👤 By person',
    recent_pick_cat: '📂 Which category?', recent_pick_who: '👤 Whose entries?',
    last_in_cat: '🕘 <b>Last in {0}</b> – check it is not already there:', none_in_cat: '🕘 Nothing in {0} yet.',
    by_person: '<b>By person</b>', nobody: '(nobody)'
  },
  es: {
    menu_expense: '➖ Gasto', menu_tax: '🧾 Impuesto', menu_income: '➕ Ingreso', menu_search: '🔎 Buscar',
    menu_month: '📊 Este mes', menu_upcoming: '📅 Próximos', menu_undo: '↩️ Deshacer',
    private: '👋 Este es un bot familiar privado.\nTu ID de Telegram es <code>{0}</code>.\nPide que lo añadan en la hoja – Configuración ▸ "{1}" (puede añadir y editar) o "{2}" (solo consultar) – y vuelve a enviar /start.',
    readonly_alert: '👀 Tienes acceso de solo lectura.',
    hi: '¡Hola {0}! ¿Qué quieres hacer?',
    help_full: '<b>{0}</b> – cualquier cosa que hayas pagado\n<b>{1}</b> – un pago de impuestos\n<b>{2}</b> – nómina, ayudas, otros ingresos\n<b>{3}</b> – buscar movimientos\n<b>{4}</b> – totales hasta hoy\n<b>{5}</b> – facturas y recurrentes de los próximos 30 días\n<b>{6}</b> – borrar tu último movimiento\n<b>{7}</b> – últimos movimientos, por categoría o persona',
    help_read: '👀 Tienes acceso de <b>solo lectura</b>.\n\n<b>{0}</b> – buscar movimientos\n<b>{1}</b> – totales hasta hoy\n<b>{2}</b> – lo que vence en los próximos 30 días\n<b>{3}</b> – últimos movimientos, por categoría o persona',
    help_tip: 'Truco: escribe una palabra (p. ej. <i>dentista</i>) para buscar.',
    lang_hint: '🌐 English: /english',
    lang_set: '✅ Idioma: español',
    cancelled: 'Cancelado.',
    search_prompt: '🔎 Escribe lo que buscas, p. ej. <b>dentista</b>, <b>colegio</b>, <b>Mercadona</b>.\nAñade un año o un mes para acotar: <b>dentista 2026</b>, <b>colegio 09/2026</b> o <b>colegio septiembre</b>.',
    readonly_no_write: '👀 Tienes acceso de solo lectura: no puedes añadir ni borrar movimientos.',
    tap_start: 'Toca un botón de abajo para empezar 👇',
    newcat_len: 'Escribe un nombre de 2 a 40 caracteres (o /cancel).',
    newcat_added: '✅ Nueva categoría añadida en Configuración: <b>{0}</b>\nTruco: pon su presupuesto mensual en Configuración.',
    newcat_existing: 'Uso la categoría que ya existe: <b>{0}</b>',
    newcat_wrongtype: '"{0}" ya existe con otro tipo. Elige otro nombre (o /cancel).',
    newcat_full: 'No quedan filas libres de {0} en Configuración. Añade una fila a mano (o /cancel).',
    newcat_prompt: '✏️ Escribe el nombre de la nueva categoría de {0}:',
    w_expense: 'gasto', w_income: 'ingreso',
    date_bad: '🤔 No entiendo esa fecha. Escríbela con el <b>día primero</b>, como <b>25/10</b> o <b>25/10/2026</b>, o usa el calendario.',
    date_future: '📅 Esa fecha es futura. Elige una fecha hasta hoy (o /cancel).',
    amount_bad: '🤔 Escribe solo el importe, p. ej. <b>45,20</b> o <b>1250</b>.',
    use_buttons: 'Usa los botones de arriba, o /cancel para empezar de nuevo.',
    expired: '⌛ Este menú ha caducado. Toca un botón de abajo para empezar de nuevo.',
    title_expense: '➖ <b>Nuevo gasto</b>', title_income: '➕ <b>Nuevo ingreso</b>', title_tax: '🧾 <b>Nuevo pago de impuesto</b>',
    choose_cat: 'Elige una categoría:', which_tax: '¿Qué impuesto?',
    category_line: 'Categoría: <b>{0}</b>', tax_line: 'Impuesto: <b>{0}</b>', new_cat_btn: '➕ Nueva categoría…',
    when: '📅 ¿Cuándo?\n<i>(o escríbela, día primero: 25/10)</i>',
    today: 'Hoy', yesterday: 'Ayer', two_days: 'Hace 2 días', calendar: '📅 Calendario', back: '↩ Volver',
    pick_day: '📅 Elige el día:', date_line: '📅 Fecha: <b>{0}</b>',
    how_much: '💶 ¿Cuánto? Escribe el importe (p. ej. <b>45,20</b>)',
    desc_prompt: '📝 Descripción corta (p. ej. <i>supermercado</i>) o toca Omitir',
    desc_prompt_tax: '📝 Añade una nota (p. ej. <i>2.º plazo 2026</i>) o toca Omitir',
    skip: 'Omitir', desc_line: '📝 Descripción: <i>{0}</i>',
    who_for: '👤 ¿Para quién es?', whose_income: '👤 ¿De quién es el ingreso?', who_paid: '💳 ¿Quién pagó?',
    for_line: '👤 Para: <b>{0}</b>', paid_line: '💳 Pagado por: <b>{0}</b>',
    check_save: '<b>Revisa y guarda</b>', kind_expense: '➖ Gasto', kind_income: '➕ Ingreso', kind_tax: '🧾 Impuesto',
    dup_title: '⚠️ <b>Posible duplicado</b> – ya está en la hoja:',
    save: '✅ Guardar', save_anyway: '✅ Guardar igualmente', cancel_btn: '❌ Cancelar',
    saved: '✅ <b>Guardado</b> – {0}, {1}, {2}',
    progress: '{0} en {1}: <b>{2}</b>', progress_budget: ' de {0} ({1} %)', over_budget: ' 🔴 por encima del presupuesto',
    progress_expected: ' de {0} previstos',
    undo_nothing: 'No hay nada que deshacer: solo recuerdo tu último movimiento de esta semana.',
    undo_ask: '¿Borrar este movimiento?', yes_delete: '🗑 Sí, borrar', keep: 'Conservar', kept: 'Vale, se conserva.', deleted: '🗑 Borrado.',
    undo_gone: '⚠️ Esa fila ya no está – bórrala a mano si hace falta.',
    undo_changed: '⚠️ La hoja ha cambiado desde entonces, así que no he borrado nada. Bórralo a mano.',
    so_far: '(hasta hoy)', income: 'Ingresos', expenses: 'Gastos', net: 'Neto',
    budget_used: '🎯 Presupuesto usado: <b>{0} %</b> de {1}', top: 'Mayores gastos',
    nothing_due: '📅 Nada pendiente en los próximos 30 días.', next30: '📅 <b>Próximos 30 días</b>', in_suffix: ' (ingreso)',
    rec_title: '🔁 <b>Movimientos recurrentes</b>', rec_added: 'Añadidos:', rec_skipped: 'Omitidos – ya los habías registrado:',
    overdue: '⚠️ vencida hace {0} días', due_today: '❗ vence HOY', due_in1: 'vence mañana', due_in: 'vence en {0} días',
    expected: 'Previsto: {0}', paid_btn: '✅ Pagado {0} – registrar',
    bill_changed: '⚠️ Esa factura ha cambiado en la hoja (o ya está pagada). No he registrado nada.',
    bill_paid: '✅ <b>{0}</b> marcada como pagada y registrada ({1}, hoy).\nSi el importe fue distinto, corrígelo en la hoja.',
    search_empty: '🔎 Escribe una palabra para buscar, p. ej. <b>dentista</b>.', nothing_found: '🔎 No hay resultados para <b>{0}</b>.',
    entry: 'movimiento', entries: 'movimientos', spent: '➖ Gastado: <b>{0}</b>', received: '➕ Recibido: <b>{0}</b>',
    more: '…y {0} más antiguos. Añade un mes o un año para acotar.', everything: 'todo',
    months: ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'],
    mon: ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sept', 'oct', 'nov', 'dic'],
    dows: ['L', 'M', 'X', 'J', 'V', 'S', 'D'],
    cmd_start: 'Mostrar el menú', cmd_cancel: 'Cancelar el movimiento en curso', cmd_lang: 'Switch to English',
    how_much_fx: '💶 ¿Cuánto? Escribe el importe (p. ej. <b>45,20</b>)\nOtra moneda: ponla detrás del número, p. ej. <b>42000 ars</b>',
    no_rate: '⚠️ Todavía no hay tipo de cambio para {0}. Añádelo en la hoja "{1}" (cuántos {0} = 1 €) y vuelve a escribir el importe.',
    unknown_cur: '🤔 Solo conozco estas monedas: {0}. Para otras, añade una columna en "{1}".',
    conv: '≈ <b>{0}</b> (1 € = {1} {2} · {3}{4})', provisional: ', tipo provisional',
    rec_norate: 'No añadidos – falta el tipo de cambio:',
    id_reply: '🔑 Tu ID de Telegram es <code>{0}</code>.\nYa se ha avisado al administrador de la familia para que te añada.',
    id_own: '🔑 Tu ID de Telegram es <code>{0}</code>.',
    id_admin: '👤 <b>{0}</b>{1} quiere entrar en el bot familiar.\nID: <code>{2}</code>',
    add_full: '➕ Acceso total', add_read: '👀 Solo lectura', add_ignore: 'Ignorar',
    added: '✅ {0} añadido – {1}. Ya puede enviar /start.', ignored: 'Ignorado.',
    full_access: 'acceso total', read_access: 'solo lectura',
    welcome: '✅ Te han añadido al bot familiar. Envía /start para empezar.',
    menu_recent: '🕘 Últimos',
    menu_export: '📤 Exportar',
    export_pick: '📤 <b>Exportar a un archivo</b>\n¿Qué periodo? Recibes un archivo CSV que se abre en Excel o en Hojas de cálculo de Google.',
    export_year: 'Todo {0}', export_none: '📤 No hay movimientos en {0}.', export_sent: '📤 {0}: archivo enviado ({1} movimientos).',
    export_caption: '📤 <b>{0}</b> · {1} movimientos\nGastado {2} · Recibido {3}', export_file: 'presupuesto',
    export_off: '📤 Exportar desde Telegram está desactivado. Quien gestiona la hoja puede activarlo en {1} ▸ "{0}".',
    export_fail: '⚠️ No se pudo enviar el archivo: {0}',
    help_export: '<b>{0}</b> – un mes o un año como archivo (CSV)',
    recent_title: '🕘 <b>Últimos {0} movimientos</b>{1}', recent_none: '🕘 Todavía no hay nada aquí.',
    recent_all: '🕘 Todos', recent_cat: '📂 Por categoría', recent_who: '👤 Por persona',
    recent_pick_cat: '📂 ¿Qué categoría?', recent_pick_who: '👤 ¿De quién?',
    last_in_cat: '🕘 <b>Últimos en {0}</b> – comprueba que no esté ya:', none_in_cat: '🕘 Todavía no hay nada en {0}.',
    by_person: '<b>Por persona</b>', nobody: '(sin persona)'
  }
};
// Spreadsheet menu & pop-ups (in the workbook's language)
const UI = {
  en: {
    menu: 'Family Budget', connect: '1. Connect Telegram bot…', auto: '2. Turn on daily automation',
    post_now: 'Add due recurring items now', remind_now: 'Send bill reminders now', status: 'Check bot status',
    p1t: 'Connect Telegram bot (1/2)', p1: 'Paste the bot token from @BotFather:',
    p2t: 'Connect Telegram bot (2/2)', p2: 'Paste the Web app URL (Deploy ▸ Manage deployments, it ends with /exec):',
    bad_url: 'That URL does not look right. It should start with https://script.google.com/ and end with /exec',
    bad_token: 'Telegram rejected the token: ', connected: 'Connected to @{0} ✅\n\nNow send /start to the bot in Telegram. It will tell you your ID – add it in Settings ▸ "Allowed Telegram user IDs".',
    hook_err: 'Webhook error: ', auto_done: 'Done ✅ Every morning around 7:00 the bot will add due recurring items and send bill reminders.',
    posted: '{0} recurring item(s) added to Transactions.', nothing_posted: 'Nothing due – everything is up to date.',
    reminded: 'Sent a reminder with {0} item(s).', nothing_remind: 'No pending bills in the reminder window.',
    st_hook: 'Webhook set: ', yes: 'yes', no: 'NO', st_pending: 'Pending updates: ', st_note: 'Last Telegram note: ',
    st_full: 'Full-access IDs: ', st_read: 'Read-only IDs: ', none: '(none)', st_302: '(If you see "302": update the deployment to a new version, then run 1. Connect Telegram bot again.)',
    fx_setup: '3. Set up currencies, subcategories & date pickers', fx_recalc: 'Recalculate currency conversions',
    import: 'Import old data (extra tabs with Category + Date)',
    fx_done: 'Done ✅\n\n• New sheet "{0}": type how many pesos (ARS) or dollars (USD) 1 € is worth, one row per month.\n• "{1}" now has the columns Currency, Original amount and Subcategory. For a payment in pesos: Currency = ARS, write the pesos in Original amount – the € Amount fills in by itself.\n• Recurring has a Currency column too.\n• Date columns now open a calendar when you double-click them.',
    recalc_done: '{0} row(s) converted to €. Still waiting for an exchange rate: {1}.',
    import_none: 'First add a tab called "{0}" (File ▸ Import ▸ Upload ▸ "Insert new sheet(s)"), with the same column titles as "{1}".',
    import_unknown: 'These categories are not in Settings yet:\n\n{0}\n\nAdd them there (expense or income rows), then run the import again. Nothing was imported.',
    import_done: 'Imported {0} row(s) ✅\nSkipped – already in the sheet: {1}\nSkipped – no amount or date: {2}{3}\n\nYou can delete these tabs now: "{4}".',
    import_norate: '\n\n{0} row(s) are waiting for an exchange rate – they fill in by themselves when you add it in "{1}".',
    repair: 'Repair the dropdowns of Transactions',
    repair_done: 'Done ✅ The dropdowns of "{0}" are reset. They now only warn (red corner) instead of blocking.',
    import_failed: 'The import stopped while writing: {0}\n\nRun Menu ▸ Repair the dropdowns of Transactions, then import again (nothing was written).',
    import_approx: '\n{0} row(s) had no date: put on the 1st of their month, noted "approximate date".',
    secret: '4. Secret word for new users…',
    secret_p: 'Type a secret word (no spaces, e.g. "mango42"). A new person sends it to the bot: they get their ID and you get a message with buttons to add them. Everyone else gets silence.\n\nNow: {0}\nLeave empty and press OK to turn it off.',
    secret_on: 'Saved ✅\nA new person sends  {0}  to the bot.', secret_off: 'Turned off – the bot shows the ID to nobody.',
    secret_bad: 'Use one word without spaces (at least 4 characters).', secret_none: '(not set)',
    import_unknown_q: 'These categories are not in Settings yet:\n\n{0}\n\nAdd them now as EXPENSE categories?\n(If one of them is income, choose No, add it in the green income rows of Settings, and import again.)'
  },
  es: {
    menu: 'Presupuesto familiar', connect: '1. Conectar bot de Telegram…', auto: '2. Activar automatización diaria',
    post_now: 'Añadir ahora los recurrentes pendientes', remind_now: 'Enviar ahora los avisos de facturas', status: 'Comprobar estado del bot',
    p1t: 'Conectar bot de Telegram (1/2)', p1: 'Pega el token del bot que te dio @BotFather:',
    p2t: 'Conectar bot de Telegram (2/2)', p2: 'Pega la URL de la aplicación web (Implementar ▸ Gestionar implementaciones, termina en /exec):',
    bad_url: 'Esa URL no parece correcta. Debe empezar por https://script.google.com/ y terminar en /exec',
    bad_token: 'Telegram ha rechazado el token: ', connected: 'Conectado a @{0} ✅\n\nAhora envía /start al bot en Telegram. Te dirá tu ID: añádelo en Configuración ▸ "IDs de Telegram con acceso total".',
    hook_err: 'Error del webhook: ', auto_done: 'Hecho ✅ Cada mañana hacia las 7:00 el bot añadirá los recurrentes y enviará los avisos de facturas.',
    posted: '{0} recurrente(s) añadido(s) a Movimientos.', nothing_posted: 'No hay nada pendiente: todo está al día.',
    reminded: 'Aviso enviado con {0} factura(s).', nothing_remind: 'No hay facturas pendientes dentro del plazo de aviso.',
    st_hook: 'Webhook configurado: ', yes: 'sí', no: 'NO', st_pending: 'Mensajes pendientes: ', st_note: 'Última nota de Telegram: ',
    st_full: 'IDs con acceso total: ', st_read: 'IDs solo lectura: ', none: '(ninguno)', st_302: '(Si ves "302": actualiza la implementación a una versión nueva y vuelve a ejecutar 1. Conectar bot de Telegram.)',
    fx_setup: '3. Preparar monedas, subcategorías y calendarios', fx_recalc: 'Recalcular conversiones de moneda',
    import: 'Importar datos antiguos (pestañas con Categoría + Fecha)',
    fx_done: 'Hecho ✅\n\n• Nueva hoja "{0}": escribe cuántos pesos (ARS) o dólares (USD) vale 1 €, una fila por mes.\n• "{1}" tiene ahora las columnas Moneda, Importe original y Subcategoría. Para un pago en pesos: Moneda = ARS y los pesos en Importe original – el Importe en € se rellena solo.\n• Recurrentes también tiene columna Moneda.\n• Las columnas de fecha abren un calendario al hacer doble clic.',
    recalc_done: '{0} fila(s) convertidas a €. Pendientes de tipo de cambio: {1}.',
    import_none: 'Primero añade una pestaña llamada "{0}" (Archivo ▸ Importar ▸ Subir ▸ "Insertar hojas nuevas"), con los mismos títulos de columna que "{1}".',
    import_unknown: 'Estas categorías todavía no están en Configuración:\n\n{0}\n\nAñádelas allí (filas de gastos o de ingresos) y vuelve a importar. No se ha importado nada.',
    import_done: 'Importadas {0} fila(s) ✅\nOmitidas – ya estaban en la hoja: {1}\nOmitidas – sin importe o sin fecha: {2}{3}\n\nYa puedes borrar estas pestañas: "{4}".',
    import_norate: '\n\n{0} fila(s) esperan el tipo de cambio – se rellenan solas cuando lo añadas en "{1}".',
    repair: 'Reparar los desplegables de Movimientos',
    repair_done: 'Hecho ✅ Los desplegables de "{0}" están reparados. Ahora solo avisan (esquina roja) en lugar de bloquear.',
    import_failed: 'La importación se paró al escribir: {0}\n\nUsa Menú ▸ Reparar los desplegables de Movimientos y vuelve a importar (no se ha escrito nada).',
    import_approx: '\n{0} fila(s) no tenían fecha: puestas el día 1 de su mes, con la nota "fecha aproximada".',
    secret: '4. Palabra secreta para nuevos usuarios…',
    secret_p: 'Escribe una palabra secreta (sin espacios, p. ej. "mango42"). Una persona nueva se la envía al bot: recibe su ID y a ti te llega un mensaje con botones para añadirla. Al resto, silencio.\n\nAhora: {0}\nDéjala vacía y pulsa Aceptar para desactivarla.',
    secret_on: 'Guardada ✅\nUna persona nueva envía  {0}  al bot.', secret_off: 'Desactivada – el bot no muestra el ID a nadie.',
    secret_bad: 'Usa una sola palabra sin espacios (mínimo 4 caracteres).', secret_none: '(sin configurar)',
    import_unknown_q: 'Estas categorías todavía no están en Configuración:\n\n{0}\n\n¿Las añado ahora como categorías de GASTO?\n(Si alguna es de ingresos, elige No, añádela en las filas verdes de ingresos de Configuración y vuelve a importar.)'
  }
};

let LANG_ = 'en';   // language of the person we are talking to right now
function t_(key) {
  let s = (TXT[LANG_] && TXT[LANG_][key] !== undefined) ? TXT[LANG_][key] : TXT.en[key];
  for (let i = 1; i < arguments.length; i++) s = String(s).split('{' + (i - 1) + '}').join(arguments[i]);
  return s;
}
function ui_(key) {
  let s = UI[wbLang_()][key];
  for (let i = 1; i < arguments.length; i++) s = String(s).split('{' + (i - 1) + '}').join(arguments[i]);
  return s;
}
function mainKeyboard_(role) {
  const k = role === 'full'
    ? [[t_('menu_expense'), t_('menu_tax')], [t_('menu_income'), t_('menu_search')], [t_('menu_month'), t_('menu_upcoming')], [t_('menu_recent'), t_('menu_undo')]]
    : [[t_('menu_search'), t_('menu_recent')], [t_('menu_month'), t_('menu_upcoming')]];
  if (exportAllowed_(role)) k.push([t_('menu_export')]);
  return { reply_markup: { keyboard: k, resize_keyboard: true, is_persistent: true } };
}
/** Which menu button was pressed – in either language (old keyboards keep working after a language switch). */
function menuAction_(text) {
  const keys = ['expense', 'tax', 'income', 'search', 'month', 'upcoming', 'undo', 'recent', 'export'];
  for (let i = 0; i < keys.length; i++) {
    if (text === TXT.en['menu_' + keys[i]] || text === TXT.es['menu_' + keys[i]]) return keys[i];
  }
  return null;
}

/* ---------- per-person language ---------- */
function userLang_(userId, telegramCode) {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('lang_' + userId);
  if (saved) return saved;
  const guess = telegramCode ? (/^es/i.test(telegramCode) ? 'es' : 'en') : wbLang_();
  props.setProperty('lang_' + userId, guess);
  return guess;
}
function setUserLang_(userId, lang) { PropertiesService.getScriptProperties().setProperty('lang_' + userId, lang); }

/* =========================================================================
 *  Spreadsheet menu & setup
 * ========================================================================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu(ui_('menu'))
    .addItem(ui_('connect'), 'setupBot')
    .addItem(ui_('auto'), 'installDailyTrigger')
    .addItem(ui_('fx_setup'), 'setupCurrencies')
    .addItem(ui_('secret'), 'setSecretWord')
    .addSeparator()
    .addItem(ui_('import'), 'importTransactions')
    .addItem(ui_('fx_recalc'), 'recalcConversionsNow')
    .addItem(ui_('repair'), 'repairValidationNow')
    .addSeparator()
    .addItem(ui_('post_now'), 'postRecurringNow')
    .addItem(ui_('remind_now'), 'sendRemindersNow')
    .addItem(ui_('status'), 'botStatus')
    .addToUi();
}

function setupBot() {
  const ui = SpreadsheetApp.getUi();
  const t = ui.prompt(ui_('p1t'), ui_('p1'), ui.ButtonSet.OK_CANCEL);
  if (t.getSelectedButton() !== ui.Button.OK) return;
  const token = t.getResponseText().trim();
  const u = ui.prompt(ui_('p2t'), ui_('p2'), ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  const url = u.getResponseText().trim();
  if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) { ui.alert(ui_('bad_url')); return; }
  const props = PropertiesService.getScriptProperties();
  const key = Utilities.getUuid().replace(/-/g, '');
  props.setProperties({ BOT_TOKEN: token, WEBHOOK_KEY: key });
  const me = tg_('getMe', {});
  if (!me.ok) { ui.alert(ui_('bad_token') + (me.description || '')); return; }
  const hook = tg_('setWebhook', { url: url + '?key=' + key, drop_pending_updates: true, allowed_updates: ['message', 'callback_query'] });
  ['en', 'es'].forEach(function (lang) {
    LANG_ = lang;
    const cmds = [{ command: 'start', description: t_('cmd_start') }, { command: 'cancel', description: t_('cmd_cancel') },
      { command: lang === 'en' ? 'espanol' : 'english', description: t_('cmd_lang') }];
    tg_('setMyCommands', lang === 'en' ? { commands: cmds } : { commands: cmds, language_code: 'es' });
  });
  ui.alert(hook.ok ? ui_('connected', me.result.username) : ui_('hook_err') + hook.description);
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (tr) { return tr.getHandlerFunction() === 'dailyJob'; })
    .forEach(function (tr) { ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert(ui_('auto_done'));
}

function botStatus() {
  const r = tg_('getWebhookInfo', {}).result || {};
  SpreadsheetApp.getUi().alert(
    ui_('st_hook') + (r.url ? ui_('yes') : ui_('no')) + '\n' + ui_('st_pending') + (r.pending_update_count || 0) +
    (r.last_error_message ? '\n' + ui_('st_note') + r.last_error_message : '') +
    '\n' + ui_('st_full') + (idsFor_('Allowed Telegram user IDs').join(', ') || ui_('none')) +
    '\n' + ui_('st_read') + (idsFor_('Read-only Telegram user IDs').join(', ') || ui_('none')) + '\n\n' + ui_('st_302'));
}

function postRecurringNow() {
  const n = postRecurring_();
  SpreadsheetApp.getUi().alert(n ? ui_('posted', n) : ui_('nothing_posted'));
}
function sendRemindersNow() {
  const n = sendReminders_(true);
  SpreadsheetApp.getUi().alert(n ? ui_('reminded', n) : ui_('nothing_remind'));
}

/* =========================================================================
 *  Currencies (pesos, dollars…), subcategory, date pickers, import
 *  • "Exchange rates / Tipos de cambio" sheet: one row per month, one column per currency
 *    ("ARS per 1 €"). Amount (€) = Original amount ÷ rate of that month.
 *  • Missing month → the latest earlier rate is used and the row is marked "provisional";
 *    when the real rate is typed, every row recalculates by itself (onEdit).
 * ========================================================================= */
function setupCurrencies() {
  const ss = ss_(), es = wbLang_() === 'es';
  // 1. Exchange-rate sheet
  let rs = sheetOpt_('Exchange rates');
  if (!rs) {
    rs = ss.insertSheet(val_('Exchange rates'));
    const per = es ? ' por 1 €' : ' per 1 €';
    rs.getRange(1, 1, 1, 4).setValues([[val_('Month'), 'ARS' + per, 'USD' + per, val_('Notes')]])
      .setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#FFFFFF');
    rs.getRange(2, 1, 1, 3).setValues([[val_('Latest'),
      '=IFERROR(INDEX(FILTER(B4:B,B4:B<>""),COUNT(B4:B)),"")', '=IFERROR(INDEX(FILTER(C4:C,C4:C<>""),COUNT(C4:C)),"")']])
      .setFontWeight('bold');
    rs.getRange(3, 1).setValue(es
      ? 'Escribe cuántos pesos / dólares vale 1 € cada mes. Si falta un mes se usa el último conocido (provisional). Para otra moneda, añade una columna con su código (p. ej. "BRL por 1 €").'
      : 'Type how many pesos / dollars 1 € is worth each month. A missing month uses the latest known rate (provisional). For another currency add a column with its code (e.g. "BRL per 1 €").')
      .setFontStyle('italic').setFontColor('#7F7F7F');
    const months = [];
    for (let y = 2024; y <= 2030; y++) for (let m = 1; m <= 12; m++) months.push([toDate_(fromParts_(y, m, 1))]);
    rs.getRange(4, 1, months.length, 1).setValues(months).setNumberFormat('mmm yyyy');
    rs.getRange(2, 2, months.length + 2, 2).setNumberFormat('#,##0.00');
    rs.getRange(4, 2, months.length, 2).setBackground('#FFF4C2');
    rs.setFrozenRows(3); rs.setColumnWidth(1, 110); rs.setColumnWidth(2, 120); rs.setColumnWidth(3, 120); rs.setColumnWidth(4, 260);
  }
  RATES_MEMO_ = null;
  const curList = currencies_();
  const curRule = SpreadsheetApp.newDataValidation().requireValueInList(curList, true).setAllowInvalid(false).build();
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();

  // 2. Transactions: Currency, Original amount, Subcategory + calendar on Date
  const tx = sheet_('Transactions');
  addHeaders_(tx, 1, ['Currency', 'Original amount', 'Subcategory']);
  const ht = headerMap_(tx, 1), n = Math.max(tx.getMaxRows() - 1, 1);
  tx.getRange(2, ht['Currency'] + 1, n, 1).setDataValidation(curRule);
  tx.getRange(2, ht['Original amount'] + 1, n, 1).setNumberFormat('#,##0.00');
  tx.getRange(2, ht['Date'] + 1, n, 1).setDataValidation(dateRule);

  // 3. Recurring: Currency column, € monthly equivalent, calendars
  const rc = sheet_('Recurring');
  addHeaders_(rc, 4, ['Currency']);
  const hr = headerMap_(rc, 4), last = recurringLastRow_(rc);
  if (last >= 5) {
    const rows = last - 4;
    rc.getRange(5, hr['Currency'] + 1, rows, 1).setDataValidation(curRule);
    ['First date', 'End date (optional)'].forEach(function (k) { if (hr[k] !== undefined) rc.getRange(5, hr[k] + 1, rows, 1).setDataValidation(dateRule); });
    if (hr['Monthly equivalent'] !== undefined && hr['Amount'] !== undefined && hr['Every (months)'] !== undefined) {
      const A = colL_(hr['Amount'] + 1), F = colL_(hr['Every (months)'] + 1), C = colL_(hr['Currency'] + 1);
      const RS = "'" + rs.getName() + "'";
      const f = [];
      for (let r = 5; r <= last; r++) {
        f.push(['=IF(OR(' + A + r + '="",' + F + r + '=""),"",' + A + r + '/' + F + r + '/IF(OR(' + C + r + '="",' + C + r + '="EUR"),1,' +
          'IFERROR(INDEX(' + RS + '!$B$2:$Z$2,MATCH(' + C + r + '&"*",' + RS + '!$B$1:$Z$1,0)),1)))']);
      }
      rc.getRange(5, hr['Monthly equivalent'] + 1, rows, 1).setFormulas(f);
    }
  }
  // 4. Bills: calendar on Due date
  const bt = sheetOpt_('Bills & Taxes');
  if (bt) { const hb = headerMap_(bt, 4); if (hb['Due date'] !== undefined) bt.getRange(5, hb['Due date'] + 1, 40, 1).setDataValidation(dateRule); }

  repairTxValidation_();
  SpreadsheetApp.getUi().alert(ui_('fx_done', rs.getName(), tx.getName()));
}

function recalcConversionsNow() {
  RATES_MEMO_ = null;
  const r = recalcRows_(2, sheet_('Transactions').getLastRow());
  SpreadsheetApp.getUi().alert(ui_('recalc_done', r.converted, r.missing));
}

/** Simple trigger: runs when someone edits the sheet by hand (not when the bot writes). */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet(), name = sh.getName();
    if (isSheetName_(name, 'Exchange rates')) { RATES_MEMO_ = null; recalcRows_(2, sheet_('Transactions').getLastRow()); return; }
    if (isSheetName_(name, 'Transactions')) {
      const r1 = Math.max(2, e.range.getRow()), r2 = e.range.getLastRow();
      if (r2 >= r1) recalcRows_(r1, r2);
    }
  } catch (err) { console.error(err && err.stack ? err.stack : err); }
}

/** Converts rows r1..r2 of Transactions: foreign-currency rows get Amount (€) = original ÷ rate. */
function recalcRows_(r1, r2) {
  const sh = sheet_('Transactions'), h = headerMap_(sh, 1);
  const out = { converted: 0, missing: 0 };
  if (h['Currency'] === undefined || h['Original amount'] === undefined || r2 < r1) return out;
  const width = sh.getLastColumn();
  const vals = sh.getRange(r1, 1, r2 - r1 + 1, width).getValues();
  const amt = [], notes = [];
  let changed = false;
  vals.forEach(function (v) {
    const cur = String(v[h['Currency']] || '').trim().toUpperCase();
    const orig = Number(v[h['Original amount']]);
    const d = ymd_(v[h['Date']]);
    let a = v[h['Amount']], note = String(v[h['Notes']] || '');
    if (orig && d && cur && cur !== 'EUR') {
      const c = convert_(orig, cur, d);
      a = c.missing ? '' : c.amount;
      note = withMarker_(note, c);
      if (c.missing) out.missing++; else out.converted++;
      changed = true;
    } else if (orig && (cur === '' || cur === 'EUR') && (a === '' || a === null)) {
      a = orig; changed = true;
    }
    amt.push([a]); notes.push([note]);
  });
  if (changed) {
    sh.getRange(r1, h['Amount'] + 1, vals.length, 1).setValues(amt);
    if (h['Notes'] !== undefined) sh.getRange(r1, h['Notes'] + 1, vals.length, 1).setValues(notes);
  }
  return out;
}

/* ---------- import ----------
 * Reads EVERY extra tab that has a table with "Categoría" and "Fecha" columns (your old monthly
 * sheets pasted or imported as they are, or a tab called "Importar") and appends the rows.
 * Understands the old layout: Categoría · Subcategoría · Fecha · Detalle · Pesos · Euros
 *  – "6.100,00 US$" style text, dates like 02/01/26, repeated title rows, empty rows.
 *  – Pesos → Currency ARS + Original amount (€ filled with that month's rate).
 *  – No date → 1st of the month of the nearest dated row, noted "fecha aproximada".
 *  – Rows already in Movimientos are skipped, so you can run it again safely.
 */
const SYSTEM_SHEETS = ['Transactions', 'Recurring', 'Bills & Taxes', 'Settings', 'Exchange rates', 'Dashboard', 'How to use',
  'Analysis', 'Year', 'Search', 'Month', 'Movimientos', 'Recurrentes', 'Facturas e impuestos', 'Configuración', 'Tipos de cambio',
  'Panel', 'Cómo usar', 'Análisis', 'Año', 'Buscar', 'Mes'];
const IMPORT_ALIASES = {
  'fecha': 'Date', 'date': 'Date', 'dia': 'Date',
  'categoria': 'Category', 'category': 'Category',
  'subcategoria': 'Subcategory', 'subcategory': 'Subcategory', 'subcat': 'Subcategory',
  'detalle': 'Description', 'descripcion': 'Description', 'description': 'Description', 'concepto': 'Description', 'detail': 'Description',
  'importe': 'Amount', 'amount': 'Amount', 'euros': 'Amount', 'eur': 'Amount', 'importe €': 'Amount', 'importe eur': 'Amount', '€': 'Amount',
  'importe original': 'Original amount', 'original amount': 'Original amount', 'moneda': 'Currency', 'currency': 'Currency',
  'pesos': 'ARS', 'ars': 'ARS', 'importe pesos': 'ARS', 'importe ars': 'ARS',
  'dolares': 'USD', 'usd': 'USD', 'dolar': 'USD', 'us$': 'USD',
  'tipo': 'Type', 'type': 'Type'
};
function importTransactions() {
  RATES_MEMO_ = null;
  const ui = SpreadsheetApp.getUi();
  const sources = ss_().getSheets().filter(function (sh) { return SYSTEM_SHEETS.indexOf(sh.getName()) < 0; })
    .map(function (sh) { const v = sh.getDataRange().getValues(); return { sh: sh, vals: v, hr: findHeaderRow_(v) }; })
    .filter(function (x) { return x.hr >= 0; });
  if (!sources.length) { ui.alert(ui_('import_none', val_('Import'), val_('Transactions'))); return; }

  const cats = {};
  const loadCats = function () {
    getCategories_('Expense').forEach(function (c) { cats[norm_(c.name)] = { name: c.name, type: 'Expense' }; });
    getCategories_('Income').forEach(function (c) { cats[norm_(c.name)] = { name: c.name, type: 'Income' }; });
  };
  loadCats();

  // 1. read every source tab into plain rows
  const rows = []; let noAmount = 0;
  sources.forEach(function (src) {
    const hi = importHeaderMap_(src.vals[src.hr]);
    const headerNorm = src.vals[src.hr].map(norm_).join('|');
    const block = [];
    src.vals.slice(src.hr + 1).forEach(function (v) {
      if (v.map(norm_).join('|') === headerNorm) return;                         // repeated title row
      const get = function (k) { return hi[k] === undefined ? '' : v[hi[k]]; };
      const catRaw = String(get('Category') || '').trim();
      const eur = anyNum_(get('Amount')), ars = anyNum_(get('ARS')), usd = anyNum_(get('USD'));
      let cur = String(get('Currency') || '').trim().toUpperCase(), orig = anyNum_(get('Original amount'));
      if (!cur) cur = ars ? 'ARS' : (usd ? 'USD' : 'EUR');
      if (cur === 'ARS' && !orig) orig = ars;
      if (cur === 'USD' && !orig) orig = usd;
      if (!catRaw && !eur && !orig) return;                                       // empty line
      if (!eur && !orig) { noAmount++; return; }                                   // e.g. "Gas" with no amount
      const d = ymd_(get('Date')) || parseDateText_(String(get('Date') || '').trim()) || '';
      block.push({ d: d, catRaw: catRaw, cur: cur === 'EUR' && orig && !eur ? 'EUR' : cur, orig: cur === 'EUR' ? 0 : orig,
        amount: eur || (cur === 'EUR' ? orig : 0), get: get, approx: false });
    });
    // undated rows → 1st of the month of the nearest dated row in the same tab
    block.forEach(function (r, i) {
      if (r.d) return;
      let ref = '';
      for (let j = i - 1; j >= 0 && !ref; j--) ref = block[j].d;
      for (let j = i + 1; j < block.length && !ref; j++) ref = block[j].d;
      if (ref) { r.d = ref.slice(0, 8) + '01'; r.approx = true; }
    });
    block.forEach(function (r) { rows.push(r); });
  });

  // 2. categories that don't exist yet → offer to create them as expense categories
  const unknown = {};
  rows.forEach(function (r) { if (!cats[norm_(r.catRaw)]) unknown[r.catRaw || '(vacía / empty)'] = true; });
  const unk = Object.keys(unknown);
  if (unk.length) {
    const ok = unk.indexOf('(vacía / empty)') < 0 &&
      ui.alert(ui_('import_unknown_q', unk.join('\n')), ui.ButtonSet.YES_NO) === ui.Button.YES;
    if (!ok) { ui.alert(ui_('import_unknown', unk.join('\n'))); return; }
    for (let i = 0; i < unk.length; i++) {
      const res = addCategory_(unk[i], 'Expense');
      if (!res.ok) { ui.alert(res.message); return; }
    }
    loadCats();
  }
  if (rows.some(function (r) { return r.cur !== 'EUR'; })) ensureCurrencyColumns_();

  repairTxValidation_();
  // 3. build the new Movimientos rows (sorted by date), skipping ones already there
  const tx = sheet_('Transactions'), ht = headerMap_(tx, 1), width = tx.getLastColumn();
  const existing = transactions_();
  const out = []; let dups = 0, waiting = 0, noDate = 0, approx = 0;
  rows.sort(function (a, b) { return a.d < b.d ? -1 : a.d > b.d ? 1 : 0; });
  rows.forEach(function (r) {
    if (!r.d) { noDate++; return; }
    const cat = cats[norm_(r.catRaw)];
    const conv = r.cur === 'EUR' ? { amount: r.amount } : (r.amount ? { amount: r.amount } : convert_(r.orig, r.cur, r.d));
    const eur = conv.missing ? '' : conv.amount;
    const isDup = existing.some(function (t) {
      return t.date === r.d && t.category === cat.name &&
        ((eur !== '' && Math.abs(t.amount - eur) < 0.005) || (r.cur !== 'EUR' && Math.abs((t.original || 0) - r.orig) < 0.005));
    });
    if (isDup) { dups++; return; }
    if (conv.missing) waiting++;
    if (r.approx) approx++;
    const row = new Array(width).fill('');
    const put = function (k, x) { if (ht[k] !== undefined) row[ht[k]] = safeCell_(x); };
    const typeIn = String(r.get('Type') || '');
    let note = String(r.get('Notes') || '');
    if (r.approx) note = (note ? note + ' ' : '') + (wbLang_() === 'es' ? 'fecha aproximada' : 'approximate date');
    put('Date', toDate_(r.d));
    put('Type', val_(typeIn ? (isVal_(typeIn, 'Income') ? 'Income' : 'Expense') : cat.type));
    put('Category', cat.name);
    put('Description', String(r.get('Description') || '').trim() || String(r.get('Subcategory') || '').trim() || cat.name);
    put('Amount', eur);
    ['For', 'Paid by', 'Payment method', 'Tax deductible', 'Subcategory'].forEach(function (k) { put(k, r.get(k) || ''); });
    put('Notes', r.cur === 'EUR' || r.amount ? note : withMarker_(note, conv));
    put('Source', r.get('Source') || val_('Manual'));
    if (r.cur !== 'EUR') { put('Currency', r.cur); put('Original amount', r.orig); }
    out.push(row);
  });
  if (out.length) {
    try { tx.getRange(tx.getLastRow() + 1, 1, out.length, width).setValues(out); }
    catch (err) { ui.alert(ui_('import_failed', err && err.message ? err.message : String(err))); return; }
  }
  ui.alert(ui_('import_done', out.length, dups, noAmount + noDate,
    (approx ? ui_('import_approx', approx) : '') + (waiting ? ui_('import_norate', waiting, val_('Exchange rates')) : ''),
    sources.map(function (x) { return x.sh.getName(); }).join('", "')));
}
/** Row index (0-based) of the first row, within the first 15, that has both a category and a date title. */
function findHeaderRow_(vals) {
  for (let r = 0; r < Math.min(15, vals.length); r++) {
    const m = importHeaderMap_(vals[r]);
    if (m['Category'] !== undefined && m['Date'] !== undefined) return r;
  }
  return -1;
}
function importHeaderMap_(hdr) {
  const m = {};
  hdr.forEach(function (x, i) {
    const raw = String(x || '').trim(); if (!raw) return;
    const k = IMPORT_ALIASES[norm_(raw)] || WB_REV[raw] || raw;
    if (m[k] === undefined) m[k] = i;
  });
  return m;
}
/** Numbers or text like "6.100,00 US$", "255,00 €", "1,234.50" → number (0 if empty). */
function anyNum_(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || String(v).trim() === '') return 0;
  return parseAmount_(String(v)) || 0;
}

/** Resets the dropdowns of Transactions: removes stray rules (e.g. pasted from another sheet) and sets the right ones
 *  as WARNINGS ONLY (red corner), so neither the bot nor the import can ever be blocked by a rule. */
function repairTxValidation_() {
  const tx = sheet_('Transactions'), h = headerMap_(tx, 1);
  const st = sheet_('Settings'), hs = headerMap_(st, 4), lastSet = st.getLastRow();
  const n = Math.max(tx.getMaxRows() - 1, 1), width = Math.max(tx.getLastColumn(), tx.getMaxColumns ? tx.getMaxColumns() : 0);
  tx.getRange(2, 1, n, width).clearDataValidations();
  const dv = function () { return SpreadsheetApp.newDataValidation(); };
  const set = function (key, builder) { if (h[key] !== undefined) tx.getRange(2, h[key] + 1, n, 1).setDataValidation(builder.setAllowInvalid(true).build()); };
  // categories: rows 5…last row that has a Type (Expense/Income) in Settings
  let catLast = 5;
  if (hs['Type'] !== undefined && lastSet >= 5) {
    st.getRange(5, hs['Type'] + 1, lastSet - 4, 1).getValues().forEach(function (v, i) {
      if (isVal_(v[0], 'Expense') || isVal_(v[0], 'Income')) catLast = 5 + i;
    });
  }
  const settingsCol = function (key) { return hs[key] === undefined ? null : st.getRange(5, hs[key] + 1, Math.max(lastSet - 4, 1), 1); };
  set('Date', dv().requireDate());
  set('Type', dv().requireValueInList([val_('Expense'), val_('Income')], true));
  if (hs['Category'] !== undefined) set('Category', dv().requireValueInRange(st.getRange(5, hs['Category'] + 1, catLast - 4, 1), true));
  [['For', 'For (person)'], ['Paid by', 'Paid by'], ['Payment method', 'Payment method']].forEach(function (p) {
    const r = settingsCol(p[1]); if (r) set(p[0], dv().requireValueInRange(r, true));
  });
  set('Tax deductible', dv().requireValueInList([val_('Yes'), 'No'], true));
  set('Source', dv().requireValueInList([val_('Manual'), 'Bot', 'Auto'], true));
  if (h['Currency'] !== undefined) set('Currency', dv().requireValueInList(currencies_(), true));
}
function repairValidationNow() {
  repairTxValidation_();
  SpreadsheetApp.getUi().alert(ui_('repair_done', val_('Transactions')));
}

/* ---------- currency helpers ---------- */
let RATES_MEMO_ = null;
/** { ARS: [{ym:'2026-01', rate:1500}, …], USD: […] } read from the Exchange rates sheet. */
function rates_() {
  if (RATES_MEMO_) return RATES_MEMO_;
  const out = {}, sh = sheetOpt_('Exchange rates');
  if (sh) {
    const vals = sh.getDataRange().getValues(), hdr = vals[0] || [];
    hdr.forEach(function (x, c) {
      const m = String(x).trim().match(/^([A-Za-z]{3})\b/);
      if (c === 0 || !m) return;
      const code = m[1].toUpperCase(), list = [];
      for (let r = 1; r < vals.length; r++) {
        const ym = ymd_(vals[r][0]).slice(0, 7), rate = Number(vals[r][c]);
        if (ym && rate > 0) list.push({ ym: ym, rate: rate });
      }
      list.sort(function (a, b) { return a.ym < b.ym ? -1 : 1; });
      out[code] = list;
    });
  }
  RATES_MEMO_ = out;
  return out;
}
function currencies_() { return ['EUR'].concat(Object.keys(rates_()).filter(function (c) { return c !== 'EUR'; })); }
/** Rate for a currency in the month of ymd: exact month, else latest earlier (provisional), else earliest later (provisional). */
function rateFor_(cur, ymd) {
  const list = rates_()[cur];
  if (!list || !list.length) return null;
  const ym = ymd.slice(0, 7);
  const exact = list.filter(function (x) { return x.ym === ym; })[0];
  if (exact) return { rate: exact.rate, ym: exact.ym, provisional: false };
  const before = list.filter(function (x) { return x.ym < ym; });
  const pick = before.length ? before[before.length - 1] : list[0];
  return { rate: pick.rate, ym: pick.ym, provisional: true };
}
/** original (in cur) → { amount (€), rate, ym, provisional } or { missing:true } */
function convert_(original, cur, ymd) {
  if (!cur || cur === 'EUR') return { amount: Math.round(original * 100) / 100, rate: 1 };
  const r = rateFor_(cur, ymd);
  if (!r) return { missing: true };
  return { amount: Math.round(original / r.rate * 100) / 100, rate: r.rate, ym: r.ym, provisional: r.provisional };
}
function withMarker_(note, conv) {
  let n = String(note || '');
  ['en', 'es'].forEach(function (l) { [MARK[l].prov, MARK[l].none].forEach(function (m) { n = n.split(m).join(''); }); });
  n = n.replace(/\s+$/, '');
  const mk = MARK[wbLang_()];
  if (conv && conv.missing) n = (n ? n + ' ' : '') + mk.none;
  else if (conv && conv.provisional) n = (n ? n + ' ' : '') + mk.prov;
  return n;
}
function ensureCurrencyColumns_() {
  const tx = sheet_('Transactions');
  addHeaders_(tx, 1, ['Currency', 'Original amount', 'Subcategory']);
}
/** Adds missing headers (in the workbook's language) at the end of a header row. */
function addHeaders_(sh, row, keys) {
  const h = headerMap_(sh, row);
  let col = sh.getLastColumn();
  keys.forEach(function (k) {
    if (h[k] !== undefined) return;
    col++;
    sh.getRange(row, col).setValue(val_(k)).setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#FFFFFF');
  });
}
function recurringLastRow_(rc) {
  const vals = rc.getRange(1, 1, rc.getLastRow(), 1).getValues();
  for (let r = 5; r < vals.length; r++) if (/fixed monthly|fijos mensuales/i.test(String(vals[r][0]))) return r - 1;
  return Math.max(rc.getLastRow(), 5);
}
function isSheetName_(name, en) { return name === en || name === WB_ES[en]; }
function sheetOpt_(name) { return ss_().getSheetByName(name) || ss_().getSheetByName(WB_ES[name] || name); }
function headerMapValues_(hdr) {
  const m = {};
  hdr.forEach(function (x, i) { const k = String(x).trim(); if (!k) return; const en = WB_REV[k] || k; if (m[en] === undefined) m[en] = i; });
  return m;
}
function colL_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
/** "42000 ars", "42.000 pesos", "$42000", "100 usd", "45,20" (€) → { amount, currency } */
function parseMoney_(t) {
  let s = String(t).trim(), cur = 'EUR';
  const pats = [[/(us\$|u\$s|usd|d[oó]lar(es)?|dollars?)/i, 'USD'], [/(ars|pesos?|\$)/i, 'ARS'], [/(eur|euros?|€)/i, 'EUR'],
    [/\b([a-z]{3})\b/i, null]];
  for (let i = 0; i < pats.length; i++) {
    const m = s.match(pats[i][0]);
    if (m) { cur = pats[i][1] || m[1].toUpperCase(); s = s.replace(pats[i][0], ' '); break; }
  }
  const a = parseAmount_(s);
  return a === null ? null : { amount: a, currency: cur };
}
function fmtNum_(n, dec) {
  const v = Math.abs(Math.round(n * Math.pow(10, dec)) / Math.pow(10, dec)).toFixed(dec).split('.');
  const th = LANG_ === 'es' ? '.' : ',', ds = LANG_ === 'es' ? ',' : '.';
  return (n < 0 ? '-' : '') + v[0].replace(/\B(?=(\d{3})+(?!\d))/g, th) + (dec ? ds + v[1] : '');
}
/** 42.000 ARS · 1.234,56 € */
function fmtCur_(n, cur) {
  if (!cur || cur === 'EUR') return money_(n);
  return fmtNum_(n, Math.round(n) === n ? 0 : 2) + ' ' + cur;
}

/** Runs every morning (installed by installDailyTrigger). */
function dailyJob() {
  if (isVal_(getSetting_('Auto-post recurring items', 'Yes'), 'Yes')) postRecurring_();
  sendReminders_(false);
}

/* =========================================================================
 *  Telegram webhook
 * ========================================================================= */
function doPost(e) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('WEBHOOK_KEY');
    if (!key || !e || !e.parameter || e.parameter.key !== key) return ok_();
    const update = JSON.parse(e.postData.contents);
    // Telegram may resend the same update – process each one only once.
    const cache = CacheService.getScriptCache();
    const seen = 'upd_' + update.update_id;
    if (cache.get(seen)) return ok_();
    cache.put(seen, '1', 21600);
    handleUpdate_(update);
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
  }
  return ok_();
}
// HtmlService (not ContentService): Telegram needs a plain 200 OK – ContentService answers with a 302 redirect,
// which Telegram treats as a failure and keeps retrying, so the bot looks frozen.
function ok_() { return HtmlService.createHtmlOutput('ok'); }

function handleUpdate_(u) {
  const msg = u.message, cb = u.callback_query;
  const from = msg ? msg.from : (cb ? cb.from : null);
  const chat = msg ? msg.chat : (cb && cb.message ? cb.message.chat : null);
  if (!from || !chat) return;
  // Closed bot: only private one-to-one chats – never groups or channels, even if someone adds the bot to one.
  if (chat.type && chat.type !== 'private') return;
  if (String(chat.id) !== String(from.id)) return;
  const chatId = chat.id;
  let role = null;
  try { role = role_(from.id); } catch (e) { console.error(e && e.stack ? e.stack : e); return; }
  // Not on the list → complete silence (nothing is read from or written to the sheet for them).
  if (!role) { stranger_(msg, cb, from, chatId); return; }
  try {
    handleUpdateInner_(msg, cb, from, chatId, role);
  } catch (err) {
    // Only family members ever see an error message.
    console.error(err && err.stack ? err.stack : err);
    try { send_(chatId, '⚠️ ' + esc_(err && err.message ? err.message : String(err))); } catch (e) { }
  }
}
/** Someone who is not in Settings. Default: no answer at all.
 *  Exceptions: they send the secret word (set from the sheet menu, stored hidden in the script),
 *  or no IDs are configured yet (very first setup). */
function stranger_(msg, cb, from, chatId) {
  if (cb || !msg || !msg.text) return;
  const firstSetup = !idsFor_('Allowed Telegram user IDs').length;
  if (!firstSetup && !isSecret_(msg.text)) return;
  LANG_ = userLang_(from.id, from.language_code);
  if (firstSetup) {
    send_(chatId, t_('private', from.id, label_('Allowed Telegram user IDs'), label_('Read-only Telegram user IDs')));
    return;
  }
  send_(chatId, t_('id_reply', from.id));
  // Tell the full-access members (at most once every 10 minutes per person, so nobody can spam you).
  const cache = CacheService.getScriptCache(), k = 'join_' + from.id;
  if (cache.get(k)) return;
  cache.put(k, '1', 600);
  const name = esc_([from.first_name, from.last_name].filter(Boolean).join(' ') || String(from.id));
  const user = from.username ? ' (@' + esc_(from.username) + ')' : '';
  broadcast_(function () { return t_('id_admin', name, user, from.id); }, function () {
    return { reply_markup: kb_([[btn_(t_('add_full'), 'ad|' + from.id + '|f'), btn_(t_('add_read'), 'ad|' + from.id + '|r')],
      [btn_(t_('add_ignore'), 'ad|' + from.id + '|x')]]) };
  }, false);
}
/** The text is the secret word (also accepted as "/id word" or "/start word"). */
function isSecret_(text) {
  const secret = PropertiesService.getScriptProperties().getProperty('JOIN_SECRET');
  if (!secret) return false;
  const t = String(text).trim().replace(/^\/(id|start)(@\w+)?\s+/i, '');
  return norm_(t) === norm_(secret);
}
/** Menu ▸ Secret word for new users… */
function setSecretWord() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  const now = props.getProperty('JOIN_SECRET');
  const r = ui.prompt(ui_('secret'), ui_('secret_p', now || ui_('secret_none')), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const w = r.getResponseText().trim();
  if (!w) { props.deleteProperty('JOIN_SECRET'); ui.alert(ui_('secret_off')); return; }
  if (/\s/.test(w) || w.length < 4) { ui.alert(ui_('secret_bad')); return; }
  props.setProperty('JOIN_SECRET', w);
  ui.alert(ui_('secret_on', w));
}
/** Button on the "wants to join" message: adds the ID to the full-access or read-only list in Settings. */
function addMember_(chatId, msgId, id, how) {
  id = String(id).replace(/\D/g, '');
  if (!id) return;
  if (how === 'x') return edit_(chatId, msgId, t_('ignored'));
  const full = 'Allowed Telegram user IDs', read = 'Read-only Telegram user IDs';
  const target = how === 'f' ? full : read, other = how === 'f' ? read : full;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    setSettingIds_(other, idsFor_(other).filter(function (x) { return x !== id; }));
    const list = idsFor_(target);
    if (list.indexOf(id) < 0) list.push(id);
    setSettingIds_(target, list);
  } finally { lock.releaseLock(); }
  edit_(chatId, msgId, t_('added', '<code>' + id + '</code>', t_(how === 'f' ? 'full_access' : 'read_access')));
  const before = LANG_;
  LANG_ = userLang_(id); send_(id, t_('welcome')); LANG_ = before;
}
/** Writes a list of IDs into the value cell next to a setting label (either language). */
function setSettingIds_(label, ids) {
  const sh = sheet_('Settings'), vals = sh.getDataRange().getValues();
  const want = [norm_(label), norm_(WB_ES[label] || label)];
  for (let r = 0; r < vals.length; r++) for (let c = 0; c < vals[r].length - 1; c++) {
    if (want.indexOf(norm_(vals[r][c])) >= 0) {
      const cell = sh.getRange(r + 1, c + 2);
      if (cell.setNumberFormat) cell.setNumberFormat('@');
      cell.setValue(ids.join(', '));
      return true;
    }
  }
  throw new Error('"' + val_(label) + '" not found in ' + val_('Settings'));
}
function handleUpdateInner_(msg, cb, from, chatId, role) {
  LANG_ = userLang_(from.id, from.language_code);

  if (cb) {
    if (role !== 'full') {
      if (/^rc\|/.test(String(cb.data || ''))) {
        tg_('answerCallbackQuery', { callback_query_id: cb.id });
        return recentCallback_(chatId, cb.message.message_id, String(cb.data).split('|'));
      }
      tg_('answerCallbackQuery', { callback_query_id: cb.id, text: t_('readonly_alert'), show_alert: true });
      return;
    }
    onCallback_(chatId, cb, from); return;
  }
  if (msg.text) onText_(chatId, msg.text.trim(), from, role);
}

/** 'full' = can add/undo/mark paid · 'read' = can only search and see reports · null = not allowed */
function role_(userId) {
  const id = String(userId);
  if (idsFor_('Allowed Telegram user IDs').indexOf(id) >= 0) return 'full';
  if (idsFor_('Read-only Telegram user IDs').indexOf(id) >= 0) return 'read';
  return null;
}

/* =========================================================================
 *  Conversation
 * ========================================================================= */
function onText_(chatId, text, from, role) {
  const t = text.replace(/^(\/\w+)@\w+/, '$1');
  const cmd = t.toLowerCase();

  // language switch (works for everyone)
  if (cmd === '/english' || cmd === '/espanol' || cmd === '/español' || cmd === '/language' || cmd === '/idioma') {
    if (cmd === '/language' || cmd === '/idioma') LANG_ = LANG_ === 'es' ? 'en' : 'es';
    else LANG_ = cmd === '/english' ? 'en' : 'es';
    setUserLang_(from.id, LANG_);
    clearState_(chatId);
    return send_(chatId, t_('lang_set'), mainKeyboard_(role));
  }
  const KB = mainKeyboard_(role);
  if (cmd === '/id' || isSecret_(t)) return send_(chatId, t_('id_own', from.id), KB);
  if (cmd === '/start' || cmd === '/menu' || cmd === '/help') {
    clearState_(chatId);
    const help = role === 'full'
      ? t_('help_full', t_('menu_expense'), t_('menu_tax'), t_('menu_income'), t_('menu_search'), t_('menu_month'), t_('menu_upcoming'), t_('menu_undo'), t_('menu_recent'))
      : t_('help_read', t_('menu_search'), t_('menu_month'), t_('menu_upcoming'), t_('menu_recent'));
    const helpX = exportAllowed_(role) ? '\n' + t_('help_export', t_('menu_export')) : '';
    return send_(chatId, t_('hi', esc_(from.first_name || '')) + '\n\n' + help + helpX + '\n\n' + t_('help_tip') + '\n' + t_('lang_hint'), KB);
  }
  if (cmd === '/cancel') { clearState_(chatId); return send_(chatId, t_('cancelled'), KB); }

  const action = menuAction_(t);
  if (action === 'month') { clearState_(chatId); return send_(chatId, monthReport_(), KB); }
  if (action === 'upcoming') { clearState_(chatId); return send_(chatId, upcomingReport_(), KB); }
  if (action === 'recent') { clearState_(chatId); const rv = recentView_({}); return send_(chatId, rv.text, { reply_markup: kb_(rv.rows) }); }
  if (action === 'search') { saveState_(chatId, { step: 'search' }); return send_(chatId, t_('search_prompt'), KB); }
  if (/^\/(search|buscar)\b/.test(cmd)) { clearState_(chatId); return send_(chatId, searchReport_(t.replace(/^\/\w+\s*/, '')), KB); }
  if (action && role !== 'full') return send_(chatId, t_('readonly_no_write'), KB);
  if (action === 'export') {
    clearState_(chatId);
    if (!exportAllowed_(role)) return send_(chatId, t_('export_off', val_('Allow export from Telegram'), val_('Settings')), KB);
    return send_(chatId, t_('export_pick'), { reply_markup: kb_(exportRows_()) });
  }
  if (action === 'expense' || action === 'tax' || action === 'income') return startFlow_(chatId, action);
  if (action === 'undo') { clearState_(chatId); return askUndo_(chatId); }

  const s = getState_(chatId);
  // No entry in progress (or a search was asked for): treat whatever was typed as a search.
  if (!s || s.step === 'search' || role !== 'full') {
    clearState_(chatId);
    if (t.charAt(0) === '/') return send_(chatId, t_('tap_start'), KB);
    return send_(chatId, searchReport_(t), KB);
  }

  if (s.step === 'newcat') {
    const name = t.replace(/\s+/g, ' ').trim();
    if (name.length < 2 || name.length > 40 || /[|]/.test(name)) return send_(chatId, t_('newcat_len'));
    const res = addCategory_(name, s.type);
    if (!res.ok) return send_(chatId, res.message);
    s.category = res.name;
    send_(chatId, res.created ? t_('newcat_added', esc_(res.name)) : t_('newcat_existing', esc_(res.name)));
    return askDate_(chatId, s);
  }
  if (s.step === 'date') {   // typed instead of tapping a button
    const d = parseDateText_(t);
    if (!d) return send_(chatId, t_('date_bad'));
    if (d > today_()) return send_(chatId, t_('date_future'));
    if (s.msgId) edit_(chatId, s.msgId, t_('date_line', showDate_(d)));
    s.date = d; return askAmount_(chatId, s);
  }
  if (s.step === 'amount') {
    const m = parseMoney_(t);
    if (m === null) return send_(chatId, t_('amount_bad'));
    s.currency = 'EUR'; s.original = null; s.rate = null; s.provisional = false;
    if (m.currency !== 'EUR') {
      const known = currencies_();
      if (known.indexOf(m.currency) < 0) return send_(chatId, t_('unknown_cur', known.join(', '), val_('Exchange rates')));
      const c = convert_(m.amount, m.currency, s.date);
      if (c.missing) return send_(chatId, t_('no_rate', m.currency, val_('Exchange rates')));
      s.currency = m.currency; s.original = m.amount; s.rate = c.rate; s.rateYm = c.ym; s.provisional = c.provisional;
      s.amount = c.amount;
    } else {
      s.amount = m.amount;
    }
    return askDesc_(chatId, s);
  }
  if (s.step === 'desc') {
    s.desc = s.flow === 'tax' ? s.taxType + ' – ' + t : t;
    if (s.msgId) edit_(chatId, s.msgId, t_('desc_line', esc_(s.desc)));
    return askWho_(chatId, s);
  }
  send_(chatId, t_('use_buttons'));
}

function onCallback_(chatId, cb, from) {
  tg_('answerCallbackQuery', { callback_query_id: cb.id });
  const parts = String(cb.data || '').split('|');
  const kind = parts[0], val = parts[1];
  const msgId = cb.message.message_id;
  if (kind === 'nop') return;

  // Buttons that don't belong to an entry in progress
  if (kind === 'u') return doUndo_(chatId, msgId, val === 'y');
  if (kind === 'b') return markBillPaid_(chatId, msgId, parseInt(val, 10), parts[2]);
  if (kind === 'ad') return addMember_(chatId, msgId, val, parts[2]);
  if (kind === 'rc') return recentCallback_(chatId, msgId, parts);
  if (kind === 'ex') return exportCallback_(chatId, msgId, parts);

  const s = getState_(chatId);
  if (kind === 'x') { clearState_(chatId); return edit_(chatId, msgId, '❌ ' + t_('cancelled')); }
  if (!s || s.msgId !== msgId) return edit_(chatId, msgId, t_('expired'));

  if (kind === 'n' && s.step === 'cat') {
    s.step = 'newcat'; s.msgId = null; saveState_(chatId, s);
    return edit_(chatId, msgId, title_(s) + '\n' + t_('newcat_prompt', t_(s.type === 'Income' ? 'w_income' : 'w_expense')));
  }
  if (kind === 'c' && s.step === 'cat') {
    s.category = s.options[+val]; if (!s.category) return;
    edit_(chatId, msgId, title_(s) + '\n' + t_('category_line', esc_(s.category)) + '\n\n' + lastInCategory_(s.category));
    return askDate_(chatId, s);
  }
  if (kind === 't' && s.step === 'taxtype') {
    s.taxType = s.options[+val]; if (!s.taxType) return;
    edit_(chatId, msgId, title_(s) + '\n' + t_('tax_line', esc_(s.taxType)) + '\n\n' + lastInCategory_(s.category));
    return askDate_(chatId, s);
  }
  if (s.step === 'date') {
    if (kind === 'd' && val === 'c') return edit_(chatId, msgId, t_('pick_day'), calendarRows_(today_().slice(0, 7)));
    if (kind === 'cm') return edit_(chatId, msgId, t_('pick_day'), calendarRows_(val));
    if (kind === 'cb') return edit_(chatId, msgId, t_('when'), quickDateRows_());
    let d = null;
    if (kind === 'd') d = addDays_(today_(), -(+val));
    if (kind === 'cd') d = val.slice(0, 4) + '-' + val.slice(4, 6) + '-' + val.slice(6, 8);
    if (!d || d > today_()) return;
    s.date = d;
    edit_(chatId, msgId, t_('date_line', showDate_(d)));
    return askAmount_(chatId, s);
  }
  if (kind === 'k' && s.step === 'desc') {
    s.desc = s.flow === 'tax' ? s.taxType : s.category;
    edit_(chatId, msgId, t_('desc_line', esc_(s.desc)));
    return askWho_(chatId, s);
  }
  if (kind === 'w' && s.step === 'who') {
    s.who = s.options[+val] || '';
    edit_(chatId, msgId, t_('for_line', esc_(s.who)));
    return s.type === 'Expense' ? askPayer_(chatId, s) : askConfirm_(chatId, s);
  }
  if (kind === 'p' && s.step === 'payer') {
    s.payer = s.options[+val] || '';
    edit_(chatId, msgId, t_('paid_line', esc_(s.payer)));
    return askConfirm_(chatId, s);
  }
  if (kind === 'ok' && s.step === 'confirm') return save_(chatId, msgId, s, from);
}

/* ---------- the steps ---------- */
function startFlow_(chatId, flow) {
  const s = { flow: flow, type: flow === 'income' ? 'Income' : 'Expense' };
  if (flow === 'tax') {
    s.category = taxCategory_();
    s.step = 'taxtype';
    s.options = getList_('Tax types');
    return ask_(chatId, s, title_(s) + '\n' + t_('which_tax'), grid_(s.options, 't'));
  }
  s.step = 'cat';
  s.options = getCategories_(s.type).map(function (c) { return c.name; });
  ask_(chatId, s, title_(s) + '\n' + t_('choose_cat'), grid_(s.options, 'c').concat([[btn_(t_('new_cat_btn'), 'n|1')]]));
}
function quickDateRows_() {
  return [[btn_(t_('today'), 'd|0'), btn_(t_('yesterday'), 'd|1')], [btn_(t_('two_days'), 'd|2'), btn_(t_('calendar'), 'd|c')]];
}
function askDate_(chatId, s) {
  s.step = 'date';
  ask_(chatId, s, t_('when'), quickDateRows_());
}
/** Month grid (Monday first). Future days can't be picked; ◀ ▶ change month. ym = 'yyyy-MM' */
function calendarRows_(ym) {
  const y = +ym.slice(0, 4), m = +ym.slice(5, 7), t = today_();
  const first = fromParts_(y, m, 1);
  const prev = addMonths_(first, -1).slice(0, 7), next = addMonths_(first, 1).slice(0, 7);
  const rows = [[btn_('◀', 'cm|' + prev), btn_(monthName_(first), 'nop|1'),
    next <= t.slice(0, 7) ? btn_('▶', 'cm|' + next) : btn_('·', 'nop|1')]];
  rows.push(t_('dows').map(function (d) { return btn_(d, 'nop|1'); }));
  const lead = (new Date(Date.UTC(y, m - 1, 1)).getUTCDay() + 6) % 7;
  const days = new Date(Date.UTC(y, m, 0)).getUTCDate();
  let row = [];
  for (let i = 0; i < lead; i++) row.push(btn_('·', 'nop|1'));
  for (let d = 1; d <= days; d++) {
    const ymd = fromParts_(y, m, d);
    row.push(ymd <= t ? btn_(ymd === t ? '[' + d + ']' : String(d), 'cd|' + ymd.replace(/-/g, '')) : btn_('·', 'nop|1'));
    if (row.length === 7) { rows.push(row); row = []; }
  }
  if (row.length) { while (row.length < 7) row.push(btn_('·', 'nop|1')); rows.push(row); }
  rows.push([btn_(t_('back'), 'cb|1'), btn_(t_('cancel_btn'), 'x|1')]);
  return rows;
}
function askAmount_(chatId, s) {
  s.step = 'amount'; s.msgId = null; saveState_(chatId, s);
  send_(chatId, t_(currencies_().length > 1 ? 'how_much_fx' : 'how_much'));
}
function askDesc_(chatId, s) {
  s.step = 'desc';
  ask_(chatId, s, t_(s.flow === 'tax' ? 'desc_prompt_tax' : 'desc_prompt'), [[btn_(t_('skip'), 'k|1')]]);
}
function askWho_(chatId, s) {
  s.step = 'who';
  s.options = getList_('For (person)');
  ask_(chatId, s, t_(s.type === 'Income' ? 'whose_income' : 'who_for'), grid_(s.options, 'w'));
}
function askPayer_(chatId, s) {
  s.step = 'payer';
  s.options = getList_('Paid by');
  ask_(chatId, s, t_('who_paid'), grid_(s.options, 'p'));
}
function askConfirm_(chatId, s) {
  s.step = 'confirm';
  const lines = [
    t_('check_save'),
    t_(s.type === 'Income' ? 'kind_income' : (s.flow === 'tax' ? 'kind_tax' : 'kind_expense')),
    '📂 ' + esc_(s.category),
    '📅 <b>' + showDate_(s.date, true) + '</b>',
    s.currency && s.currency !== 'EUR'
      ? '💶 <b>' + fmtCur_(s.original, s.currency) + '</b> ' + t_('conv', money_(s.amount), fmtNum_(s.rate, s.rate % 1 ? 2 : 0), s.currency,
        monthName_(s.rateYm + '-01'), s.provisional ? t_('provisional') : '')
      : '💶 <b>' + money_(s.amount) + '</b>',
    '📝 ' + esc_(s.desc),
    '👤 ' + esc_(s.who || '–') + (s.type === 'Expense' ? '   💳 ' + esc_(s.payer || '–') : '')
  ];
  const dups = findDuplicates_(s.category, s.amount, s.date, DUP_DAYS);
  if (dups.length) {
    lines.push('', t_('dup_title'));
    dups.slice(0, 3).forEach(function (d) {
      lines.push('• ' + showDate_(d.date) + ' · ' + money_(d.amount) + ' · <i>' + esc_(d.desc || d.category) + '</i>' + (d.source ? ' (' + esc_(d.source) + ')' : ''));
    });
  }
  ask_(chatId, s, lines.join('\n'), [[btn_(t_(dups.length ? 'save_anyway' : 'save'), 'ok|1'), btn_(t_('cancel_btn'), 'x|1')]], true);
}

function save_(chatId, msgId, s, from) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let row;
  try {
    row = appendTransaction_({
      'Date': s.date, 'Type': val_(s.type), 'Category': s.category, 'Description': s.desc, 'Amount': s.amount,
      'For': s.who || '', 'Paid by': s.payer || '', 'Payment method': '', 'Tax deductible': '',
      'Notes': withMarker_((wbLang_() === 'es' ? 'por ' : 'by ') + (from.first_name || from.id),
        s.currency && s.currency !== 'EUR' ? { provisional: s.provisional } : null),
      'Source': val_('Bot'),
      'Currency': s.currency && s.currency !== 'EUR' ? s.currency : '',
      'Original amount': s.currency && s.currency !== 'EUR' ? s.original : ''
    });
  } finally { lock.releaseLock(); }
  clearState_(chatId);
  CacheService.getScriptCache().put('last_' + chatId, JSON.stringify({
    row: row, date: s.date, category: s.category, amount: s.amount, desc: s.desc
  }), 7 * 24 * 3600);
  edit_(chatId, msgId, t_('saved', esc_(s.category), money_(s.amount), showDate_(s.date)));
  send_(chatId, categoryProgress_(s.category, s.date, s.type), mainKeyboard_('full'));
}

/** Entries with the same category and amount within ±days of the date. */
function findDuplicates_(category, amount, date, days) {
  return transactions_().filter(function (r) {
    return r.category === category && Math.abs(r.amount - amount) < 0.005 && Math.abs(daysBetween_(r.date, date)) <= days;
  });
}

/* ---------- undo ---------- */
function askUndo_(chatId) {
  const last = JSON.parse(CacheService.getScriptCache().get('last_' + chatId) || 'null');
  if (!last) return send_(chatId, t_('undo_nothing'), mainKeyboard_('full'));
  send_(chatId, t_('undo_ask') + '\n' + showDate_(last.date) + ' · ' + esc_(last.category) + ' · <b>' + money_(last.amount) +
    '</b>\n<i>' + esc_(last.desc) + '</i>', { reply_markup: kb_([[btn_(t_('yes_delete'), 'u|y'), btn_(t_('keep'), 'u|n')]]) });
}
function doUndo_(chatId, msgId, yes) {
  const cache = CacheService.getScriptCache();
  const last = JSON.parse(cache.get('last_' + chatId) || 'null');
  if (!yes) return edit_(chatId, msgId, t_('kept'));
  if (!last) return edit_(chatId, msgId, t_('undo_nothing'));
  const sh = sheet_('Transactions'), h = headerMap_(sh, 1);
  if (last.row > sh.getLastRow()) return edit_(chatId, msgId, t_('undo_gone'));
  const v = sh.getRange(last.row, 1, 1, sh.getLastColumn()).getValues()[0];
  const same = ymd_(v[h['Date']]) === last.date && v[h['Category']] === last.category &&
    Math.abs(Number(v[h['Amount']]) - last.amount) < 0.005;
  if (!same) return edit_(chatId, msgId, t_('undo_changed'));
  sh.deleteRow(last.row);
  cache.remove('last_' + chatId);
  edit_(chatId, msgId, t_('deleted'));
}

/* =========================================================================
 *  Reports & search
 * ========================================================================= */
function monthReport_() {
  const t = today_(), start = t.slice(0, 8) + '01', end = nextMonth_(start);
  let inc = 0, exp = 0; const byCat = {};
  transactions_().forEach(function (r) {
    if (r.date < start || r.date >= end) return;
    if (r.type === 'Income') inc += r.amount;
    else { exp += r.amount; byCat[r.category] = (byCat[r.category] || 0) + r.amount; }
  });
  const cats = getCategories_('Expense');
  const budget = cats.reduce(function (a, c) { return a + c.budget; }, 0);
  const bud = {}; cats.forEach(function (c) { bud[c.name] = c.budget; });
  const top = Object.keys(byCat).sort(function (a, b) { return byCat[b] - byCat[a]; }).slice(0, 8);
  const out = [
    '📊 <b>' + cap_(monthName_(start)) + '</b> ' + t_('so_far'),
    '➕ ' + t_('income') + ': <b>' + money_(inc) + '</b>',
    '➖ ' + t_('expenses') + ': <b>' + money_(exp) + '</b>',
    (inc - exp >= 0 ? '🟢' : '🔴') + ' ' + t_('net') + ': <b>' + money_(inc - exp) + '</b>'
  ];
  if (budget > 0) out.push(t_('budget_used', Math.round(exp / budget * 100), money_(budget)));
  if (top.length) {
    out.push('', '<b>' + t_('top') + '</b>');
    top.forEach(function (c) {
      const b = bud[c] || 0;
      const flag = b > 0 ? (byCat[c] > b ? ' 🔴' : (byCat[c] > 0.85 * b ? ' 🟠' : '')) : '';
      out.push('• ' + esc_(c) + ': ' + money_(byCat[c]) + (b > 0 ? ' / ' + money_(b) : '') + flag);
    });
  }
  return out.join('\n');
}

function categoryProgress_(category, date, type) {
  const start = date.slice(0, 8) + '01', end = nextMonth_(start);
  const total = transactions_().reduce(function (a, r) {
    return a + (r.category === category && r.date >= start && r.date < end ? r.amount : 0);
  }, 0);
  const c = getCategories_(type).filter(function (x) { return x.name === category; })[0];
  const b = c ? c.budget : 0;
  let line = t_('progress', esc_(category), monthName_(start), money_(total));
  if (b > 0) {
    const pct = Math.round(total / b * 100);
    line += type === 'Income' ? t_('progress_expected', money_(b))
      : t_('progress_budget', money_(b), pct) + (pct > 100 ? t_('over_budget') : pct > 85 ? ' 🟠' : ' 🟢');
  }
  return line;
}

function upcomingReport_() {
  const t = today_(), until = addDays_(t, 30), items = [];
  billRows_().forEach(function (b) {
    if (b.pending && b.due && b.due <= until) items.push({ date: b.due, text: '🧾 ' + esc_(b.item) + (b.amount ? ' – ' + money_(b.amount) : '') });
  });
  recurringRows_().forEach(function (r) {
    if (!r.active || !r.first || !r.amount) return;
    for (let k = 0; k < 1200; k += r.every) {
      const d = addMonths_(r.first, k);
      if ((r.end && d > r.end) || d > until) break;
      if (d >= t) items.push({ date: d, text: '🔁 ' + esc_(r.name) + ' – ' + fmtCur_(r.amount, r.currency) + (r.type === 'Income' ? t_('in_suffix') : '') });
    }
  });
  if (!items.length) return t_('nothing_due');
  items.sort(function (a, b) { return a.date < b.date ? -1 : 1; });
  return t_('next30') + '\n' + items.map(function (i) {
    return (i.date < t ? '⚠️ ' : '') + showDate_(i.date) + '  ' + i.text;
  }).join('\n');
}

/* ---------- recent entries ---------- */
const RECENT_MAX = 10;          // entries in the 🕘 Recent list
const RECENT_IN_CATEGORY = 5;   // entries shown when you pick a category while adding
/** One entry, two lines: date · amount · category / subcategory, then description · 👤 for whom · ✍️ who logged it. */
function entryLine_(r) {
  const amt = r.currency && r.currency !== 'EUR' && r.original
    ? fmtCur_(r.original, r.currency) + ' (' + money_(r.amount) + ')' : money_(r.amount);
  const cat = esc_(r.category) + (r.subcat ? ' / ' + esc_(r.subcat) : '');
  const desc = r.desc && r.desc !== r.category && r.desc !== r.subcat ? '<i>' + esc_(r.desc) + '</i>' : '';
  const logged = loggedBy_(r.notes);
  const bits = [desc, r.who ? '👤 ' + esc_(r.who) : '', logged ? '✍️ ' + esc_(logged) : ''].filter(String);
  return '<b>' + shortDate_(r.date) + '</b> · ' + (r.type === 'Income' ? '+' : '') + amt + ' · ' + cat +
    (bits.length ? '\n      ' + bits.join(' · ') : '');
}
/** "por Lucía" / "by Lucía" written in Notes by the bot → "Lucía". */
function loggedBy_(notes) {
  const m = String(notes || '').match(/(?:^|\s)(?:por|by)\s+([^\[\]]+?)\s*(?:\[|$)/i);
  return m ? m[1].trim() : '';
}
/** "8 nov" this year, "8 nov 2025" other years. */
function shortDate_(ymd) {
  const s = (+ymd.slice(8, 10)) + ' ' + t_('mon')[+ymd.slice(5, 7) - 1];
  return ymd.slice(0, 4) === today_().slice(0, 4) ? s : s + ' ' + ymd.slice(0, 4);
}
/** Entries newest first (by date, then by row), optionally only one category or one person. */
function recentEntries_(f, max) {
  return transactions_().map(function (r, i) { r.i = i; return r; })
    .filter(function (r) {
      if (f.category && r.category !== f.category) return false;
      if (f.who && norm_(r.who) !== norm_(f.who)) return false;
      return true;
    })
    .sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : b.i - a.i; })
    .slice(0, max);
}
/** Shown under the category while adding an entry, so you can see if it is already logged. */
function lastInCategory_(category) {
  const items = recentEntries_({ category: category }, RECENT_IN_CATEGORY);
  if (!items.length) return t_('none_in_cat', esc_(category));
  return t_('last_in_cat', esc_(category)) + '\n' + items.map(entryLine_).join('\n');
}
/** Categories from Settings (expenses, then income) that already have entries – for the filter buttons. */
function recentCategories_() {
  const used = {};
  transactions_().forEach(function (r) { used[r.category] = true; });
  return getCategories_('Expense').concat(getCategories_('Income'))
    .map(function (c) { return c.name; }).filter(function (n) { return used[n]; });
}
function recentView_(f) {
  const items = recentEntries_(f, RECENT_MAX);
  const label = f.category ? ' · ' + esc_(f.category) : (f.who ? ' · 👤 ' + esc_(f.who) : '');
  const text = items.length
    ? t_('recent_title', items.length, label) + '\n\n' + items.map(entryLine_).join('\n')
    : t_('recent_none') + (label ? '\n' + label.replace(/^ · /, '') : '');
  return { text: text, rows: [[btn_(t_('recent_all'), 'rc|a'), btn_(t_('recent_cat'), 'rc|c'), btn_(t_('recent_who'), 'rc|w')]] };
}
/** Buttons on the Recent message: rc|a all · rc|c pick category · rc|ci|n category n · rc|w pick person · rc|wi|n person n */
function recentCallback_(chatId, msgId, parts) {
  const k = parts[1];
  if (k === 'c') return edit_(chatId, msgId, t_('recent_pick_cat'), grid_(recentCategories_(), 'rc|ci').concat([[btn_(t_('back'), 'rc|a')]]));
  if (k === 'w') return edit_(chatId, msgId, t_('recent_pick_who'), grid_(getList_('For (person)'), 'rc|wi').concat([[btn_(t_('back'), 'rc|a')]]));
  const f = {};
  if (k === 'ci') f.category = recentCategories_()[+parts[2]];
  if (k === 'wi') f.who = getList_('For (person)')[+parts[2]];
  const rv = recentView_(f);
  return edit_(chatId, msgId, rv.text, rv.rows);
}

/** Month words in both languages, for searches like "school september" / "colegio septiembre". */
function monthWord_(w) {
  const x = norm_(w);
  for (let i = 0; i < 12; i++) {
    const cands = [TXT.en.months[i], TXT.es.months[i], TXT.en.mon[i], TXT.es.mon[i]].map(norm_);
    if (cands.indexOf(x) >= 0) return i + 1;
  }
  return 0;
}
/** Every word must appear in the description, category, notes, "for" or "paid by".
 *  A year (2026), a month (09/2026) or a month name (september, septiembre [2026]) narrows the period. */
function searchReport_(query) {
  let q = String(query || '').trim();
  if (!q) return t_('search_empty');
  let from = '', to = '', period = '';
  const mm = q.match(/\b(\d{1,2})[\/.\-](\d{4}|\d{2})\b/);
  const yy = q.match(/\b(20\d{2})\b/);
  if (mm && +mm[1] >= 1 && +mm[1] <= 12) {
    const y = +mm[2] < 100 ? 2000 + +mm[2] : +mm[2];
    from = fromParts_(y, +mm[1], 1); q = q.replace(mm[0], ' ');
  } else {
    const words0 = q.split(/\s+/);
    const mw = words0.filter(function (w) { return monthWord_(w) > 0 && w.length >= 3; })[0];
    if (mw) {
      const y = yy ? +yy[1] : +today_().slice(0, 4);
      let m = monthWord_(mw);
      from = fromParts_(y, m, 1);
      if (!yy && from > today_()) from = fromParts_(y - 1, m, 1);  // "december" in October = last December
      q = q.replace(mw, ' '); if (yy) q = q.replace(yy[0], ' ');
    } else if (yy) {
      from = yy[1] + '-01-01'; to = (+yy[1] + 1) + '-01-01'; period = yy[1]; q = q.replace(yy[0], ' ');
    }
  }
  if (from && !to) { to = nextMonth_(from); period = monthName_(from); }
  const words = q.split(/\s+/).map(norm_).filter(String);
  const hits = transactions_().filter(function (r) {
    if (from && (r.date < from || r.date >= to)) return false;
    const hay = norm_([r.desc, r.category, r.subcat, r.notes, r.who, r.payer].join(' '));
    return words.every(function (w) { return hay.indexOf(w) >= 0; });
  }).sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  const label = esc_(q.trim().replace(/\s+/g, ' ') || t_('everything')) + (period ? ' · ' + period : '');
  if (!hits.length) return t_('nothing_found', label);
  let out = 0, inn = 0;
  hits.forEach(function (r) { if (r.type === 'Income') inn += r.amount; else out += r.amount; });
  const lines = ['🔎 <b>' + label + '</b> – ' + hits.length + ' ' + t_(hits.length === 1 ? 'entry' : 'entries')];
  if (out) lines.push(t_('spent', money_(out)));
  if (inn) lines.push(t_('received', money_(inn)));
  lines.push('');
  // Who it belongs to: a total per person when the results are for more than one.
  const per = {};
  hits.forEach(function (r) { if (r.type !== 'Income') { const k = r.who || t_('nobody'); per[k] = (per[k] || 0) + r.amount; } });
  const people = Object.keys(per).sort(function (x, y) { return per[y] - per[x]; });
  if (people.length > 1) {
    lines.push(t_('by_person'));
    people.forEach(function (k) { lines.push('👤 ' + esc_(k) + ': ' + money_(per[k])); });
    lines.push('');
  }
  hits.slice(0, SEARCH_MAX).forEach(function (r) { lines.push(entryLine_(r)); });
  if (hits.length > SEARCH_MAX) lines.push(t_('more', hits.length - SEARCH_MAX));
  return lines.join('\n');
}

/** Adds a category to the first empty spare row of its block on Settings. */
function addCategory_(name, type) {
  const all = getCategories_('Expense').concat(getCategories_('Income'));
  const same = all.filter(function (c) { return norm_(c.name) === norm_(name); })[0];
  if (same) {
    const sameType = getCategories_(type).some(function (c) { return c.name === same.name; });
    return sameType ? { ok: true, created: false, name: same.name } : { ok: false, message: t_('newcat_wrongtype', esc_(same.name)) };
  }
  const sh = sheet_('Settings'), h = headerMap_(sh, 4), last = sh.getLastRow();
  const vals = sh.getRange(5, 1, last - 4, sh.getLastColumn()).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (isVal_(vals[i][h['Type']], type) && vals[i][h['Category']] === '') {
      sh.getRange(5 + i, h['Category'] + 1).setValue(safeCell_(name));
      return { ok: true, created: true, name: name };
    }
  }
  return { ok: false, message: t_('newcat_full', t_(type === 'Income' ? 'w_income' : 'w_expense')) };
}

/* =========================================================================
 *  Automation: recurring items & bill reminders
 * ========================================================================= */
/** Adds every recurring occurrence that is due (up to today) and not yet posted. Returns how many. */
function postRecurring_() {
  const sh = sheet_('Recurring'), h = headerMap_(sh, 4);
  const last = sh.getLastRow(); if (last < 5) return 0;
  const values = sh.getRange(5, 1, last - 4, sh.getLastColumn()).getValues();
  const t = today_(); const posted = [], skipped = [], noRate = [];
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    values.forEach(function (v, i) {
      const name = v[h['Name']], amount = Number(v[h['Amount']]);
      const first = ymd_(v[h['First date']]);
      if (!name || !amount || !first || !isVal_(v[h['Active']], 'Yes')) return;
      const every = Number(v[h['Every (months)']]) || 1;
      const end = ymd_(v[h['End date (optional)']]);
      // If never posted, start from today (no back-filling of old months).
      const lastPosted = ymd_(v[h['Last auto-posted']]) || addDays_(t, -1);
      const cur = h['Currency'] === undefined ? 'EUR' : (String(v[h['Currency']] || 'EUR').trim().toUpperCase() || 'EUR');
      let newest = null;
      for (let k = 0; k < 1200; k += every) {
        const d = addMonths_(first, k);
        if (d > t || (end && d > end)) break;
        if (d <= lastPosted) continue;
        // Amount in pesos/dollars → € with that month's rate. No rate at all → wait (retried tomorrow).
        const conv = convert_(amount, cur, d);
        if (conv.missing) { noRate.push({ date: d, name: name }); break; }
        newest = d;
        // Already logged by hand (same category & amount within a few days)? Don't add it twice.
        if (findDuplicates_(v[h['Category']], conv.amount, d, RECURRING_DUP_DAYS).length) { skipped.push({ date: d, name: name }); continue; }
        const baseNote = wbLang_() === 'es' ? 'recurrente' : 'recurring';
        appendTransaction_({
          'Date': d, 'Type': v[h['Type']] || val_('Expense'), 'Category': v[h['Category']], 'Description': name,
          'Amount': conv.amount, 'For': v[h['For']], 'Paid by': v[h['Paid by']], 'Payment method': v[h['Payment method']],
          'Tax deductible': '', 'Notes': cur === 'EUR' ? baseNote : withMarker_(baseNote, conv), 'Source': val_('Auto'),
          'Currency': cur === 'EUR' ? '' : cur, 'Original amount': cur === 'EUR' ? '' : amount
        });
        posted.push({ date: d, name: name, amount: conv.amount, cur: cur, orig: amount });
      }
      if (newest) sh.getRange(5 + i, h['Last auto-posted'] + 1).setValue(toDate_(newest));
    });
  } finally { lock.releaseLock(); }
  if (posted.length || skipped.length || noRate.length) {
    broadcast_(function () {
      return t_('rec_title') +
        (posted.length ? '\n' + t_('rec_added') + '\n' + posted.map(function (p) {
          return showDate_(p.date) + ' ' + esc_(p.name) + ' – ' + (p.cur === 'EUR' ? money_(p.amount) : fmtCur_(p.orig, p.cur) + ' (' + money_(p.amount) + ')');
        }).join('\n') : '') +
        (skipped.length ? '\n' + t_('rec_skipped') + '\n' + skipped.map(function (p) { return showDate_(p.date) + ' ' + esc_(p.name); }).join('\n') : '') +
        (noRate.length ? '\n' + t_('rec_norate') + '\n' + noRate.map(function (p) { return showDate_(p.date) + ' ' + esc_(p.name); }).join('\n') : '');
    }, null, false);
  }
  return posted.length;
}

/** Reminds about pending bills: N days before, 1 day before, on the day, and weekly when overdue. */
function sendReminders_(force) {
  const before = Number(getSetting_('Remind bills this many days before', 3)) || 3;
  const t = today_(), sh = sheet_('Bills & Taxes'), h = headerMap_(sh, 4);
  const due = billRows_().filter(function (b) {
    if (!b.pending || !b.due) return false;
    const left = daysBetween_(t, b.due);
    if (force) return left <= before;
    return left === before || left === 1 || left === 0 || (left < 0 && (-left) % 7 === 0);
  });
  due.forEach(function (b) {
    const left = daysBetween_(t, b.due);
    // Everyone gets the reminder, each in their own language; only full-access users get the "Paid" button.
    broadcast_(function () {
      const when = left < 0 ? t_('overdue', -left) : left === 0 ? t_('due_today') : left === 1 ? t_('due_in1') : t_('due_in', left);
      return '🔔 <b>' + esc_(b.item) + '</b>\n' + showDate_(b.due, true) + ' – ' + when + (b.amount ? '\n' + t_('expected', money_(b.amount)) : '');
    }, b.amount ? function () {
      return { reply_markup: kb_([[btn_(t_('paid_btn', money_(b.amount)), 'b|' + b.row + '|' + b.due.replace(/-/g, ''))]]) };
    } : null, true);
    if (h['Reminded'] !== undefined) sh.getRange(b.row, h['Reminded'] + 1).setValue(toDate_(t));
  });
  return due.length;
}

function markBillPaid_(chatId, msgId, row, dueCompact) {
  const sh = sheet_('Bills & Taxes'), h = headerMap_(sh, 4);
  const v = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const due = ymd_(v[h['Due date']]);
  if (!due || due.replace(/-/g, '') !== dueCompact || !isVal_(v[h['Status']], 'Pending')) return edit_(chatId, msgId, t_('bill_changed'));
  const amount = Number(v[h['Expected amount']]);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    appendTransaction_({
      'Date': today_(), 'Type': val_('Expense'), 'Category': v[h['Category']], 'Description': v[h['Item']],
      'Amount': amount, 'For': getList_('For (person)')[0] || '', 'Paid by': '', 'Payment method': '', 'Tax deductible': '',
      'Notes': wbLang_() === 'es' ? 'desde Facturas e impuestos' : 'from Bills & Taxes', 'Source': val_('Bot')
    });
    sh.getRange(row, h['Status'] + 1).setValue(val_('Paid'));
  } finally { lock.releaseLock(); }
  edit_(chatId, msgId, t_('bill_paid', esc_(v[h['Item']]), money_(amount)));
}

/* =========================================================================
 *  Sheet access (English or Spanish workbook)
 * ========================================================================= */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
let WB_LANG_MEMO_ = null;
function wbLang_() {
  if (!WB_LANG_MEMO_) WB_LANG_MEMO_ = ss_().getSheetByName('Movimientos') ? 'es' : 'en';
  return WB_LANG_MEMO_;
}
/** A workbook word (sheet value) in the workbook's language, e.g. val_('Expense') → 'Gasto' in the Spanish file. */
function val_(en) { return wbLang_() === 'es' && WB_ES[en] ? WB_ES[en] : en; }
/** A setting label as written in the workbook. */
function label_(en) { return val_(en); }
/** Compares a cell to a workbook word in either language, ignoring case and accents ("si" = "Sí"). */
function isVal_(cell, en) {
  const c = norm_(cell);
  return c === norm_(en) || (WB_ES[en] !== undefined && c === norm_(WB_ES[en]));
}
function norm_(s) { return String(s === undefined || s === null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase(); }

function sheet_(name) {
  const sh = ss_().getSheetByName(name) || ss_().getSheetByName(WB_ES[name] || name);
  if (!sh) throw new Error('Sheet "' + val_(name) + '" not found – don\'t rename the tabs.');
  return sh;
}
/** Column positions by header name. Keys are the English names, whatever the workbook language. */
function headerMap_(sh, row) {
  const hdr = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const m = {};
  hdr.forEach(function (x, i) {
    const k = String(x).trim(); if (!k) return;
    const en = WB_REV[k] || k;
    if (m[en] === undefined) m[en] = i;
  });
  return m;
}
function appendTransaction_(obj) {
  const sh = sheet_('Transactions'), h = headerMap_(sh, 1), width = sh.getLastColumn();
  const row = new Array(width).fill('');
  Object.keys(obj).forEach(function (k) {
    if (h[k] === undefined) return;
    row[h[k]] = k === 'Date' ? toDate_(obj[k]) : safeCell_(obj[k]);
  });
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, width).setValues([row]);
  return r;
}
function transactions_() {
  const sh = sheet_('Transactions'), h = headerMap_(sh, 1), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues().map(function (v) {
    return {
      date: ymd_(v[h['Date']]), type: isVal_(v[h['Type']], 'Income') ? 'Income' : 'Expense', category: v[h['Category']],
      amount: Number(v[h['Amount']]) || 0, desc: String(v[h['Description']] || ''), notes: String(v[h['Notes']] || ''),
      who: String(v[h['For']] || ''), payer: String(v[h['Paid by']] || ''), source: String(v[h['Source']] || ''),
      currency: h['Currency'] === undefined ? '' : String(v[h['Currency']] || '').trim().toUpperCase(),
      original: h['Original amount'] === undefined ? 0 : (Number(v[h['Original amount']]) || 0),
      subcat: h['Subcategory'] === undefined ? '' : String(v[h['Subcategory']] || '')
    };
  }).filter(function (r) { return r.date; });
}
function billRows_() {
  const sh = sheet_('Bills & Taxes'), h = headerMap_(sh, 4), last = sh.getLastRow();
  if (last < 5) return [];
  return sh.getRange(5, 1, last - 4, sh.getLastColumn()).getValues().map(function (v, i) {
    return { row: 5 + i, item: v[h['Item']], due: ymd_(v[h['Due date']]), amount: Number(v[h['Expected amount']]) || 0, pending: isVal_(v[h['Status']], 'Pending') };
  }).filter(function (b) { return b.item; });
}
function recurringRows_() {
  const sh = sheet_('Recurring'), h = headerMap_(sh, 4), last = sh.getLastRow();
  if (last < 5) return [];
  return sh.getRange(5, 1, last - 4, sh.getLastColumn()).getValues().map(function (v) {
    return {
      name: v[h['Name']], type: isVal_(v[h['Type']], 'Income') ? 'Income' : 'Expense', amount: Number(v[h['Amount']]) || 0,
      every: Number(v[h['Every (months)']]) || 1, first: ymd_(v[h['First date']]),
      end: ymd_(v[h['End date (optional)']]), active: isVal_(v[h['Active']], 'Yes'),
      currency: h['Currency'] === undefined ? 'EUR' : (String(v[h['Currency']] || 'EUR').trim().toUpperCase() || 'EUR')
    };
  }).filter(function (r) { return r.name; });
}
/** Settings categories table (header row 4: Category | Type | Monthly budget / expected). type = 'Expense' | 'Income' */
function getCategories_(type) {
  const sh = sheet_('Settings'), h = headerMap_(sh, 4), last = sh.getLastRow();
  return sh.getRange(5, 1, last - 4, sh.getLastColumn()).getValues()
    .filter(function (v) { return v[h['Category']] !== '' && isVal_(v[h['Type']], type); })
    .map(function (v) { return { name: String(v[h['Category']]), budget: Number(v[h['Monthly budget / expected']]) || 0 }; });
}
/** Any simple list on Settings, by its header in row 4 (e.g. "Paid by"). */
function getList_(headerName) {
  const sh = sheet_('Settings'), h = headerMap_(sh, 4), last = sh.getLastRow();
  if (h[headerName] === undefined) return [];
  return sh.getRange(5, h[headerName] + 1, last - 4, 1).getValues()
    .map(function (v) { return String(v[0]).trim(); }).filter(String);
}
/** A "label → value" setting anywhere on the Settings sheet (label in either language). */
function getSetting_(label, dflt) {
  const vals = sheet_('Settings').getDataRange().getValues();
  const want = [norm_(label), norm_(WB_ES[label] || label)];
  for (let r = 0; r < vals.length; r++) for (let c = 0; c < vals[r].length - 1; c++)
    if (want.indexOf(norm_(vals[r][c])) >= 0) return vals[r][c + 1] === '' ? dflt : vals[r][c + 1];
  return dflt;
}
function idsFor_(label) {
  return String(getSetting_(label, '')).split(/[,;\s]+/)
    .map(function (x) { return x.replace(/\.0+$/, '').trim(); }).filter(String);
}
function taxCategory_() {
  const cats = getCategories_('Expense').map(function (c) { return c.name; });
  return cats.filter(function (c) { return isVal_(c, 'Taxes'); })[0] ||
    cats.filter(function (c) { return /tax|impuest/i.test(c); })[0] || val_('Taxes');
}

/* =========================================================================
 *  Telegram helpers
 * ========================================================================= */
/* ---------- 📤 Export: a month or a year as a CSV file, sent to the chat ----------
 *  Off unless Settings ▸ "Allow export from Telegram" = Yes, and only for full-access users. */
function exportAllowed_(role) {
  return role === 'full' && isVal_(getSetting_('Allow export from Telegram', 'No'), 'Yes');
}
function exportRows_() {
  const t = today_().split('-').map(Number), items = [], rows = [];
  let y = t[0], m = t[1];
  for (let i = 0; i < 6; i++) {
    items.push(btn_(TXT[LANG_].mon[m - 1] + ' ' + y, 'ex|m|' + y + '-' + pad_(m)));
    m--; if (m === 0) { m = 12; y--; }
  }
  for (let i = 0; i < items.length; i += 3) rows.push(items.slice(i, i + 3));
  rows.push([btn_(t_('export_year', t[0]), 'ex|y|' + t[0]), btn_(t_('export_year', t[0] - 1), 'ex|y|' + (t[0] - 1))]);
  return rows;
}
function exportCallback_(chatId, msgId, parts) {
  if (!exportAllowed_('full')) return edit_(chatId, msgId, t_('export_off', val_('Allow export from Telegram'), val_('Settings')));
  const key = String(parts[2] || '');
  if (!/^\d{4}(-\d{2})?$/.test(key)) return;
  const label = key.length === 4 ? key : TXT[LANG_].months[Number(key.slice(5)) - 1] + ' ' + key.slice(0, 4);
  const sh = sheet_('Transactions'), h = headerMap_(sh, 1), last = sh.getLastRow(), lastCol = sh.getLastColumn();
  const hdr = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const vals = last < 2 ? [] : sh.getRange(2, 1, last - 1, lastCol).getValues();
  const rows = vals.map(function (v) { return { d: ymd_(v[h['Date']]), v: v }; })
    .filter(function (r) { return r.d && r.d.indexOf(key) === 0; })
    .sort(function (a, b) { return a.d < b.d ? -1 : a.d > b.d ? 1 : 0; });
  if (!rows.length) return edit_(chatId, msgId, t_('export_none', label));
  let spent = 0, received = 0;
  rows.forEach(function (r) {
    const a = Number(r.v[h['Amount']]) || 0;
    if (isVal_(r.v[h['Type']], 'Income')) received += a; else spent += a;
  });
  const sep = LANG_ === 'es' ? ';' : ',';
  const lines = [hdr.map(function (x) { return csvCell_(x, sep); }).join(sep)].concat(rows.map(function (r) {
    return r.v.map(function (x, i) { return csvCell_(i === h['Date'] ? r.d : x, sep); }).join(sep);
  }));
  const name = t_('export_file') + '_' + key + '.csv';
  const blob = Utilities.newBlob('﻿' + lines.join('\r\n'), 'text/csv', name);
  const res = sendFile_(chatId, blob, t_('export_caption', label, rows.length, money_(spent), money_(received)));
  if (!res || !res.ok) return edit_(chatId, msgId, t_('export_fail', esc_(String((res && res.description) || '?'))));
  return edit_(chatId, msgId, t_('export_sent', label, rows.length));
}
/** One CSV cell: dates as yyyy-mm-dd, decimal comma in Spanish, quotes when needed, no formulas. */
function csvCell_(x, sep) {
  let s;
  if (x instanceof Date) s = ymd_(x);
  else if (typeof x === 'number') s = LANG_ === 'es' ? String(x).replace('.', ',') : String(x);
  else {
    s = String(x === undefined || x === null ? '' : x);
    if (/^[=+\-@]/.test(s)) s = "'" + s;
  }
  return (s.indexOf(sep) >= 0 || /["\r\n]/.test(s)) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
function sendFile_(chatId, blob, caption) {
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendDocument', {
    method: 'post', payload: { chat_id: String(chatId), caption: caption, parse_mode: 'HTML', document: blob }, muteHttpExceptions: true
  });
  try { return JSON.parse(res.getContentText()); } catch (e) { return { ok: false, description: res.getContentText() }; }
}

function tg_(method, payload) {
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) throw new Error('Bot not connected – use the menu ▸ 1. Connect Telegram bot');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  try { return JSON.parse(res.getContentText()); } catch (e) { return { ok: false, description: res.getContentText() }; }
}
function send_(chatId, text, extra) {
  return tg_('sendMessage', Object.assign({ chat_id: chatId, text: text, parse_mode: 'HTML' }, extra || {}));
}
function edit_(chatId, msgId, text, keyboardRows) {
  const p = { chat_id: chatId, message_id: msgId, text: text, parse_mode: 'HTML' };
  if (keyboardRows) p.reply_markup = kb_(keyboardRows);
  return tg_('editMessageText', p);
}
/** Sends a question with inline buttons and remembers which message holds them. */
function ask_(chatId, s, text, rows, noCancel) {
  const kb = noCancel ? rows : rows.concat([[btn_(t_('cancel_btn'), 'x|1')]]);
  const res = send_(chatId, text, { reply_markup: kb_(kb) });
  s.msgId = res && res.result ? res.result.message_id : null;
  saveState_(chatId, s);
}
/** Sends to every full-access user (with buttons) and, if includeReadOnly, to read-only users (without buttons) –
 *  each in their own language. textFn / extraFn are called once per person. */
function broadcast_(textFn, extraFn, includeReadOnly) {
  const before = LANG_;
  idsFor_('Allowed Telegram user IDs').forEach(function (id) {
    LANG_ = userLang_(id); send_(id, textFn(), extraFn ? extraFn() : undefined);
  });
  if (includeReadOnly) idsFor_('Read-only Telegram user IDs').forEach(function (id) {
    LANG_ = userLang_(id); send_(id, textFn());
  });
  LANG_ = before;
}
function kb_(rows) { return { inline_keyboard: rows }; }
function btn_(text, data) { return { text: text, callback_data: data }; }
function grid_(items, prefix) {
  const rows = [];
  items.forEach(function (it, i) {
    if (i % 2 === 0) rows.push([]);
    rows[rows.length - 1].push(btn_(it, prefix + '|' + i));
  });
  return rows;
}
function title_(s) { return t_(s.flow === 'income' ? 'title_income' : s.flow === 'tax' ? 'title_tax' : 'title_expense'); }
function getState_(chatId) { return JSON.parse(CacheService.getScriptCache().get('st_' + chatId) || 'null'); }
function saveState_(chatId, s) { CacheService.getScriptCache().put('st_' + chatId, JSON.stringify(s), STATE_TTL); }
function clearState_(chatId) { CacheService.getScriptCache().remove('st_' + chatId); }

/* =========================================================================
 *  Dates, numbers, text
 *  Dates travel inside the bot as 'yyyy-MM-dd' strings (no time, no time zone),
 *  and are converted to real dates only when written, using the SHEET's time zone.
 * ========================================================================= */
let TZ_MEMO_ = null;
function tz_() {
  if (TZ_MEMO_) return TZ_MEMO_;
  try { TZ_MEMO_ = ss_().getSpreadsheetTimeZone(); } catch (e) { }
  if (!TZ_MEMO_) { try { TZ_MEMO_ = Session.getScriptTimeZone(); } catch (e) { } }
  return TZ_MEMO_ || 'Europe/Madrid';
}
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
  return '';
}
function toDate_(ymd) { return Utilities.parseDate(ymd, tz_(), 'yyyy-MM-dd'); }
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function fromParts_(y, m, d) { return y + '-' + pad_(m) + '-' + pad_(d); }
function addDays_(ymd, n) {
  const p = ymd.split('-').map(Number), d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return fromParts_(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
}
/** Same day k months later; clamps 31st → last day of shorter months (like EDATE). */
function addMonths_(ymd, k) {
  const p = ymd.split('-').map(Number), total = p[1] - 1 + k;
  const y = p[0] + Math.floor(total / 12), m = ((total % 12) + 12) % 12 + 1;
  return fromParts_(y, m, Math.min(p[2], new Date(Date.UTC(y, m, 0)).getUTCDate()));
}
function nextMonth_(ymd) { return addMonths_(ymd.slice(0, 8) + '01', 1); }
function daysBetween_(a, b) {
  const pa = a.split('-').map(Number), pb = b.split('-').map(Number);
  return Math.round((Date.UTC(pb[0], pb[1] - 1, pb[2]) - Date.UTC(pa[0], pa[1] - 1, pa[2])) / 86400000);
}
/** "8 Oct 2026" / "8 oct 2026" – the month is always written as a word, so day/month can never be confused.
 *  long = true adds the weekday: "Thu 8 Oct 2026" / "jue 8 oct 2026". */
function showDate_(ymd, long) {
  if (!ymd) return '';
  const y = +ymd.slice(0, 4), m = +ymd.slice(5, 7), d = +ymd.slice(8, 10);
  let s = d + ' ' + t_('mon')[m - 1] + ' ' + y;
  if (long) {
    const dow = (new Date(Date.UTC(y, m - 1, d)).getUTCDay() + 6) % 7;
    s = (LANG_ === 'es' ? ['lun', 'mar', 'mié', 'jue', 'vie', 'sáb', 'dom'] : ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'])[dow] + ' ' + s;
  }
  return s;
}
function monthName_(ymd) { return t_('months')[+ymd.slice(5, 7) - 1] + ' ' + ymd.slice(0, 4); }
function cap_(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
/** Typed dates are always DAY FIRST: 25/10, 25/10/26, 25-10-2026, 25.10.2026. Also accepts 2026-10-25. */
function parseDateText_(t) {
  const iso = String(t).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  let d, mo, y;
  if (iso) { y = +iso[1]; mo = +iso[2]; d = +iso[3]; }
  else {
    const m = String(t).trim().match(/^(\d{1,2})[\/.\-](\d{1,2})(?:[\/.\-](\d{2}|\d{4}))?$/);
    if (!m) return null;
    d = +m[1]; mo = +m[2]; y = m[3] ? +m[3] : +today_().slice(0, 4);
    if (y < 100) y += 2000;
  }
  if (mo < 1 || mo > 12 || d < 1 || d > new Date(Date.UTC(y, mo, 0)).getUTCDate()) return null;
  return fromParts_(y, mo, d);
}
/** Accepts 45.20, 45,20, 1250, 1.250,50, 1,250.50, €45 */
function parseAmount_(t) {
  let s = String(t).replace(/[^\d.,]/g, '');
  if (!s) return null;
  const lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
  if (lc >= 0 && ld >= 0) s = lc > ld ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (lc >= 0) s = /,\d{1,2}$/.test(s) && s.split(',').length === 2 ? s.replace(',', '.') : s.replace(/,/g, '');
  else if (ld >= 0 && !(/\.\d{1,2}$/.test(s) && s.split('.').length === 2)) s = s.replace(/\./g, '');
  const n = Math.round(parseFloat(s) * 100) / 100;
  return isFinite(n) && n > 0 ? n : null;
}
/** 1,234.56 € (English) · 1.234,56 € (Spanish) */
function money_(n) {
  const sym = String(getSettingCached_('Currency symbol', '€'));
  const neg = n < 0, v = Math.abs(Math.round(n * 100) / 100).toFixed(2).split('.');
  const th = LANG_ === 'es' ? '.' : ',', dec = LANG_ === 'es' ? ',' : '.';
  return (neg ? '-' : '') + v[0].replace(/\B(?=(\d{3})+(?!\d))/g, th) + dec + v[1] + (sym ? ' ' + sym : '');
}
let SETTINGS_MEMO_ = {};
function getSettingCached_(label, dflt) {
  if (!(label in SETTINGS_MEMO_)) SETTINGS_MEMO_[label] = getSetting_(label, dflt);
  return SETTINGS_MEMO_[label];
}
/** Text starting with = + - @ would be read by Sheets as a formula: store it as plain text instead. */
function safeCell_(v) {
  if (v === undefined || v === null) return '';
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}
function esc_(s) { return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
