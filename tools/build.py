"""Builds the Family Budget workbook.  usage: build.py OUT.xlsx [--demo] [--es]"""
import sys, random, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter as CL
from howto import HOWTO

OUT = sys.argv[1]
DEMO = '--demo' in sys.argv
ES = '--es' in sys.argv
def L(en, es): return es if ES else en

# ---------------------------------------------------------------- vocabulary
S_DASH, S_HOW, S_ANA, S_YEAR, S_SEARCH = L('Dashboard', 'Panel'), L('How to use', 'Cómo usar'), L('Analysis', 'Análisis'), L('Year', 'Año'), L('Search', 'Buscar')
S_TX, S_REC, S_BILLS, S_SET = L('Transactions', 'Movimientos'), L('Recurring', 'Recurrentes'), L('Bills & Taxes', 'Facturas e impuestos'), L('Settings', 'Configuración')
def q(n): return "'" + n + "'"
EXP, INC, YES, NO = L('Expense', 'Gasto'), L('Income', 'Ingreso'), L('Yes', 'Sí'), 'No'
PEND, PAID, SKIP = L('Pending', 'Pendiente'), L('Paid', 'Pagado'), L('Skipped', 'Omitido')
ALLYEAR, ENDED = L('All year', 'Todo el año'), L('Ended', 'Finalizado')
MANUAL, BOT, AUTO = 'Manual', 'Bot', 'Auto'
MONTHS = L(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
           ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'])

INCOME = [  # en, es, demo budget, help en, help es
    ('Salary – Adult 1', 'Sueldo – Adulto 1', 2900, 'Net salary of the first adult', 'Sueldo neto del primer adulto'),
    ('Salary – Adult 2', 'Sueldo – Adulto 2', 2150, 'Net salary of the second adult', 'Sueldo neto del segundo adulto'),
    ('Rental income', 'Alquiler cobrado', 750, 'Rent you receive from a property', 'Alquiler que cobras por una vivienda'),
    ('Child benefit', 'Ayuda por hijos', 250, 'Public child allowances', 'Ayudas públicas por hijos'),
    ('Side income', 'Ingresos extra', 150, 'Freelance work, sales, extras', 'Trabajos puntuales, ventas, extras'),
    ('Investments & Crypto', 'Inversiones y cripto', 0, 'Dividends, interest, crypto sales', 'Dividendos, intereses, ventas de cripto'),
    ('Other income', 'Otros ingresos', 0, 'Anything else that comes in', 'Cualquier otro ingreso'),
]
EXPENSES = [
    ('Housing', 'Vivienda', 1200, 'Rent or mortgage, community fees', 'Alquiler o hipoteca, comunidad'),
    ('Utilities', 'Suministros', 300, 'Electricity, gas, water, internet, phones', 'Luz, gas, agua, internet, teléfonos'),
    ('Groceries', 'Supermercado', 650, 'Supermarket, market, bakery', 'Supermercado, mercado, panadería'),
    ('Kids – School', 'Hijos – Colegio', 180, 'Fees, canteen, books, uniforms, trips', 'Cuotas, comedor, libros, uniformes, excursiones'),
    ('Kids – Childcare', 'Hijos – Guardería', 450, 'Nursery, babysitter, holiday camps', 'Guardería, canguro, campamentos'),
    ('Kids – Activities', 'Hijos – Actividades', 120, 'Sports, music, languages', 'Deportes, música, idiomas'),
    ('Kids – Clothes & Other', 'Hijos – Ropa y otros', 100, 'Clothes, shoes, toys, birthdays', 'Ropa, calzado, juguetes, cumpleaños'),
    ('Taxes', 'Impuestos', 250, 'Income, property, vehicle and local taxes', 'Renta, inmuebles, vehículos y tasas locales'),
    ('Insurance', 'Seguros', 120, 'Home, car, health, life', 'Hogar, coche, salud, vida'),
    ('Transport', 'Transporte', 220, 'Fuel, public transport, parking, car repairs', 'Combustible, transporte público, aparcamiento, taller'),
    ('Health', 'Salud', 80, 'Pharmacy, dentist, glasses', 'Farmacia, dentista, gafas'),
    ('Subscriptions', 'Suscripciones', 70, 'Streaming, apps, cloud, gym', 'Streaming, apps, nube, gimnasio'),
    ('Home & Maintenance', 'Hogar y mantenimiento', 100, 'Repairs, furniture, appliances', 'Reparaciones, muebles, electrodomésticos'),
    ('Leisure & Eating out', 'Ocio y restaurantes', 300, 'Restaurants, outings, holidays', 'Restaurantes, salidas, vacaciones'),
    ('Personal & Gifts', 'Personal y regalos', 120, 'Hairdresser, clothes for adults, presents', 'Peluquería, ropa de adultos, regalos'),
    ('Other', 'Otros', 50, 'Anything that does not fit elsewhere', 'Lo que no encaja en otra categoría'),
]
def cat(en):
    for row in INCOME + EXPENSES:
        if row[0] == en: return L(row[0], row[1])
    raise KeyError(en)
PERSONS = L(['Family', 'Child 1', 'Child 2', 'Adult 1', 'Adult 2'], ['Familia', 'Hijo/a 1', 'Hijo/a 2', 'Adulto 1', 'Adulto 2'])
PAYERS = L(['Adult 1', 'Adult 2', 'Joint account'], ['Adulto 1', 'Adulto 2', 'Cuenta conjunta'])
METHODS = L(['Bank transfer', 'Direct debit', 'Debit card', 'Credit card', 'Cash', 'Crypto'],
            ['Transferencia', 'Domiciliación', 'Tarjeta de débito', 'Tarjeta de crédito', 'Efectivo', 'Cripto'])
TAXTYPES = L(['Income tax', 'Property tax', 'Vehicle tax', 'Waste tax', 'Social contributions', 'Other tax'],
             ['Impuesto sobre la renta', 'Impuesto de bienes inmuebles', 'Impuesto de vehículos', 'Tasa de basuras', 'Cotizaciones sociales', 'Otro impuesto'])
FREQS = [(L('Monthly', 'Mensual'), 1), (L('Every 2 months', 'Cada 2 meses'), 2), (L('Quarterly', 'Trimestral'), 3),
         (L('Every 6 months', 'Semestral'), 6), (L('Yearly', 'Anual'), 12)]
def P(i): return PERSONS[i]
def PB(i): return PAYERS[i]
def M(i): return METHODS[i]
def F(i): return FREQS[i][0]

# ---------------------------------------------------------------- style
NAVY, TEAL, CORAL, SLATE, GREY, LIGHT, INPUT, PAPER, GREEN = '1F3A5F', '2A9D8F', 'E76F51', '264653', '7F7F7F', 'EEF2F7', 'FFF4C2', 'F7F9FC', 'E6F4EA'
MONEY = '#,##0.00;[Red]-#,##0.00;"–"'
PCT = '0.0%;[Red]-0.0%;"–"'
DATEF = 'dd mmm yyyy'
FONT = 'Arial'
def fill(c): return PatternFill('solid', start_color=c, end_color=c)
def font(**k): k.setdefault('name', FONT); k.setdefault('size', 10); return Font(**k)
THIN = Side(style='thin', color='D0D7E2')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical='center')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

def banner(ws, title, sub, ncols):
    for r in (1, 2):
        for c in range(1, ncols + 1):
            ws.cell(r, c).fill = fill(NAVY)
    ws.cell(1, 1, title).font = font(size=16, bold=True, color='FFFFFF')
    ws.cell(2, 1, sub).font = font(size=10, italic=True, color='DCE3EC')
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18

def header(ws, row, c1, labels, color=NAVY, height=None):
    for i, t in enumerate(labels):
        c = ws.cell(row, c1 + i, t)
        c.font = font(bold=True, color='FFFFFF'); c.fill = fill(color); c.alignment = CENTER; c.border = BOX
    if height: ws.row_dimensions[row].height = height

def section(ws, row, col, text):
    ws.cell(row, col, text).font = font(size=12, bold=True, color=SLATE)

def inp(c, fmt=None):
    c.fill = fill(INPUT); c.font = font(color='1A4FB4', bold=True); c.border = BOX; c.alignment = Alignment(horizontal='center', vertical='center')
    if fmt: c.number_format = fmt

def tile(ws, row, col, label, formula, fmt=MONEY, color=TEAL):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
    a = ws.cell(row, col, label); a.font = font(size=9, bold=True, color='FFFFFF'); a.fill = fill(color); a.alignment = CENTER
    ws.cell(row, col + 1).fill = fill(color)
    b = ws.cell(row + 1, col, formula); b.font = font(size=15, bold=True, color=SLATE); b.alignment = CENTER; b.number_format = fmt
    b.fill = fill(LIGHT); ws.cell(row + 1, col + 1).fill = fill(LIGHT)
    ws.row_dimensions[row + 1].height = 30

def view(ws, color):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = color
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

def widths(ws, d):
    for k, v in d.items(): ws.column_dimensions[k].width = v

def hide(ws, c1, c2):
    for c in range(c1, c2 + 1): ws.column_dimensions[CL(c)].hidden = True

def dv(ws, formula, rng, warn=True):
    v = DataValidation(type='list', formula1=formula, allow_blank=True)
    v.errorStyle = 'warning' if warn else 'stop'
    v.showErrorMessage = True
    v.error = L('This value is not in the list in Settings.', 'Este valor no está en la lista de Configuración.')
    v.errorTitle = L('Not in the list', 'No está en la lista')
    ws.add_data_validation(v); v.add(rng)

def dv_date(ws, rng):
    v = DataValidation(type='date', operator='greaterThan', formula1='36526', allow_blank=True)
    v.errorStyle = 'warning'; v.showErrorMessage = True
    v.error = L('Please enter a date, e.g. 15/03/2026.', 'Escribe una fecha, por ejemplo 15/03/2026.')
    ws.add_data_validation(v); v.add(rng)

def setref(col, r1, r2): return q(S_SET) + '!$' + col + '$' + str(r1) + ':$' + col + '$' + str(r2)

NTX = 20001      # formulas look at Transactions rows 2..NTX (same last row everywhere)
TX_ROWS = NTX - 1   # formatted rows 2..NTX, so the imported sheet has enough rows for every formula
def TX(col): return q(S_TX) + '!$' + col + '$2:$' + col + '$' + str(NTX)
TXA, TXB, TXC, TXE, TXF, TXG = TX('A'), TX('B'), TX('C'), TX('E'), TX('F'), TX('G')
def sum_period(start, end, extra=''):
    return 'SUMIFS(' + TXE + ',' + TXA + ',">="&' + start + ',' + TXA + ',"<"&' + end + extra + ')'

wb = Workbook()
wb.remove(wb.active)
order = [S_DASH, S_HOW, S_ANA, S_YEAR, S_SEARCH, S_TX, S_REC, S_BILLS, S_SET]
WS = {n: wb.create_sheet(n) for n in order}

# ================================================================ Settings
st = WS[S_SET]
banner(st, L('Settings', 'Configuración'), L('Your categories, people, lists and bot settings. Yellow cells are yours to fill in.',
       'Tus categorías, personas, listas y ajustes del bot. Las celdas amarillas son para rellenar.'), 24)
section(st, 3, 1, L('Categories – income first (green rows), then expenses', 'Categorías – primero ingresos (filas verdes), luego gastos'))
section(st, 3, 6, L('Lists', 'Listas'))
section(st, 3, 14, L('Telegram bot', 'Bot de Telegram'))
section(st, 3, 18, L('Fixed lists (do not edit)', 'Listas fijas (no editar)'))
header(st, 4, 1, [L('Category', 'Categoría'), L('Type', 'Tipo'), L('Monthly budget / expected', 'Presupuesto / previsto mensual'), L('What goes here', 'Qué va aquí')], height=32)
INC_ROWS, EXP_ROWS = (5, 16), (17, 56)
cats = [(r, INC) for r in range(INC_ROWS[0], INC_ROWS[1] + 1)] + [(r, EXP) for r in range(EXP_ROWS[0], EXP_ROWS[1] + 1)]
for i, (r, typ) in enumerate(cats):
    src = INCOME if typ == INC else EXPENSES
    k = r - (INC_ROWS[0] if typ == INC else EXP_ROWS[0])
    item = src[k] if k < len(src) else None
    a = st.cell(r, 1, L(item[0], item[1]) if item else None)
    b = st.cell(r, 2, typ)
    c = st.cell(r, 3, item[2] if (item and DEMO and item[2]) else None)
    d = st.cell(r, 4, L(item[3], item[4]) if item else None)
    for x in (a, b, c, d):
        x.border = BOX; x.font = font()
        x.fill = fill(GREEN if typ == INC else 'FFFFFF')
    a.fill = fill(INPUT) if typ == EXP else fill('D4EDDA'); a.font = font(color='1A4FB4', bold=True)
    c.fill = fill(INPUT) if typ == EXP else fill('D4EDDA'); c.font = font(color='1A4FB4'); c.number_format = MONEY
    b.font = font(color=GREY); d.font = font(color=GREY, italic=True, size=9)
lists = [(6, L('For (person)', 'Para (persona)'), PERSONS, 12), (7, L('Paid by', 'Pagado por'), PAYERS, 8),
         (8, L('Payment method', 'Forma de pago'), METHODS, 12), (9, L('Tax types', 'Tipos de impuesto'), TAXTYPES, 12)]
for col, title, values, slots in lists:
    header(st, 4, col, [title])
    for i in range(slots):
        c = st.cell(5 + i, col, values[i] if i < len(values) else None); inp(c); c.alignment = Alignment(vertical='center')
header(st, 4, 11, [L('Frequency', 'Frecuencia'), L('Months', 'Meses')])
for i, (n, m) in enumerate(FREQS):
    st.cell(5 + i, 11, n).border = BOX; st.cell(5 + i, 12, m).border = BOX
header(st, 4, 14, [L('Bot setting', 'Ajuste del bot'), L('Value', 'Valor')])
bot = [(L('Allowed Telegram user IDs', 'IDs de Telegram con acceso total'), None),
       (L('Read-only Telegram user IDs', 'IDs de Telegram solo lectura'), None),
       (L('Currency symbol', 'Símbolo de moneda'), '€'),
       (L('Remind bills this many days before', 'Avisar de facturas con estos días de antelación'), 3),
       (L('Auto-post recurring items', 'Registrar recurrentes automáticamente'), YES),
       (L('Allow export from Telegram', 'Permitir exportar desde Telegram'), NO)]
for i, (lab, val) in enumerate(bot):
    a = st.cell(5 + i, 14, lab); a.font = font(bold=True); a.border = BOX; a.alignment = WRAP
    b = st.cell(5 + i, 15, val); inp(b)
st['O5'].number_format = '@'; st['O6'].number_format = '@'
dv(st, setref('V', 5, 6), 'O9:O10')
notes = L(['Telegram IDs: numbers separated by commas, e.g. 123456789, 987654321.',
           'Leave them empty at first: the bot tells you your ID when you send /start.',
           'Full access = can add and edit. Read-only = can only search and look.',
           'Tip: protect these two cells (Data ▸ Protect sheets and ranges).',
           'Export from Telegram: Yes = full-access users get a 📤 Export button (CSV file). Read-only users never do.'],
          ['IDs de Telegram: números separados por comas, p. ej. 123456789, 987654321.',
           'Déjalos vacíos al principio: el bot te dice tu ID cuando le envías /start.',
           'Acceso total = puede añadir y editar. Solo lectura = solo puede buscar y consultar.',
           'Consejo: protege estas dos celdas (Datos ▸ Proteger hojas e intervalos).',
           'Exportar desde Telegram: Sí = los usuarios con acceso total ven un botón 📤 Exportar (archivo CSV). Los de solo lectura, nunca.'])
for i, t in enumerate(notes):
    st.cell(12 + i, 14, t).font = font(size=9, italic=True, color=GREY)
fixed = [(18, L('Months', 'Meses'), MONTHS), (19, L('Type', 'Tipo'), [EXP, INC]), (20, L('Source', 'Origen'), [MANUAL, BOT, AUTO]),
         (21, L('Status', 'Estado'), [PEND, PAID, SKIP]), (22, L('Yes/No', 'Sí/No'), [YES, NO]),
         (23, L('Years', 'Años'), list(range(2024, 2041))), (24, L('Period', 'Periodo'), [ALLYEAR] + MONTHS)]
for col, title, values in fixed:
    header(st, 4, col, [title], color=GREY)
    for i, v in enumerate(values):
        c = st.cell(5 + i, col, v); c.font = font(color=GREY); c.border = BOX
widths(st, {'A': 26, 'B': 10, 'C': 16, 'D': 40, 'E': 3, 'F': 16, 'G': 16, 'H': 18, 'I': 26, 'J': 3, 'K': 16, 'L': 8, 'M': 3,
            'N': 34, 'O': 26, 'P': 3, 'Q': 3, 'R': 12, 'S': 10, 'T': 10, 'U': 11, 'V': 8, 'W': 8, 'X': 13})
st.freeze_panes = 'A5'
st.sheet_properties.tabColor = GREY

# ================================================================ Transactions
tx = WS[S_TX]
TXH = [L('Date', 'Fecha'), L('Type', 'Tipo'), L('Category', 'Categoría'), L('Description', 'Descripción'), L('Amount', 'Importe'),
       L('For', 'Para'), L('Paid by', 'Pagado por'), L('Payment method', 'Forma de pago'), L('Tax deductible', 'Deducible'),
       L('Notes', 'Notas'), L('Source', 'Origen')]
header(tx, 1, 1, TXH, height=26)
widths(tx, {'A': 13, 'B': 10, 'C': 24, 'D': 36, 'E': 13, 'F': 13, 'G': 15, 'H': 17, 'I': 11, 'J': 30, 'K': 9})
for r in range(2, TX_ROWS + 2):
    tx.cell(r, 1).number_format = DATEF
    tx.cell(r, 5).number_format = MONEY
end = TX_ROWS + 1
dv_date(tx, 'A2:A%d' % end)
dv(tx, setref('S', 5, 6), 'B2:B%d' % end)
dv(tx, setref('A', 5, 56), 'C2:C%d' % end)
dv(tx, setref('F', 5, 16), 'F2:F%d' % end)
dv(tx, setref('G', 5, 12), 'G2:G%d' % end)
dv(tx, setref('H', 5, 16), 'H2:H%d' % end)
dv(tx, setref('V', 5, 6), 'I2:I%d' % end)
dv(tx, setref('T', 5, 7), 'K2:K%d' % end)
tx.conditional_formatting.add('A2:K%d' % end, FormulaRule(formula=['$B2="%s"' % INC], font=Font(color='1E7F55')))
tx.freeze_panes = 'A2'
tx.auto_filter.ref = 'A1:K%d' % end
tx.sheet_properties.tabColor = TEAL

# ---------------------------------------------------------------- demo data
DESC = {  # en -> es
    'Salary': 'Nómina', 'Flat rental income': 'Alquiler del piso', 'Child benefit': 'Ayuda por hijos', 'Freelance job': 'Trabajo freelance',
    'Second-hand sale': 'Venta de segunda mano', 'Savings interest': 'Intereses de la cuenta', 'Rent': 'Alquiler',
    'Electricity': 'Luz', 'Water': 'Agua', 'Gas': 'Gas', 'Internet': 'Internet', 'Mobile phones': 'Móviles', 'Supermarket': 'Supermercado',
    'Weekly market': 'Mercado semanal', 'Bakery': 'Panadería', 'School canteen': 'Comedor escolar', 'School books': 'Libros del colegio',
    'School trip': 'Excursión del colegio', 'Nursery fee': 'Cuota de guardería', 'Summer camp': 'Campamento de verano',
    'Swimming club': 'Club de natación', 'Music lessons': 'Clases de música', 'Kids clothes': 'Ropa de niños', 'Shoes': 'Zapatos',
    'Birthday party': 'Fiesta de cumpleaños', 'Property tax': 'Impuesto de bienes inmuebles', 'Vehicle tax': 'Impuesto de circulación',
    'Waste tax': 'Tasa de basuras', 'Income tax return': 'Declaración de la renta', 'Car insurance': 'Seguro del coche',
    'Home insurance': 'Seguro del hogar', 'Health insurance': 'Seguro de salud', 'Fuel': 'Gasolina', 'Public transport pass': 'Abono de transporte',
    'Parking': 'Aparcamiento', 'Car service': 'Revisión del coche', 'Pharmacy': 'Farmacia', 'Dentist': 'Dentista', 'Streaming': 'Streaming',
    'Music app': 'App de música', 'Cloud storage': 'Almacenamiento en la nube', 'Gym': 'Gimnasio', 'Hardware store': 'Ferretería',
    'Plumber': 'Fontanero', 'New kettle': 'Hervidor nuevo', 'Restaurant': 'Restaurante', 'Cinema': 'Cine', 'Day trip': 'Excursión de un día',
    'Summer holiday': 'Vacaciones de verano', 'Hairdresser': 'Peluquería', 'Birthday present': 'Regalo de cumpleaños', 'Post office': 'Correos',
    'EXAMPLE – school canteen (delete this row)': 'EJEMPLO – comedor escolar (borra esta fila)', 'recurring': 'recurrente',
}
def D(en): return DESC[en] if ES else en
def by(i): return L('by ', 'por ') + PAYERS[i]

def demo_rows():
    rnd = random.Random(7)
    rows = []
    def add(d, typ, c, desc, amt, who, payer, method, tax='', note='', src=MANUAL):
        rows.append([d, typ, cat(c), D(desc), round(amt, 2), who, payer, method, tax, note, src])
    for m in range(1, 11):
        day = lambda x: dt.date(2026, m, x)
        # income
        add(day(28), INC, 'Salary – Adult 1', 'Salary', 2900, P(3), '', M(0), note=D('recurring'), src=AUTO)
        add(day(27), INC, 'Salary – Adult 2', 'Salary', 2150 if m != 6 else 3250, P(4), '', M(0), note=D('recurring'), src=AUTO)
        add(day(5), INC, 'Rental income', 'Flat rental income', 750, P(0), '', M(0), note=D('recurring'), src=AUTO)
        add(day(15), INC, 'Child benefit', 'Child benefit', 250, P(0), '', M(0), note=D('recurring'), src=AUTO)
        if m in (2, 5, 9): add(day(rnd.randint(8, 20)), INC, 'Side income', 'Freelance job', rnd.choice([220, 340, 480]), P(4), '', M(0))
        if m == 7: add(day(12), INC, 'Side income', 'Second-hand sale', 85, P(0), '', M(4), note=by(0), src=BOT)
        if m in (3, 6, 9): add(day(30), INC, 'Investments & Crypto', 'Savings interest', round(rnd.uniform(18, 30), 2), P(0), '', M(0))
        # fixed costs
        add(day(1), EXP, 'Housing', 'Rent', 1200, P(0), PB(2), M(1), note=D('recurring'), src=AUTO)
        add(day(rnd.randint(8, 12)), EXP, 'Utilities', 'Electricity', rnd.uniform(62, 118), P(0), PB(2), M(1))
        if m % 2 == 0: add(day(20), EXP, 'Utilities', 'Water', rnd.uniform(28, 41), P(0), PB(2), M(1), note=D('recurring'), src=AUTO)
        if m in (1, 2, 3, 10): add(day(14), EXP, 'Utilities', 'Gas', rnd.uniform(45, 90), P(0), PB(2), M(1))
        add(day(3), EXP, 'Utilities', 'Internet', 45, P(0), PB(2), M(1), note=D('recurring'), src=AUTO)
        add(day(3), EXP, 'Utilities', 'Mobile phones', 30, P(0), PB(0), M(1), note=D('recurring'), src=AUTO)
        # groceries
        for _ in range(rnd.randint(8, 11)):
            src = rnd.choice([MANUAL, BOT, BOT])
            add(day(rnd.randint(1, 28)), EXP, 'Groceries', rnd.choice(['Supermarket', 'Supermarket', 'Weekly market', 'Bakery']),
                rnd.uniform(12, 118), P(0), PB(rnd.randint(0, 1)), M(rnd.choice([2, 3, 4])), note=by(rnd.randint(0, 1)) if src == BOT else '', src=src)
        # kids
        if m not in (7, 8):
            add(day(5), EXP, 'Kids – School', 'School canteen', 85.5, P(1), PB(2), M(1), note=D('recurring'), src=AUTO)
            add(day(2), EXP, 'Kids – Childcare', 'Nursery fee', 450, P(2), PB(2), M(1), tax=YES, note=D('recurring'), src=AUTO)
            add(day(6), EXP, 'Kids – Activities', 'Swimming club', 40, P(1), PB(0), M(1), note=D('recurring'), src=AUTO)
            add(day(rnd.randint(7, 12)), EXP, 'Kids – Activities', 'Music lessons', 35, P(2), PB(1), M(2), note=by(1), src=BOT)
        if m == 9: add(day(4), EXP, 'Kids – School', 'School books', 186.4, P(1), PB(0), M(3), tax=YES)
        if m in (4, 10): add(day(rnd.randint(14, 24)), EXP, 'Kids – School', 'School trip', 25, rnd.choice([P(1), P(2)]), PB(1), M(4), note=by(1), src=BOT)
        if m == 7: add(day(1), EXP, 'Kids – Childcare', 'Summer camp', 520, P(1), PB(2), M(0), tax=YES)
        if m in (3, 5, 9, 10):
            add(day(rnd.randint(10, 25)), EXP, 'Kids – Clothes & Other', rnd.choice(['Kids clothes', 'Shoes']), rnd.uniform(35, 95),
                rnd.choice([P(1), P(2)]), PB(rnd.randint(0, 1)), M(2))
        if m in (2, 6): add(day(rnd.randint(10, 20)), EXP, 'Kids – Clothes & Other', 'Birthday party', rnd.uniform(70, 140), P(1 if m == 2 else 2), PB(2), M(2))
        # taxes & insurance
        if m == 4: add(day(28), EXP, 'Taxes', 'Vehicle tax', 112.5, P(0), PB(2), M(1))
        if m == 6:
            add(day(25), EXP, 'Taxes', 'Income tax return', 380, P(3), PB(0), M(0))
            add(day(30), EXP, 'Taxes', 'Property tax', 420, P(0), PB(2), M(1))
        if m == 10: add(day(20), EXP, 'Taxes', 'Waste tax', 96, P(0), PB(2), M(1))
        if m == 3: add(day(10), EXP, 'Insurance', 'Car insurance', 480, P(0), PB(2), M(1), note=D('recurring'), src=AUTO)
        if m == 9: add(day(15), EXP, 'Insurance', 'Home insurance', 300, P(0), PB(2), M(1), note=D('recurring'), src=AUTO)
        add(day(8), EXP, 'Insurance', 'Health insurance', 62, P(0), PB(2), M(1), tax=YES, note=D('recurring'), src=AUTO)
        # transport
        for _ in range(rnd.randint(2, 3)):
            add(day(rnd.randint(1, 28)), EXP, 'Transport', 'Fuel', rnd.uniform(48, 72), P(0), PB(rnd.randint(0, 1)), M(2), note=by(0), src=BOT)
        add(day(1), EXP, 'Transport', 'Public transport pass', 40, P(4), PB(1), M(2))
        if m in (2, 8): add(day(rnd.randint(5, 25)), EXP, 'Transport', 'Parking', rnd.uniform(8, 25), P(0), PB(0), M(2))
        if m == 5: add(day(17), EXP, 'Transport', 'Car service', 265, P(0), PB(2), M(3))
        # health, subscriptions, home
        if rnd.random() < 0.7: add(day(rnd.randint(1, 28)), EXP, 'Health', 'Pharmacy', rnd.uniform(8, 45), rnd.choice(PERSONS), PB(rnd.randint(0, 1)), M(2))
        if m in (3, 9): add(day(rnd.randint(5, 25)), EXP, 'Health', 'Dentist', 90 if m == 3 else 140, P(1) if m == 3 else P(3), PB(0), M(3), tax=YES)
        add(day(12), EXP, 'Subscriptions', 'Streaming', 15.99, P(0), PB(0), M(3), note=D('recurring'), src=AUTO)
        add(day(18), EXP, 'Subscriptions', 'Music app', 10.99, P(0), PB(1), M(3))
        add(day(21), EXP, 'Subscriptions', 'Cloud storage', 2.99, P(0), PB(0), M(3))
        add(day(4), EXP, 'Subscriptions', 'Gym', 39, P(3), PB(0), M(1))
        if m in (2, 4, 7, 10): add(day(rnd.randint(5, 25)), EXP, 'Home & Maintenance', 'Hardware store', rnd.uniform(15, 70), P(0), PB(0), M(2))
        if m == 4: add(day(9), EXP, 'Home & Maintenance', 'Plumber', 140, P(0), PB(2), M(0))
        if m == 8: add(day(22), EXP, 'Home & Maintenance', 'New kettle', 34.9, P(0), PB(1), M(2), note=by(1), src=BOT)
        # leisure & personal
        for _ in range(rnd.randint(2, 4)):
            add(day(rnd.randint(1, 28)), EXP, 'Leisure & Eating out', rnd.choice(['Restaurant', 'Restaurant', 'Cinema', 'Day trip']),
                rnd.uniform(18, 95), P(0), PB(rnd.randint(0, 1)), M(rnd.choice([2, 3])), note=by(0) if rnd.random() < .4 else '', src=BOT if rnd.random() < .4 else MANUAL)
        if m == 8: add(day(3), EXP, 'Leisure & Eating out', 'Summer holiday', 1450, P(0), PB(2), M(3))
        add(day(rnd.randint(1, 28)), EXP, 'Personal & Gifts', rnd.choice(['Hairdresser', 'Birthday present']), rnd.uniform(18, 60),
            rnd.choice([P(3), P(4)]), PB(rnd.randint(0, 1)), M(2))
        if m in (5, 11): add(day(9), EXP, 'Other', 'Post office', 6.4, P(0), PB(0), M(4))
    for r in rows:
        if r[9] == '' and r[10] == BOT: r[9] = by(0)
    rows.sort(key=lambda r: r[0])
    return rows

if DEMO:
    data = demo_rows()
else:
    data = [[dt.date(2026, 10, 1), EXP, cat('Kids – School'), D('EXAMPLE – school canteen (delete this row)'), 85.5, P(1), PB(2), M(1), '', '', MANUAL]]
for i, row in enumerate(data):
    for j, v in enumerate(row):
        c = tx.cell(2 + i, 1 + j, v if v != '' else None)
        if not DEMO: c.font = font(italic=True, color=GREY)

# ================================================================ Recurring
rc = WS[S_REC]
banner(rc, L('Recurring', 'Recurrentes'), L('Things that repeat (rent, salaries, fees). The bot can add them to Transactions automatically on their date.',
       'Lo que se repite (alquiler, sueldos, cuotas). El bot puede añadirlos solo a Movimientos en su fecha.'), 15)
RCH = [L('Name', 'Nombre'), L('Type', 'Tipo'), L('Category', 'Categoría'), L('Amount', 'Importe'), L('For', 'Para'), L('Paid by', 'Pagado por'),
       L('Payment method', 'Forma de pago'), L('Frequency', 'Frecuencia'), L('Every (months)', 'Cada (meses)'), L('First date', 'Primera fecha'),
       L('End date (optional)', 'Fecha fin (opcional)'), L('Active', 'Activo'), L('Next date', 'Próxima fecha'),
       L('Monthly equivalent', 'Equivalente mensual'), L('Last auto-posted', 'Último registro automático')]
header(rc, 4, 1, RCH, height=36)
R1, R2 = 5, 54
for r in range(R1, R2 + 1):
    for c in range(1, 16):
        x = rc.cell(r, c); x.border = BOX; x.font = font()
        if c in (1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12): x.fill = fill(INPUT)
    rc.cell(r, 4).number_format = MONEY; rc.cell(r, 14).number_format = MONEY
    for c in (10, 11, 13, 15): rc.cell(r, c).number_format = DATEF
    rc.cell(r, 9, '=IF(H{0}="","",IFERROR(INDEX({1},MATCH(H{0},{2},0)),""))'.format(r, setref('L', 5, 9), setref('K', 5, 9)))
    X = 'EDATE(J{0},CEILING(MAX(0,(YEAR(TODAY())-YEAR(J{0}))*12+MONTH(TODAY())-MONTH(J{0}))/I{0},1)*I{0})'.format(r)
    X2 = 'EDATE(J{0},(CEILING(MAX(0,(YEAR(TODAY())-YEAR(J{0}))*12+MONTH(TODAY())-MONTH(J{0}))/I{0},1)+1)*I{0})'.format(r)
    nd = 'IF(J{0}>=TODAY(),J{0},IF({1}>=TODAY(),{1},{2}))'.format(r, X, X2)
    rc.cell(r, 13, '=IF(OR(A{0}="",J{0}="",I{0}=""),"",IF(L{0}<>"{1}","",IF(AND(K{0}<>"",{2}>K{0}),"{3}",{2})))'.format(r, YES, nd, ENDED))
    rc.cell(r, 14, '=IF(OR(D{0}="",I{0}=""),"",D{0}/I{0})'.format(r))
    for c in (9, 13, 14, 15): rc.cell(r, c).font = font(color=GREY)
dv(rc, setref('S', 5, 6), 'B%d:B%d' % (R1, R2))
dv(rc, setref('A', 5, 56), 'C%d:C%d' % (R1, R2))
dv(rc, setref('F', 5, 16), 'E%d:E%d' % (R1, R2))
dv(rc, setref('G', 5, 12), 'F%d:F%d' % (R1, R2))
dv(rc, setref('H', 5, 16), 'G%d:G%d' % (R1, R2))
dv(rc, setref('K', 5, 9), 'H%d:H%d' % (R1, R2))
dv(rc, setref('V', 5, 6), 'L%d:L%d' % (R1, R2))
dv_date(rc, 'J%d:K%d' % (R1, R2))
tot = [(55, L('Fixed monthly income (monthly equivalent)', 'Ingresos fijos mensuales (equivalente mensual)'), INC),
       (56, L('Fixed monthly costs (monthly equivalent)', 'Gastos fijos mensuales (equivalente mensual)'), EXP)]
for r, lab, typ in tot:
    rc.cell(r, 1, lab).font = font(bold=True, color=SLATE)
    c = rc.cell(r, 14, '=SUMIFS(N{0}:N{1},B{0}:B{1},"{2}",L{0}:L{1},"{3}")'.format(R1, R2, typ, YES))
    c.font = font(bold=True, color=SLATE); c.number_format = MONEY; c.fill = fill(LIGHT)
widths(rc, {'A': 26, 'B': 10, 'C': 22, 'D': 12, 'E': 12, 'F': 15, 'G': 16, 'H': 15, 'I': 9, 'J': 13, 'K': 13, 'L': 8, 'M': 13, 'N': 13, 'O': 14})
rc.freeze_panes = 'B5'
rc.sheet_properties.tabColor = TEAL
if DEMO:
    def ld(y, m, d): return dt.date(y, m, d)
    REC = [  # name, type, category, amount, for, payer, method, freq idx, first, last posted
        ('Salary', INC, 'Salary – Adult 1', 2900, P(3), '', M(0), 0, ld(2026, 1, 28), ld(2026, 10, 28)),
        ('Salary', INC, 'Salary – Adult 2', 2150, P(4), '', M(0), 0, ld(2026, 1, 27), ld(2026, 10, 27)),
        ('Flat rental income', INC, 'Rental income', 750, P(0), '', M(0), 0, ld(2026, 1, 5), ld(2026, 10, 5)),
        ('Child benefit', INC, 'Child benefit', 250, P(0), '', M(0), 0, ld(2026, 1, 15), ld(2026, 10, 15)),
        ('Rent', EXP, 'Housing', 1200, P(0), PB(2), M(1), 0, ld(2026, 1, 1), ld(2026, 10, 1)),
        ('Internet', EXP, 'Utilities', 45, P(0), PB(2), M(1), 0, ld(2026, 1, 3), ld(2026, 10, 3)),
        ('Mobile phones', EXP, 'Utilities', 30, P(0), PB(0), M(1), 0, ld(2026, 1, 3), ld(2026, 10, 3)),
        ('Water', EXP, 'Utilities', 35, P(0), PB(2), M(1), 1, ld(2026, 2, 20), ld(2026, 10, 20)),
        ('Nursery fee', EXP, 'Kids – Childcare', 450, P(2), PB(2), M(1), 0, ld(2026, 1, 2), ld(2026, 10, 2)),
        ('School canteen', EXP, 'Kids – School', 85.5, P(1), PB(2), M(1), 0, ld(2026, 1, 5), ld(2026, 10, 5)),
        ('Swimming club', EXP, 'Kids – Activities', 40, P(1), PB(0), M(1), 0, ld(2026, 1, 6), ld(2026, 10, 6)),
        ('Health insurance', EXP, 'Insurance', 62, P(0), PB(2), M(1), 0, ld(2026, 1, 8), ld(2026, 10, 8)),
        ('Streaming', EXP, 'Subscriptions', 15.99, P(0), PB(0), M(3), 0, ld(2026, 1, 12), ld(2026, 10, 12)),
        ('Car insurance', EXP, 'Insurance', 480, P(0), PB(2), M(1), 4, ld(2026, 3, 10), ld(2026, 3, 10)),
        ('Home insurance', EXP, 'Insurance', 300, P(0), PB(2), M(1), 4, ld(2026, 9, 15), ld(2026, 9, 15)),
    ]
    for i, (n, typ, c, a, who, pb, meth, fi, first, lp) in enumerate(REC):
        r = R1 + i
        for col, v in ((1, D(n)), (2, typ), (3, cat(c)), (4, a), (5, who), (6, pb or None), (7, meth), (8, F(fi)), (10, first), (12, YES), (15, lp)):
            rc.cell(r, col, v)

# ================================================================ Bills & Taxes
bt = WS[S_BILLS]
banner(bt, L('Bills & Taxes', 'Facturas e impuestos'), L('One-off or yearly payments with a due date. The bot reminds you before the date.',
       'Pagos puntuales o anuales con fecha límite. El bot te avisa antes de la fecha.'), 8)
header(bt, 4, 1, [L('Item', 'Concepto'), L('Category', 'Categoría'), L('Due date', 'Vencimiento'), L('Expected amount', 'Importe previsto'),
                  L('Status', 'Estado'), L('Days left', 'Días restantes'), L('Notes', 'Notas'), L('Reminded', 'Recordado')], height=30)
B1, B2 = 5, 44
DAYS = L('0" days";[Red]"overdue "0;"today"', '0" días";[Red]"vencida "0;"hoy"')
for r in range(B1, B2 + 1):
    for c in range(1, 9):
        x = bt.cell(r, c); x.border = BOX; x.font = font()
        if c in (1, 2, 3, 4, 5, 7): x.fill = fill(INPUT)
    bt.cell(r, 3).number_format = DATEF; bt.cell(r, 8).number_format = DATEF; bt.cell(r, 4).number_format = MONEY
    bt.cell(r, 6, '=IF(OR(C{0}="",E{0}<>"{1}"),"",C{0}-TODAY())'.format(r, PEND)).number_format = DAYS
    bt.cell(r, 8).font = font(color=GREY)
dv(bt, setref('A', 5, 56), 'B%d:B%d' % (B1, B2))
dv(bt, setref('U', 5, 7), 'E%d:E%d' % (B1, B2))
dv_date(bt, 'C%d:C%d' % (B1, B2))
bt.conditional_formatting.add('A%d:H%d' % (B1, B2), FormulaRule(formula=['AND($E{0}="{1}",$C{0}<>"",$C{0}-TODAY()<=7)'.format(B1, PEND)], fill=fill('FBE3DC')))
bt.conditional_formatting.add('A%d:H%d' % (B1, B2), FormulaRule(formula=['$E{0}="{1}"'.format(B1, PAID)], font=Font(color=GREY)))
bt.cell(46, 1, L('Still to pay (pending)', 'Pendiente de pagar')).font = font(bold=True, color=SLATE)
c = bt.cell(46, 4, '=SUMIFS(D{0}:D{1},E{0}:E{1},"{2}")'.format(B1, B2, PEND)); c.number_format = MONEY; c.font = font(bold=True, color=SLATE); c.fill = fill(LIGHT)
widths(bt, {'A': 30, 'B': 22, 'C': 14, 'D': 14, 'E': 12, 'F': 13, 'G': 30, 'H': 13})
bt.freeze_panes = 'A5'
bt.sheet_properties.tabColor = CORAL
if DEMO:
    BILLS = [
        (L('Vehicle tax 2026', 'Impuesto de circulación 2026'), 'Taxes', dt.date(2026, 4, 30), 112.5, PAID),
        (L('Income tax return 2025', 'Declaración de la renta 2025'), 'Taxes', dt.date(2026, 6, 25), 380, PAID),
        (L('Property tax 2026', 'Impuesto de bienes inmuebles 2026'), 'Taxes', dt.date(2026, 6, 30), 420, PAID),
        (L('Waste tax 2026', 'Tasa de basuras 2026'), 'Taxes', dt.date(2026, 10, 31), 96, PAID),
        (L('School trip – Child 2', 'Excursión del colegio – Hijo/a 2'), 'Kids – School', dt.date(2026, 11, 25), 60, PEND),
        (L('Car inspection', 'Inspección del coche'), 'Transport', dt.date(2026, 12, 10), 45, PEND),
        (L('Vehicle tax 2027', 'Impuesto de circulación 2027'), 'Taxes', dt.date(2027, 4, 30), 115, PEND),
        (L('Property tax 2027', 'Impuesto de bienes inmuebles 2027'), 'Taxes', dt.date(2027, 6, 30), 430, PEND),
    ]
    for i, (it, c, d, a, s) in enumerate(BILLS):
        r = B1 + i
        bt.cell(r, 1, it); bt.cell(r, 2, cat(c)); bt.cell(r, 3, d); bt.cell(r, 4, a); bt.cell(r, 5, s)

# ================================================================ Dashboard
ds = WS[S_DASH]
view(ds, NAVY)
banner(ds, L('Family Budget – Dashboard', 'Presupuesto familiar – Panel'),
       L('Choose a year in the yellow cell. Everything here is calculated from Transactions.',
         'Elige un año en la celda amarilla. Todo se calcula a partir de Movimientos.'), 11)
ds['A4'] = L('Year', 'Año'); ds['A4'].font = font(bold=True)
ds['B4'] = '=YEAR(TODAY())' if not DEMO else 2026
inp(ds['B4'], '0')
dv(ds, setref('W', 5, 21), 'B4', warn=False)
y0, y1 = 'DATE($B$4,1,1)', 'DATE($B$4+1,1,1)'
tile(ds, 6, 1, L('Income', 'Ingresos'), '=SUM(B14:B25)', color=TEAL)
tile(ds, 6, 3, L('Expenses', 'Gastos'), '=SUM(C14:C25)', color=CORAL)
tile(ds, 6, 5, L('Net (income – expenses)', 'Neto (ingresos – gastos)'), '=A7-C7', color=SLATE)
tile(ds, 6, 8, L('Savings rate', 'Tasa de ahorro'), '=IF(A7=0,"",E7/A7)', fmt=PCT, color=NAVY)
tile(ds, 9, 1, L('Fixed monthly income', 'Ingresos fijos al mes'), "=" + q(S_REC) + "!N55", color=TEAL)
tile(ds, 9, 3, L('Fixed monthly costs', 'Gastos fijos al mes'), "=" + q(S_REC) + "!N56", color=CORAL)
tile(ds, 9, 5, L('Bills still to pay', 'Facturas pendientes'), "=" + q(S_BILLS) + "!D46", color=SLATE)
tile(ds, 9, 8, L('Entries this year', 'Movimientos este año'), '=COUNTIFS(' + TXA + ',">="&' + y0 + ',' + TXA + ',"<"&' + y1 + ')', fmt='0', color=NAVY)
section(ds, 12, 1, L('Month by month', 'Mes a mes'))
header(ds, 13, 1, [L('Month', 'Mes'), L('Income', 'Ingresos'), L('Expenses', 'Gastos'), L('Net', 'Neto'), L('Savings rate', 'Tasa de ahorro'),
                   L('Cumulative net', 'Neto acumulado')], height=30)
for i in range(12):
    r = 14 + i
    s, e = 'DATE($B$4,%d,1)' % (i + 1), 'DATE($B$4,%d,1)' % (i + 2)
    ds.cell(r, 1, MONTHS[i]).font = font(bold=True)
    ds.cell(r, 2, '=' + sum_period(s, e, ',' + TXB + ',"' + INC + '"'))
    ds.cell(r, 3, '=' + sum_period(s, e, ',' + TXB + ',"' + EXP + '"'))
    ds.cell(r, 4, '=B{0}-C{0}'.format(r))
    ds.cell(r, 5, '=IF(B{0}=0,"",D{0}/B{0})'.format(r)).number_format = PCT
    ds.cell(r, 6, '=SUM($D$14:D{0})'.format(r))
    for c in (2, 3, 4, 6): ds.cell(r, c).number_format = MONEY
    for c in range(1, 7):
        ds.cell(r, c).border = BOX
        if i % 2: ds.cell(r, c).fill = fill(PAPER)
ds.cell(26, 1, 'Total').font = font(bold=True)
for c, f in ((2, '=SUM(B14:B25)'), (3, '=SUM(C14:C25)'), (4, '=B26-C26'), (5, '=IF(B26=0,"",D26/B26)')):
    x = ds.cell(26, c, f); x.font = font(bold=True); x.number_format = PCT if c == 5 else MONEY
for c in range(1, 7): ds.cell(26, c).fill = fill(LIGHT); ds.cell(26, c).border = BOX
ds.cell(28, 1, L('Check – entries this year whose category is not in Settings (should be 0):',
                 'Control – movimientos de este año con una categoría que no está en Configuración (debería ser 0):')).font = font(size=9, italic=True, color=GREY)
chk = ds.cell(28, 6, '=SUMPRODUCT((' + TXA + '>=' + y0 + ')*(' + TXA + '<' + y1 + ')*(COUNTIF(' + setref('A', 5, 56) + ',' + TXC + ')=0))')
chk.font = font(bold=True); chk.number_format = '0'
ds.conditional_formatting.add('F28', FormulaRule(formula=['$F$28>0'], fill=fill('FBE3DC'), font=Font(color='C0392B', bold=True)))
# helpers Z..AD rows 5..56 (one per Settings category row)
for i in range(52):
    r, sr = 5 + i, 5 + i
    ds.cell(r, 26, '=IF({0}!A{1}="","",{0}!A{1})'.format(q(S_SET), sr))
    ds.cell(r, 27, '={0}!B{1}'.format(q(S_SET), sr))
    ds.cell(r, 28, '=IF(Z{0}="","",'.format(r) + sum_period(y0, y1, ',' + TXC + ',Z%d,' % r + TXB + ',AA%d' % r) + ')')
    ds.cell(r, 29, '=IF(AND(AA{0}="{1}",N(AB{0})>0),AB{0}+(60-ROW())/10000000,"")'.format(r, EXP))
    ds.cell(r, 30, '=IF(AND(AA{0}="{1}",N(AB{0})>0),AB{0}+(60-ROW())/10000000,"")'.format(r, INC))
ds['Z4'], ds['AA4'], ds['AB4'], ds['AC4'], ds['AD4'] = 'cat', 'type', 'year', 'exp key', 'inc key'
hide(ds, 26, 30)
def toplist(row0, n, keycol, typ, title, total_ref):
    section(ds, row0 - 1, 8, title)
    header(ds, row0, 8, ['#', L('Category', 'Categoría'), L('Amount', 'Importe'), L('Share', 'Peso')])
    for k in range(1, n + 1):
        r = row0 + k
        ds.cell(r, 8, k).alignment = CENTER
        ds.cell(r, 9, '=IFERROR(INDEX($Z$5:$Z$56,MATCH(LARGE(${0}$5:${0}$56,{1}),${0}$5:${0}$56,0)),"")'.format(keycol, k))
        ds.cell(r, 10, '=IF(I{0}="","",SUMIFS($AB$5:$AB$56,$Z$5:$Z$56,I{0},$AA$5:$AA$56,"{1}"))'.format(r, typ)).number_format = MONEY
        ds.cell(r, 11, '=IF(OR(I{0}="",{1}=0),"",J{0}/{1})'.format(r, total_ref)).number_format = PCT
        for c in range(8, 12):
            ds.cell(r, c).border = BOX
            if k % 2 == 0: ds.cell(r, c).fill = fill(PAPER)
toplist(13, 10, 'AC', EXP, L('Top 10 expenses (year)', 'Top 10 gastos (año)'), '$C$26')
toplist(26, 6, 'AD', INC, L('Top income (year)', 'Principales ingresos (año)'), '$B$26')
widths(ds, {'A': 15, 'B': 14, 'C': 14, 'D': 14, 'E': 13, 'F': 15, 'G': 3, 'H': 8, 'I': 26, 'J': 14, 'K': 10})
ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'
ch.title = L('Income vs expenses per month', 'Ingresos y gastos por mes')
ch.add_data(Reference(ds, min_col=2, max_col=3, min_row=13, max_row=25), titles_from_data=True)
ch.set_categories(Reference(ds, min_col=1, min_row=14, max_row=25))
ch.height, ch.width = 8, 17
try:
    ch.series[0].graphicalProperties.solidFill = TEAL; ch.series[1].graphicalProperties.solidFill = CORAL
except Exception: pass
ds.add_chart(ch, 'A35')
ch2 = BarChart(); ch2.type = 'bar'; ch2.title = L('Top 10 expenses', 'Top 10 gastos')
ch2.add_data(Reference(ds, min_col=10, min_row=13, max_row=23), titles_from_data=True)
ch2.set_categories(Reference(ds, min_col=9, min_row=14, max_row=23))
ch2.legend = None; ch2.height, ch2.width = 8, 13
ch2.y_axis.scaling.orientation = 'minMax'; ch2.x_axis.scaling.orientation = 'maxMin'
try: ch2.series[0].graphicalProperties.solidFill = CORAL
except Exception: pass
ds.add_chart(ch2, 'H35')
ch3 = LineChart(); ch3.title = L('Cumulative net over the year', 'Neto acumulado durante el año')
ch3.add_data(Reference(ds, min_col=6, min_row=13, max_row=25), titles_from_data=True)
ch3.set_categories(Reference(ds, min_col=1, min_row=14, max_row=25))
ch3.legend = None; ch3.height, ch3.width = 8, 17
ds.add_chart(ch3, 'A52')

# ================================================================ Analysis
an = WS[S_ANA]
view(an, NAVY)
banner(an, L('Analysis – where the money goes', 'Análisis – en qué se va el dinero'),
       L('Pick a year and a month (or "All year"). Categories are ranked from highest to lowest.',
         'Elige un año y un mes (o "Todo el año"). Las categorías se ordenan de mayor a menor.'), 11)
an['A4'] = L('Year', 'Año'); an['A5'] = L('Period', 'Periodo')
for c in ('A4', 'A5'): an[c].font = font(bold=True)
an['B4'] = '=YEAR(TODAY())' if not DEMO else 2026; inp(an['B4'], '0')
an['B5'] = ALLYEAR; inp(an['B5'])
dv(an, setref('W', 5, 21), 'B4', warn=False)
dv(an, setref('X', 5, 17), 'B5', warn=False)
# period helpers row 2, columns AA..AH
an['AA1'], an['AB1'], an['AC1'], an['AD1'], an['AE1'], an['AF1'], an['AG1'], an['AH1'] = 'month', 'start', 'end', 'prev start', 'prev end', 'ly start', 'ly end', 'months'
an['AA2'] = '=IF($B$5="{0}",0,IFERROR(MATCH($B$5,{1},0),0))'.format(ALLYEAR, setref('R', 5, 16))
an['AB2'] = '=IF(AA2=0,DATE($B$4,1,1),DATE($B$4,AA2,1))'
an['AC2'] = '=IF(AA2=0,DATE($B$4+1,1,1),DATE($B$4,AA2+1,1))'
an['AD2'] = '=IF(AA2=0,DATE($B$4-1,1,1),DATE($B$4,AA2-1,1))'
an['AE2'] = '=AB2'
an['AF2'] = '=DATE($B$4-1,IF(AA2=0,1,AA2),1)'
an['AG2'] = '=IF(AA2=0,DATE($B$4,1,1),DATE($B$4-1,AA2+1,1))'
an['AH2'] = '=IF(AA2=0,12,1)'
for i in range(52):
    r, sr = 9 + i, 5 + i
    tsum = lambda s, e: sum_period(s, e, ',' + TXC + ',$Z%d,' % r + TXB + ',$AA%d' % r)
    an.cell(r, 26, '=IF({0}!A{1}="","",{0}!A{1})'.format(q(S_SET), sr))
    an.cell(r, 27, '={0}!B{1}'.format(q(S_SET), sr))
    an.cell(r, 28, '=IF($Z{0}="","",{1})'.format(r, tsum('$AB$2', '$AC$2')))
    an.cell(r, 29, '=IF($Z{0}="","",{1})'.format(r, tsum('$AD$2', '$AE$2')))
    an.cell(r, 30, '=IF($Z{0}="","",{1})'.format(r, tsum('$AF$2', '$AG$2')))
    an.cell(r, 31, '=IF($Z{0}="","",N({1}!C{2})*$AH$2)'.format(r, q(S_SET), sr))
    an.cell(r, 32, '=IF(AND(AA{0}="{1}",N(AB{0})>0),AB{0}+(80-ROW())/10000000,"")'.format(r, EXP))
    an.cell(r, 33, '=IF(AND(AA{0}="{1}",N(AB{0})>0),AB{0}+(80-ROW())/10000000,"")'.format(r, INC))
an['Z8'], an['AA8'], an['AB8'], an['AC8'], an['AD8'], an['AE8'], an['AF8'], an['AG8'] = 'cat', 'type', 'period', 'prev', 'last year', 'budget', 'exp key', 'inc key'
hide(an, 26, 40)
pinc = '=' + sum_period('$AB$2', '$AC$2', ',' + TXB + ',"' + INC + '"')
pexp = '=' + sum_period('$AB$2', '$AC$2', ',' + TXB + ',"' + EXP + '"')
tile(an, 7, 1, L('Income in the period', 'Ingresos del periodo'), pinc, color=TEAL)
tile(an, 7, 3, L('Expenses in the period', 'Gastos del periodo'), pexp, color=CORAL)
tile(an, 7, 5, L('Net', 'Neto'), '=A8-C8', color=SLATE)
tile(an, 7, 7, L('Expenses – previous period', 'Gastos – periodo anterior'), '=' + sum_period('$AD$2', '$AE$2', ',' + TXB + ',"' + EXP + '"'), color=GREY)
tile(an, 7, 9, L('Expenses – same period last year', 'Gastos – mismo periodo año anterior'), '=' + sum_period('$AF$2', '$AG$2', ',' + TXB + ',"' + EXP + '"'), color=GREY)
AH = ['#', L('Category', 'Categoría'), L('Amount', 'Importe'), L('Share', 'Peso'), '', L('Budget', 'Presupuesto'),
      L('% of budget used', '% presup. usado'), L('Previous period', 'Periodo anterior'), L('Same period last year', 'Mismo periodo año anterior')]
def ranked(r0, n, keycol, typ, title, total):
    section(an, r0 - 1, 1, title)
    header(an, r0, 1, AH, height=44)
    for k in range(1, n + 1):
        r = r0 + k
        an.cell(r, 1, k).alignment = CENTER
        an.cell(r, 2, '=IFERROR(INDEX($Z$9:$Z$60,MATCH(LARGE(${0}$9:${0}$60,{1}),${0}$9:${0}$60,0)),"")'.format(keycol, k))
        for c, hc in ((3, 'AB'), (6, 'AE'), (8, 'AC'), (9, 'AD')):
            an.cell(r, c, '=IF($B{0}="","",SUMIFS(${1}$9:${1}$60,$Z$9:$Z$60,$B{0},$AA$9:$AA$60,"{2}"))'.format(r, hc, typ)).number_format = MONEY
        an.cell(r, 4, '=IF(OR($B{0}="",{1}=0),"",C{0}/{1})'.format(r, total)).number_format = PCT
        x = an.cell(r, 5, '=IF(D{0}="","",REPT("█",ROUND(D{0}*25,0)))'.format(r)); x.font = font(color=TEAL if typ == INC else CORAL)
        an.cell(r, 7, '=IF(OR($B{0}="",N(F{0})=0),"",C{0}/F{0})'.format(r)).number_format = PCT
        for c in range(1, 10):
            an.cell(r, c).border = BOX
            if k % 2 == 0: an.cell(r, c).fill = fill(PAPER)
    rng = 'G%d:G%d' % (r0 + 1, r0 + n)
    if typ == EXP:
        an.conditional_formatting.add(rng, FormulaRule(formula=['AND(ISNUMBER(G%d),G%d>1)' % (r0 + 1, r0 + 1)], fill=fill('FBE3DC'), font=Font(color='C0392B', bold=True)))
ranked(12, 20, 'AF', EXP, L('Expenses, highest first', 'Gastos, de mayor a menor'), '$C$8')
ranked(36, 12, 'AG', INC, L('Income, highest first', 'Ingresos, de mayor a menor'), '$A$8')
widths(an, {'A': 6, 'B': 26, 'C': 14, 'D': 10, 'E': 26, 'F': 14, 'G': 12, 'H': 14, 'I': 15, 'J': 3})
ch4 = BarChart(); ch4.type = 'bar'; ch4.title = L('Expenses in the period', 'Gastos del periodo')
ch4.add_data(Reference(an, min_col=3, min_row=12, max_row=22), titles_from_data=True)
ch4.set_categories(Reference(an, min_col=2, min_row=13, max_row=22))
ch4.legend = None; ch4.height, ch4.width = 9, 14; ch4.x_axis.scaling.orientation = 'maxMin'
try: ch4.series[0].graphicalProperties.solidFill = CORAL
except Exception: pass
an.add_chart(ch4, 'K12')
ch5 = BarChart(); ch5.type = 'bar'; ch5.title = L('Income in the period', 'Ingresos del periodo')
ch5.add_data(Reference(an, min_col=3, min_row=36, max_row=42), titles_from_data=True)
ch5.set_categories(Reference(an, min_col=2, min_row=37, max_row=42))
ch5.legend = None; ch5.height, ch5.width = 7, 14; ch5.x_axis.scaling.orientation = 'maxMin'
try: ch5.series[0].graphicalProperties.solidFill = TEAL
except Exception: pass
an.add_chart(ch5, 'K36')

# ================================================================ Year
yr = WS[S_YEAR]
view(yr, NAVY)
banner(yr, L('Year at a glance', 'El año de un vistazo'), L('Every category, month by month, against its budget.',
       'Cada categoría, mes a mes, frente a su presupuesto.'), 16)
yr['A4'] = L('Year', 'Año'); yr['A4'].font = font(bold=True)
yr['B4'] = '=YEAR(TODAY())' if not DEMO else 2026; inp(yr['B4'], '0')
dv(yr, setref('W', 5, 21), 'B4', warn=False)
MSHORT = [m[:3] for m in MONTHS]
def grid(r0, title, rows, typ, label_fn, crit_col, color, budget=True):
    section(yr, r0 - 1, 1, title)
    header(yr, r0, 1, [L('Category', 'Categoría')] + MSHORT + ['Total'] + ([L('Budget (year)', 'Presupuesto (año)'), L('Difference', 'Diferencia')] if budget else []), color=color, height=30)
    for i, sr in enumerate(rows):
        r = r0 + 1 + i
        yr.cell(r, 1, label_fn(sr)).font = font(bold=True)
        for m in range(12):
            s, e = 'DATE($B$4,%d,1)' % (m + 1), 'DATE($B$4,%d,1)' % (m + 2)
            extra = ',' + crit_col + ',$A%d' % r + (',' + TXB + ',"' + typ + '"' if typ else '')
            yr.cell(r, 2 + m, '=IF($A{0}="","",{1})'.format(r, sum_period(s, e, extra))).number_format = MONEY
        yr.cell(r, 14, '=IF($A{0}="","",SUM(B{0}:M{0}))'.format(r)).number_format = MONEY
        yr.cell(r, 14).font = font(bold=True)
        for c in range(1, 17 if budget else 15):
            yr.cell(r, c).border = BOX
            if i % 2: yr.cell(r, c).fill = fill(PAPER)
    return r0 + len(rows)
def setlab(sr): return '=IF({0}!A{1}="","",{0}!A{1})'.format(q(S_SET), sr)
last_inc = grid(7, L('Income', 'Ingresos'), list(range(5, 17)), INC, setlab, TXC, TEAL)
for i in range(12):
    r, sr = 8 + i, 5 + i
    yr.cell(r, 15, '=IF($A{0}="","",N({1}!C{2})*12)'.format(r, q(S_SET), sr)).number_format = MONEY
    yr.cell(r, 16, '=IF($A{0}="","",N(N{0})-N(O{0}))'.format(r)).number_format = MONEY
TI = last_inc + 1
yr.cell(TI, 1, L('Total income', 'Total ingresos')).font = font(bold=True)
for c in range(2, 17):
    x = yr.cell(TI, c, '=SUM({0}8:{0}{1})'.format(CL(c), last_inc)); x.number_format = MONEY; x.font = font(bold=True)
for c in range(1, 17): yr.cell(TI, c).fill = fill(GREEN)
E0 = TI + 3
last_exp = grid(E0, L('Expenses', 'Gastos'), list(range(17, 57)), EXP, setlab, TXC, CORAL)
for i in range(40):
    r, sr = E0 + 1 + i, 17 + i
    yr.cell(r, 15, '=IF($A{0}="","",N({1}!C{2})*12)'.format(r, q(S_SET), sr)).number_format = MONEY
    yr.cell(r, 16, '=IF($A{0}="","",N(O{0})-N(N{0}))'.format(r)).number_format = MONEY
TE = last_exp + 1
yr.cell(TE, 1, L('Total expenses', 'Total gastos')).font = font(bold=True)
for c in range(2, 17):
    x = yr.cell(TE, c, '=SUM({0}{1}:{0}{2})'.format(CL(c), E0 + 1, last_exp)); x.number_format = MONEY; x.font = font(bold=True)
for c in range(1, 17): yr.cell(TE, c).fill = fill('FBE3DC')
yr.conditional_formatting.add('P{0}:P{1}'.format(E0 + 1, last_exp), FormulaRule(formula=['AND(ISNUMBER(P{0}),P{0}<0,N(O{0})>0)'.format(E0 + 1)], font=Font(color='C0392B', bold=True)))
RN = TE + 2
section(yr, RN, 1, L('Result', 'Resultado'))
yr.cell(RN + 1, 1, L('Net (income – expenses)', 'Neto (ingresos – gastos)')).font = font(bold=True)
yr.cell(RN + 2, 1, L('Savings rate', 'Tasa de ahorro')).font = font(bold=True)
for c in range(2, 15):
    L_ = CL(c)
    yr.cell(RN + 1, c, '={0}{1}-{0}{2}'.format(L_, TI, TE)).number_format = MONEY
    yr.cell(RN + 2, c, '=IF(N({0}{1})=0,"",{0}{2}/{0}{1})'.format(L_, TI, RN + 1)).number_format = PCT
    for rr in (RN + 1, RN + 2): yr.cell(rr, c).border = BOX; yr.cell(rr, c).fill = fill(LIGHT)
P0 = RN + 5
def plab(col):
    return lambda sr: '=IF({0}!{1}{2}="","",{0}!{1}{2})'.format(q(S_SET), col, sr)
lastp = grid(P0, L('Expenses by person (column "For")', 'Gastos por persona (columna "Para")'), list(range(5, 17)), EXP, plab('F'), TXF, SLATE, budget=False)
Q0 = lastp + 3
lastq = grid(Q0, L('Expenses by who paid (column "Paid by")', 'Gastos por quién pagó (columna "Pagado por")'), list(range(5, 13)), EXP, plab('G'), TXG, SLATE, budget=False)
widths(yr, dict([('A', 28)] + [(CL(c), 10.5) for c in range(2, 14)] + [('N', 12), ('O', 13), ('P', 12)]))
yr.freeze_panes = 'B5'

# ================================================================ Search
se = WS[S_SEARCH]
view(se, NAVY)
banner(se, L('Search', 'Buscar'), L('Fill in any of the yellow filters (leave empty = all). Shows the newest 100 matches.',
       'Rellena los filtros amarillos que quieras (vacío = todos). Muestra las 100 coincidencias más recientes.'), 9)
flt = [(4, L('Word (description, notes, category, person)', 'Palabra (descripción, notas, categoría, persona)')), (5, L('Type', 'Tipo')),
       (6, L('Category', 'Categoría')), (7, L('Year', 'Año')), (8, L('Month', 'Mes')), (9, L('For (person)', 'Para (persona)'))]
for r, lab in flt:
    se.cell(r, 1, lab).font = font(bold=True); inp(se.cell(r, 2)); se.cell(r, 2).alignment = Alignment(horizontal='left', vertical='center')
se['B7'].number_format = '0'
if DEMO: se['B4'] = L('school', 'colegio')
dv(se, setref('S', 5, 6), 'B5', warn=False)
dv(se, setref('A', 5, 56), 'B6', warn=False)
dv(se, setref('W', 5, 21), 'B7', warn=False)
dv(se, setref('R', 5, 16), 'B8', warn=False)
dv(se, setref('F', 5, 16), 'B9', warn=False)
se['A10'] = L('Found', 'Encontrados'); se['B10'] = '=COUNT($AA$2:$AA$%d)' % NTX
se['C10'] = L('Expenses', 'Gastos'); se['D10'] = '=SUMIFS(%s,$AA$2:$AA$%d,">0",%s,"%s")' % (TX('E'), NTX, TXB, EXP)
se['E10'] = L('Income', 'Ingresos'); se['F10'] = '=SUMIFS(%s,$AA$2:$AA$%d,">0",%s,"%s")' % (TX('E'), NTX, TXB, INC)
for c in ('A10', 'C10', 'E10'): se[c].font = font(bold=True, color=SLATE)
for c in ('B10', 'D10', 'F10'): se[c].font = font(bold=True); se[c].fill = fill(LIGHT); se[c].border = BOX
se['D10'].number_format = MONEY; se['F10'].number_format = MONEY
SH = [L('Date', 'Fecha'), L('Type', 'Tipo'), L('Category', 'Categoría'), L('Description', 'Descripción'), L('Amount', 'Importe'),
      L('For', 'Para'), L('Paid by', 'Pagado por'), L('Notes', 'Notas'), L('Row in Transactions', 'Fila en Movimientos')]
header(se, 12, 1, SH, height=30)
T = q(S_TX)
for i in range(2, NTX + 1):
    tr = i
    cond = ('OR($B$4="",ISNUMBER(SEARCH($B$4,{T}!$D{r}&" "&{T}!$C{r}&" "&{T}!$J{r}&" "&{T}!$F{r})))'.format(T=T, r=tr),
            'OR($B$5="",{T}!$B{r}=$B$5)'.format(T=T, r=tr),
            'OR($B$6="",{T}!$C{r}=$B$6)'.format(T=T, r=tr),
            'IF($B$7="",TRUE,YEAR({T}!$A{r})=$B$7)'.format(T=T, r=tr),
            'IF($B$8="",TRUE,MONTH({T}!$A{r})=MATCH($B$8,{S},0))'.format(T=T, r=tr, S=setref('R', 5, 16)),
            'OR($B$9="",{T}!$F{r}=$B$9)'.format(T=T, r=tr))
    se.cell(i, 27, '=IF({T}!$A{r}="","",IFERROR(IF(AND({c}),INT({T}!$A{r})+{r}/100000,""),""))'.format(T=T, r=tr, c=','.join(cond)))
COLS = [('A', None), ('B', 1), ('C', 1), ('D', 1), ('E', None), ('F', 1), ('G', 1), ('J', 1)]
for k in range(1, 101):
    r = 12 + k
    se.cell(r, 28, '=IFERROR(LARGE($AA$2:$AA${0},{1}),"")'.format(NTX, k))
    se.cell(r, 29, '=IF(AB{0}="","",MATCH(AB{0},$AA$2:$AA${1},0)+1)'.format(r, NTX))
    for j, (col, txt) in enumerate(COLS):
        f = 'INDEX({T}!${c}:${c},$AC{r})'.format(T=T, c=col, r=r)
        se.cell(r, 1 + j, '=IF($AC{0}="","",{1}{2})'.format(r, f, '&""' if txt else ''))
    se.cell(r, 9, '=IF($AC{0}="","",$AC{0})'.format(r))
    se.cell(r, 1).number_format = DATEF; se.cell(r, 5).number_format = MONEY
    for c in range(1, 10):
        se.cell(r, c).border = BOX
        if k % 2 == 0: se.cell(r, c).fill = fill(PAPER)
se['AA1'], se['AB12'], se['AC12'] = 'key', 'k-th', 'row'
hide(se, 26, 29)
widths(se, {'A': 40, 'B': 22, 'C': 22, 'D': 34, 'E': 13, 'F': 13, 'G': 15, 'H': 30, 'I': 12})
se.freeze_panes = 'A13'

# ================================================================ How to use
hw = WS[S_HOW]
view(hw, TEAL)
banner(hw, L('How to use this file', 'Cómo usar este archivo'), L('Read this once. It takes five minutes.', 'Léelo una vez. Son cinco minutos.'), 2)
r = 4
for kind, text in HOWTO['es' if ES else 'en']:
    if kind == 'h':
        r += 1
        c = hw.cell(r, 1, text)
        c.font = font(size=12, bold=True, color='FFFFFF'); c.fill = fill(SLATE)
        hw.row_dimensions[r].height = 22
    elif kind == 'q':
        c = hw.cell(r, 1, text); c.font = font(bold=True, color=NAVY)
    else:
        c = hw.cell(r, 1, text); c.font = font(); c.alignment = Alignment(wrap_text=True, vertical='top')
    r += 1
hw.column_dimensions['A'].width = 120
hw.column_dimensions['B'].width = 2

# ---------------------------------------------------------------- finish
wb.properties.creator = ''
wb.properties.lastModifiedBy = ''
wb.properties.title = L('Family Budget', 'Presupuesto familiar')
wb.active = 0
wb.save(OUT)
print('saved', OUT, 'rows', len(data))
