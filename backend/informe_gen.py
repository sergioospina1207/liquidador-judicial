"""
Generador del Informe Técnico de Liquidación de Prestaciones Sociales
Formato neutro institucional — Times New Roman, sin colores de app
Para entrega a fondos de prestaciones sociales.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
import io

# ── Paleta neutral ──────────────────────────────────────────
BL   = 'FFFFFF'
NG   = '000000'
GR_L = 'F2F2F2'   # gris muy claro para cabeceras
GR_M = 'D9D9D9'   # gris medio para totales y ficha
GR_D = '595959'   # gris oscuro para texto secundario

TNR  = 'Times New Roman'

def _f(bold=False, size=11, color=NG, italic=False):
    return Font(name=TNR, bold=bold, size=size, color=color, italic=italic)

def _a(h='left', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _s(style='thin', color='000000'):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _st(style='thin', color='000000'):
    """Solo borde superior e inferior (top/bottom)"""
    s = Side(style=style, color=color)
    n = Side(style=None)
    return Border(top=s, bottom=s, left=n, right=n)

THIN  = _s('thin',   '000000')
MED   = _s('medium', '000000')
HAIR  = _s('hair',   '595959')

def _fill(c):
    return PatternFill('solid', fgColor=c)

def sc(ws, r, c, v, bold=False, size=11, color=NG, bg=None,
       h='left', wrap=False, nf=None, italic=False, border=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font     = _f(bold=bold, size=size, color=color, italic=italic)
    cell.alignment = _a(h=h, wrap=wrap)
    if bg:     cell.fill = _fill(bg)
    if nf:     cell.number_format = nf
    if border: cell.border = border
    return cell

def fr(ws, r, c1, c2, bg=None, border=None):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=r, column=c)
        if bg:     cell.fill   = _fill(bg)
        if border: cell.border = border

def fmt_fecha(s):
    if not s: return '—'
    s = str(s)
    if len(s) == 10 and s[4] == '-':
        return s[8:] + '/' + s[5:7] + '/' + s[:4]
    return s

NUM = '#,##0'
DEC2 = '0.00'
DEC4 = '0.0000'

# Columnas del informe (8 columnas de datos)
INF_COLS = [
    'Concepto',
    'Mes Causación',
    'Valor Bruto ($)',
    'Valor Neto ($)',
    'IPC Inicial',
    'IPC Final',
    'Vr. Indexación ($)',
    'Valor Indexado ($)',
]
N_COLS = len(INF_COLS)   # 8
LAST_COL = get_column_letter(N_COLS)

# Claves de prestaciones (excluye cesantías e intereses)
PREST_KEYS = ['prodS1','prodS2','servicios','primaServ','primaVac','sueldoVac10','primaNov']
PREST_LABELS = {
    'prodS1':      'Bonif. Productividad S1',
    'prodS2':      'Bonif. Productividad S2',
    'servicios':   'Bonif. de Servicios',
    'primaServ':   'Prima de Servicios',
    'primaVac':    'Prima de Vacaciones',
    'sueldoVac10': 'Sueldo Vacacional (1/10)',
    'primaNov':    'Prima de Navidad',
}


def generar_informe(data: dict) -> bytes:
    """
    Genera el Excel del Informe Técnico de Liquidación.

    Espera en `data`:
      cliente       – dict con nombre, cc, entidad, cargo, ejecutoria, turno, ipc, periodo
      prestsByYear  – dict {año: [{key, label, mes, bruto, neto, ipcIni, ipcFin, idxVal, idxTot}]}
      bjPorPeriodo  – list [{desde, hasta, cargo, bon}]
      observaciones – dict con criterios calculados
      resumen       – dict con totales globales (nomPrest, capPrest, etc.)
    """
    cliente      = data.get('cliente', {})
    prests_año   = data.get('prestsByYear', {})
    bj_periodos  = data.get('bjPorPeriodo', [])
    obs          = data.get('observaciones', {})
    resumen      = data.get('resumen', {})
    ces_por_año  = data.get('cesPorAño', {})

    wb = Workbook()
    ws = wb.active
    ws.title = 'Informe Técnico'

    # Anchos de columna
    col_widths = [28, 14, 18, 18, 10, 10, 18, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Configuración de impresión (A4 vertical o carta)
    ws.page_setup.paperSize = 9      # A4
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75,
                                   header=0.3, footer=0.3)
    ws.print_options.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    cr = 1  # fila actual

    # ══════════════════════════════════════════════════════
    # ENCABEZADO INSTITUCIONAL
    # ══════════════════════════════════════════════════════
    ws.row_dimensions[cr].height = 8
    cr += 1

    ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
    sc(ws, cr, 1,
       'RAMA JUDICIAL DEL PODER PÚBLICO',
       bold=True, size=10, color=GR_D, h='center')
    ws.row_dimensions[cr].height = 14
    cr += 1

    ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
    sc(ws, cr, 1,
       'INFORME TÉCNICO DE LIQUIDACIÓN DE PRESTACIONES SOCIALES',
       bold=True, size=14, h='center', border=None)
    ws.row_dimensions[cr].height = 22
    cr += 1

    ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
    sc(ws, cr, 1,
       'Elaborado conforme al artículo 177 C.P.A.C.A. y normativa vigente en materia salarial judicial',
       size=9, color=GR_D, h='center', italic=True)
    ws.row_dimensions[cr].height = 13
    cr += 1

    # Línea separadora
    ws.row_dimensions[cr].height = 4
    for c in range(1, N_COLS + 1):
        ws.cell(row=cr, column=c).border = Border(
            bottom=Side(style='medium', color='000000'))
    cr += 1

    # ══════════════════════════════════════════════════════
    # I. DATOS DEL BENEFICIARIO
    # ══════════════════════════════════════════════════════
    cr += 1
    ws.row_dimensions[cr].height = 14
    ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
    sc(ws, cr, 1, 'I.  DATOS DEL BENEFICIARIO Y PERÍODO', bold=True, size=11)
    cr += 1

    # Ficha en cuadrícula
    ficha = [
        [('Nombre completo:', cliente.get('nombre', '—'), 7)],
        [('Documento de identidad:', cliente.get('cc', '—'), 3),
         ('Entidad:', cliente.get('entidad', '—'), 3)],
        [('Cargo(s):', cliente.get('cargo', '—'), 7)],
        [('Período total:', cliente.get('periodo', '—'), 3),
         ('Ejecutoria:', fmt_fecha(cliente.get('ejecutoria', '')), 3)],
        [('Turno de pago:', cliente.get('turno', '—'), 3),
         ('IPC ejecutoria:', cliente.get('ipc', '—'), 3)],
    ]

    for fila in ficha:
        ws.row_dimensions[cr].height = 13
        col_pos = 1
        for label, valor, span in fila:
            # Celda etiqueta (1 columna)
            sc(ws, cr, col_pos, label, bold=True, size=10, bg=GR_L, border=THIN)
            col_pos += 1
            # Celda valor (span-1 columnas)
            end_col = col_pos + span - 2
            if end_col > col_pos:
                ws.merge_cells(
                    start_row=cr, start_column=col_pos,
                    end_row=cr, end_column=end_col)
            sc(ws, cr, col_pos, valor, size=10, border=THIN)
            col_pos = end_col + 1
        cr += 1

    cr += 1  # espacio

    # ══════════════════════════════════════════════════════
    # II. PRESTACIONES POR VIGENCIA (año a año)
    # ══════════════════════════════════════════════════════
    ws.row_dimensions[cr].height = 14
    ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
    sc(ws, cr, 1, 'II.  LIQUIDACIÓN DE PRESTACIONES SOCIALES POR VIGENCIA', bold=True, size=11)
    cr += 1
    cr += 1

    años_ordenados = sorted(prests_año.keys(), key=lambda x: int(x))

    total_bruto_global = 0
    total_neto_global  = 0
    total_idx_global   = 0
    total_idxval_global= 0

    for año in años_ordenados:
        items = prests_año[año]
        if not items:
            continue

        # Encabezado del año
        ws.row_dimensions[cr].height = 14
        bj_año = items[0].get('bon', 0) if items else 0
        ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
        sc(ws, cr, 1,
           f'Vigencia {año}   —   BJ Mensual Vigente: ${bj_año:,.0f}'.replace(',', '.'),
           bold=True, size=11, italic=False,
           border=Border(bottom=Side(style='thin', color='000000')))
        cr += 1

        # Cabecera de columnas
        ws.row_dimensions[cr].height = 14
        for ci, h in enumerate(INF_COLS, 1):
            sc(ws, cr, ci, h,
               bold=True, size=9, bg=GR_L,
               h='center' if ci > 1 else 'left',
               border=THIN)
        cr += 1

        # Filas de prestaciones
        sum_bruto = sum_neto = sum_idxval = sum_idx = 0
        for item in items:
            ws.row_dimensions[cr].height = 13
            bruto   = item.get('bruto', 0)
            neto    = item.get('neto', 0)
            ipc_ini = item.get('ipcIni', 0)
            ipc_fin = item.get('ipcFin', 0)
            idx_val = item.get('idxVal', 0)  # valor de la indexación (diferencia)
            idx_tot = item.get('idxTot', 0)  # valor indexado final

            sc(ws, cr, 1, item.get('label', ''), size=10, border=THIN)
            sc(ws, cr, 2, item.get('mes', ''), size=10, h='center', border=THIN)
            sc(ws, cr, 3, bruto,   size=10, h='right', nf=NUM, border=THIN)
            sc(ws, cr, 4, neto,    size=10, h='right', nf=NUM, border=THIN)
            sc(ws, cr, 5, ipc_ini, size=10, h='center', nf=DEC2, border=THIN)
            sc(ws, cr, 6, ipc_fin, size=10, h='center', nf=DEC2, border=THIN)
            sc(ws, cr, 7, idx_val, size=10, h='right', nf=NUM, border=THIN)
            sc(ws, cr, 8, idx_tot, size=10, h='right', nf=NUM, border=THIN)

            sum_bruto  += bruto
            sum_neto   += neto
            sum_idxval += idx_val
            sum_idx    += idx_tot
            cr += 1

        # Fila TOTAL del año
        ws.row_dimensions[cr].height = 13
        lbl_tot = f'TOTAL VIGENCIA {año}'
        sc(ws, cr, 1, lbl_tot, bold=True, size=10, bg=GR_M, border=MED)
        sc(ws, cr, 2, '',      size=10, bg=GR_M, border=MED)
        sc(ws, cr, 3, sum_bruto,  bold=True, size=10, bg=GR_M, h='right', nf=NUM, border=MED)
        sc(ws, cr, 4, sum_neto,   bold=True, size=10, bg=GR_M, h='right', nf=NUM, border=MED)
        sc(ws, cr, 5, '',         size=10, bg=GR_M, border=MED)
        sc(ws, cr, 6, '',         size=10, bg=GR_M, border=MED)
        sc(ws, cr, 7, sum_idxval, bold=True, size=10, bg=GR_M, h='right', nf=NUM, border=MED)
        sc(ws, cr, 8, sum_idx,    bold=True, size=10, bg=GR_M, h='right', nf=NUM, border=MED)
        cr += 1

        total_bruto_global  += sum_bruto
        total_neto_global   += sum_neto
        total_idxval_global += sum_idxval
        total_idx_global    += sum_idx

        cr += 1  # espacio entre años

    # TOTAL GENERAL
    ws.row_dimensions[cr].height = 15
    sc(ws, cr, 1, 'TOTAL GENERAL — PRESTACIONES SOCIALES',
       bold=True, size=11, bg=GR_D, color=BL, border=MED)
    sc(ws, cr, 2, '', bg=GR_D, border=MED)
    sc(ws, cr, 3, total_bruto_global,  bold=True, size=11, color=BL,
       bg=GR_D, h='right', nf=NUM, border=MED)
    sc(ws, cr, 4, total_neto_global,   bold=True, size=11, color=BL,
       bg=GR_D, h='right', nf=NUM, border=MED)
    sc(ws, cr, 5, '', bg=GR_D, border=MED)
    sc(ws, cr, 6, '', bg=GR_D, border=MED)
    sc(ws, cr, 7, total_idxval_global, bold=True, size=11, color=BL,
       bg=GR_D, h='right', nf=NUM, border=MED)
    sc(ws, cr, 8, total_idx_global,    bold=True, size=11, color=BL,
       bg=GR_D, h='right', nf=NUM, border=MED)
    cr += 2

    # ══════════════════════════════════════════════════════
    # III. CESANTÍAS E INTERESES SOBRE CESANTÍAS
    # ══════════════════════════════════════════════════════
    if ces_por_año:
        ws.row_dimensions[cr].height = 14
        ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
        sc(ws, cr, 1, 'III.  CESANTÍAS E INTERESES SOBRE CESANTÍAS', bold=True, size=11)
        cr += 1
        cr += 1

        # Cabecera tabla cesantías
        ces_hdrs = ['Vigencia', 'Ces. Valor Nominal ($)', 'Ces. Valor Indexado ($)',
                    'Int. Ces. Nominal ($)', 'Int. Ces. Indexado ($)', 'Total Cesantías ($)']
        ws.row_dimensions[cr].height = 14
        for ci, h in enumerate(ces_hdrs, 1):
            sc(ws, cr, ci, h, bold=True, size=9, bg=GR_L,
               h='right' if ci > 1 else 'left', border=THIN)
        cr += 1

        sum_ces_neto = sum_ces_idx = sum_int_neto = sum_int_idx = 0
        for año_c in sorted(ces_por_año.keys(), key=lambda x: int(x)):
            d = ces_por_año[año_c]
            ces_n = d.get('cesNeto', 0)
            ces_i = d.get('cesIdx',  0)
            int_n = d.get('intNeto', 0)
            int_i = d.get('intIdx',  0)
            tot   = ces_i + int_i
            ws.row_dimensions[cr].height = 13
            sc(ws, cr, 1, str(año_c),  size=10, border=THIN)
            sc(ws, cr, 2, ces_n, size=10, h='right', nf=NUM, border=THIN)
            sc(ws, cr, 3, ces_i, size=10, h='right', nf=NUM, border=THIN)
            sc(ws, cr, 4, int_n, size=10, h='right', nf=NUM, border=THIN)
            sc(ws, cr, 5, int_i, size=10, h='right', nf=NUM, border=THIN)
            sc(ws, cr, 6, tot,   size=10, h='right', nf=NUM, border=THIN)
            sum_ces_neto += ces_n; sum_ces_idx += ces_i
            sum_int_neto += int_n; sum_int_idx += int_i
            cr += 1

        # Fila TOTAL cesantías
        ws.row_dimensions[cr].height = 13
        sc(ws, cr, 1, 'TOTAL CESANTÍAS',    bold=True, size=10, bg=GR_M, border=MED)
        sc(ws, cr, 2, sum_ces_neto, bold=True, size=10, bg=GR_M, h='right', nf=NUM, border=MED)
        sc(ws, cr, 3, sum_ces_idx,  bold=True, size=10, bg=GR_M, h='right', nf=NUM, border=MED)
        sc(ws, cr, 4, sum_int_neto, bold=True, size=10, bg=GR_M, h='right', nf=NUM, border=MED)
        sc(ws, cr, 5, sum_int_idx,  bold=True, size=10, bg=GR_M, h='right', nf=NUM, border=MED)
        sc(ws, cr, 6, sum_ces_idx + sum_int_idx, bold=True, size=10, bg=GR_M, h='right', nf=NUM, border=MED)
        cr += 2

    # ── RESUMEN CONSOLIDADO DE LA SENTENCIA ────────────────
    total_prest_idx  = resumen.get('totalIndexado', total_idx_global)
    total_ces_idx    = resumen.get('totalCesTot', 0)
    total_sentencia  = resumen.get('totalSentencia', total_prest_idx + total_ces_idx)

    ws.row_dimensions[cr].height = 6
    cr += 1

    # Bloque resumen en 3 filas
    ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
    sc(ws, cr, 1, 'RESUMEN VALOR DE LA SENTENCIA', bold=True, size=11,
       border=Border(top=Side(style='medium', color='000000'),
                     bottom=Side(style='thin', color='000000')))
    cr += 1

    def fila_resumen(label, valor):
        nonlocal cr
        ws.merge_cells(f'A{cr}:F{cr}')
        sc(ws, cr, 1, label, bold=False, size=11)
        sc(ws, cr, 7, '',    size=11)
        sc(ws, cr, 8, valor, bold=True, size=11, h='right', nf=NUM,
           border=Border(bottom=Side(style='hair', color='595959')))
        ws.row_dimensions[cr].height = 14
        cr += 1

    fila_resumen('Prestaciones sociales indexadas:', total_prest_idx)
    fila_resumen('Cesantías e intereses indexados:', total_ces_idx)

    # Fila TOTAL SENTENCIA
    ws.row_dimensions[cr].height = 16
    ws.merge_cells(f'A{cr}:F{cr}')
    sc(ws, cr, 1, 'VALOR TOTAL SENTENCIA:', bold=True, size=12, bg=GR_M, border=MED)
    sc(ws, cr, 7, '', bg=GR_M, border=MED)
    sc(ws, cr, 8, total_sentencia, bold=True, size=12, bg=GR_M, h='right', nf=NUM, border=MED)
    cr += 2

    # ══════════════════════════════════════════════════════
    # IV. CRITERIOS Y OBSERVACIONES
    # ══════════════════════════════════════════════════════
    ws.row_dimensions[cr].height = 14
    ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
    sc(ws, cr, 1, 'IV.  CRITERIOS Y OBSERVACIONES DE LIQUIDACIÓN', bold=True, size=11)
    cr += 1

    obs_lines = obs.get('texto', [])
    for linea in obs_lines:
        ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
        sc(ws, cr, 1, linea, size=10, wrap=True, color=NG)
        ws.row_dimensions[cr].height = 28 if len(linea) > 120 else 16
        cr += 1

    cr += 1

    # ══════════════════════════════════════════════════════
    # IV. BJ UTILIZADA POR PERÍODO
    # ══════════════════════════════════════════════════════
    ws.row_dimensions[cr].height = 14
    ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
    sc(ws, cr, 1, 'V.  BONIFICACIÓN JUDICIAL (BJ) UTILIZADA POR PERÍODO', bold=True, size=11)
    cr += 1

    bj_hdrs = ['Desde', 'Hasta', 'Cargo / Grado', 'BJ Mensual ($)']
    bj_widths_cols = [1, 2, 3, 4]  # columnas reales
    ws.row_dimensions[cr].height = 13
    for ci, h in enumerate(bj_hdrs, 1):
        sc(ws, cr, ci, h, bold=True, size=10, bg=GR_L,
           h='right' if ci == 4 else 'left', border=THIN)
    cr += 1

    for p in bj_periodos:
        ws.row_dimensions[cr].height = 13
        sc(ws, cr, 1, fmt_fecha(p.get('desde', '')), size=10, border=THIN)
        sc(ws, cr, 2, fmt_fecha(p.get('hasta', '')), size=10, border=THIN)
        sc(ws, cr, 3, p.get('cargo', ''), size=10, border=THIN)
        sc(ws, cr, 4, p.get('bon', 0), size=10, h='right', nf=NUM, border=THIN)
        cr += 1

    cr += 2

    # ══════════════════════════════════════════════════════
    # NOTA FINAL
    # ══════════════════════════════════════════════════════
    nota = (
        'Nota: Los valores consignados en el presente informe están sujetos a cambios en función '
        'de la verificación final de los datos por parte de la entidad pagadora. Documento generado '
        'para efectos de gestión ante fondo de prestaciones sociales — no constituye acto administrativo.'
    )
    ws.merge_cells(f'A{cr}:{LAST_COL}{cr}')
    cell_nota = ws.cell(row=cr, column=1, value=nota)
    cell_nota.font      = Font(name=TNR, size=9, italic=True, color=GR_D)
    cell_nota.alignment = _a(h='left', wrap=True)
    cell_nota.border    = Border(top=Side(style='thin', color='000000'))
    ws.row_dimensions[cr].height = 32

    ws.freeze_panes = 'A2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
