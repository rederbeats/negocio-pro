"""Generador reproducible de NEGOCIO PRO (primera versión)."""
from __future__ import annotations

import argparse
from copy import copy
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.properties import CalcProperties
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.protection import SheetProtection

NAVY = "18324D"
GREEN = "DDEFE6"
GRAY = "E9EDF1"
WHITE = "FFFFFF"
INPUT = "FFF8E5"
RED = "F8D7DA"
EURO = '#,##0.00 "€";[Red]-#,##0.00 "€"'
PERCENT = '0.0%'
DATE = 'DD/MM/YYYY'


def style_sheet(ws, title, headers, widths, last_row):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws.cell(1, 1, title)
    c.fill = PatternFill('solid', fgColor=NAVY)
    c.font = Font(name='Aptos Display', size=18, bold=True, color=WHITE)
    c.alignment = Alignment(vertical='center')
    ws.row_dimensions[1].height = 36
    for j, (header, width) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(4, j, header)
        cell.fill = PatternFill('solid', fgColor=NAVY)
        cell.font = Font(name='Aptos', bold=True, color=WHITE)
        cell.alignment = Alignment(wrap_text=True, vertical='center')
        ws.column_dimensions[get_column_letter(j)].width = width
    ws.row_dimensions[4].height = 34
    ws.freeze_panes = 'A5'
    ws.auto_filter.ref = f'A4:{get_column_letter(len(headers))}{last_row}'
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def format_body(ws, first, last, ncols, editable, formats=None):
    formats = formats or {}
    for row in ws.iter_rows(min_row=first, max_row=last, max_col=ncols):
        for c in row:
            c.font = Font(name='Aptos', size=10, color=NAVY)
            c.fill = PatternFill('solid', fgColor=INPUT if c.column in editable else (WHITE if c.row % 2 else 'F5F7F9'))
            protection = copy(c.protection)
            protection.locked = c.column not in editable
            c.protection = protection
            c.number_format = formats.get(c.column, 'General')
    ws.protection = SheetProtection(sheet=True, selectLockedCells=False, selectUnlockedCells=False,
                                    autoFilter=False, sort=False, formatCells=False)


def add_list(ws, formula, target, error='Selecciona un valor de la lista.'):
    dv = DataValidation(type='list', formula1=f'={formula}', allow_blank=True)
    dv.error = error
    dv.errorTitle = 'Dato no válido'
    dv.showErrorMessage = True
    dv.errorStyle = 'stop'
    ws.add_data_validation(dv)
    dv.add(target)


def add_decimal(ws, target, lower='0', upper='999999999'):
    dv = DataValidation(type='decimal', operator='between', formula1=lower, formula2=upper, allow_blank=True)
    dv.showErrorMessage = True
    dv.error = f'Introduce un número entre {lower} y {upper}.'
    ws.add_data_validation(dv)
    dv.add(target)


def add_custom(ws, target, formula, message):
    dv = DataValidation(type='custom', formula1=formula, allow_blank=True)
    dv.showErrorMessage = True
    dv.errorStyle = 'stop'
    dv.error = message
    ws.add_data_validation(dv)
    dv.add(target)


def add_date(ws, target):
    dv = DataValidation(type='date', operator='between', formula1='DATE(2000,1,1)', formula2='DATE(2100,12,31)', allow_blank=True)
    dv.showErrorMessage = True
    dv.errorStyle = 'stop'
    dv.error = 'Introduce una fecha válida entre 01/01/2000 y 31/12/2100.'
    ws.add_data_validation(dv)
    dv.add(target)


def flag(ws, formula, target):
    ws.conditional_formatting.add(target, FormulaRule(formula=[formula], fill=PatternFill('solid', fgColor=RED)))


def build(output: Path, demo=False, ventas=500, gastos=500, productos=200, clientes=200, proveedores=200, movimientos=500, caja=500, objetivos=120):
    if min(ventas, gastos, productos, clientes, proveedores, movimientos, caja, objetivos) < 2:
        raise ValueError('Cada capacidad debe ser al menos 2')
    output.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    dash = wb.active
    dash.title = 'Dashboard'
    vs = wb.create_sheet('Ventas')
    gs = wb.create_sheet('Gastos')
    ps = wb.create_sheet('Productos')
    cs = wb.create_sheet('Clientes')
    ss = wb.create_sheet('Proveedores')
    inv = wb.create_sheet('Inventario')
    cash = wb.create_sheet('Tesorería')
    goals = wb.create_sheet('Objetivos')
    reports = wb.create_sheet('Informes')
    cfg = wb.create_sheet('Configuración')
    wb.calculation = CalcProperties(calcMode='auto', fullCalcOnLoad=True, forceFullCalc=True)
    ve, ge, pe = 4 + ventas, 4 + gastos, 4 + productos
    ce, se, ie = 4 + clientes, 4 + proveedores, 4 + movimientos
    te, oe = 4 + caja, 4 + objetivos

    style_sheet(cfg, 'Configuración', ['Parámetro', 'Valor', 'Categorías de gasto', 'Tipos IVA'], [32, 22, 28, 18], 20)
    cfg['A5'], cfg['B5'] = 'Empresa', 'Mi negocio'
    cfg['A6'], cfg['B6'] = 'Ejercicio', 2026
    cfg['A7'], cfg['B7'] = 'Moneda', 'EUR'
    cfg['A9'] = 'Entradas: amarillo claro · Cálculos: blanco/gris'
    cfg['A10'] = 'Todos los importes de operación se introducen sin IVA.'
    cfg['A11'] = 'Demo: datos ficticios; no usar para decisiones reales.' if demo else 'Plantilla limpia: introduce tus datos en las hojas de entrada.'
    for i, cat in enumerate(['Material', 'Alquiler', 'Servicios', 'Marketing', 'Transporte', 'Personal', 'Otros'], 5):
        cfg.cell(i, 3, cat)
    for i, rate in enumerate([0, 0.04, 0.10, 0.21], 5):
        cfg.cell(i, 4, rate).number_format = PERCENT
    wb.defined_names.add(DefinedName('CategoriasGasto', attr_text="'Configuración'!$C$5:$C$11"))
    wb.defined_names.add(DefinedName('TiposIVA', attr_text="'Configuración'!$D$5:$D$8"))
    wb.defined_names.add(DefinedName('IdsProducto', attr_text=f"Productos!$A$5:$A${pe}"))
    wb.defined_names.add(DefinedName('IdsCliente', attr_text=f"Clientes!$A$5:$A${ce}"))
    wb.defined_names.add(DefinedName('IdsProveedor', attr_text=f"Proveedores!$A$5:$A${se}"))
    wb.defined_names.add(DefinedName('EstadosEntidad', attr_text="'Configuración'!$E$5:$E$6"))
    wb.defined_names.add(DefinedName('TiposProducto', attr_text="'Configuración'!$F$5:$F$7"))
    wb.defined_names.add(DefinedName('TiposMovimiento', attr_text="'Configuración'!$G$5:$G$8"))
    wb.defined_names.add(DefinedName('CuentasCaja', attr_text="'Configuración'!$H$5:$H$14"))
    wb.defined_names.add(DefinedName('TiposCaja', attr_text="'Configuración'!$K$5:$K$6"))
    for col, values in [(5, ['Activo', 'Inactivo']), (6, ['Físico', 'Servicio', 'Sin control']),
                        (7, ['Entrada', 'Salida', 'Ajuste +', 'Ajuste -'])]:
        for row, value in enumerate(values, 5):
            cfg.cell(row, col, value)
    for col, header, width in [(5, 'Estado entidad', 20), (6, 'Tipo producto', 20), (7, 'Movimiento', 20)]:
        cfg.cell(4, col, header)
        cfg.cell(4, col).fill = PatternFill('solid', fgColor=NAVY)
        cfg.cell(4, col).font = Font(name='Aptos', bold=True, color=WHITE)
        cfg.column_dimensions[get_column_letter(col)].width = width
    for col, header, width in [(8, 'Cuenta de caja', 22), (9, 'Saldo inicial', 20), (10, 'Saldo final', 20), (11, 'Tipo caja', 16)]:
        cfg.cell(4, col, header)
        cfg.cell(4, col).fill = PatternFill('solid', fgColor=NAVY)
        cfg.cell(4, col).font = Font(name='Aptos', bold=True, color=WHITE)
        cfg.column_dimensions[get_column_letter(col)].width = width
    for row, account in enumerate(('Banco', 'Efectivo'), 5):
        cfg.cell(row, 8, account)
        cfg.cell(row, 9, 0).number_format = EURO
    cfg['K5'], cfg['K6'] = 'Cobro', 'Pago'
    for row in range(5, 15):
        for col in (8, 9):
            cell = cfg.cell(row, col)
            protection = copy(cell.protection)
            protection.locked = False
            cell.protection = protection
            cell.fill = PatternFill('solid', fgColor=INPUT)
        cfg.cell(row, 10, f'=IF(H{row}="","",I{row}+SUMIFS(Tesorería!$I$5:$I${te},Tesorería!$F$5:$F${te},H{row},Tesorería!$B$5:$B${te},">="&DATE($B$6,1,1),Tesorería!$B$5:$B${te},"<"&DATE($B$6+1,1,1))-SUMIFS(Tesorería!$J$5:$J${te},Tesorería!$F$5:$F${te},H{row},Tesorería!$B$5:$B${te},">="&DATE($B$6,1,1),Tesorería!$B$5:$B${te},"<"&DATE($B$6+1,1,1)))').number_format = EURO
    add_custom(cfg, 'H5:H14', 'OR(H5="",COUNTIF($H$5:$H$14,H5)=1)', 'La cuenta debe tener un nombre único.')
    add_decimal(cfg, 'I5:I14')
    flag(cfg, 'AND($H5<>"",NOT(ISNUMBER($I5)))', 'H5:I14')
    for row in cfg.iter_rows(min_row=5, max_row=20, max_col=4):
        for c in row:
            c.font = Font(name='Aptos', color=NAVY)
            protection = copy(c.protection)
            protection.locked = c.coordinate not in {'B5', 'B6'}
            c.protection = protection
    for pos in ('B5', 'B6'):
        cfg[pos].fill = PatternFill('solid', fgColor=INPUT)
    add_custom(cfg, 'B6', 'AND(ISNUMBER(B6),B6=INT(B6),B6>=2000,B6<=2100)', 'El ejercicio debe ser un año entre 2000 y 2100.')
    cfg.protection = SheetProtection(sheet=True, selectLockedCells=False, selectUnlockedCells=True)

    style_sheet(ps, 'Productos', ['ID producto', 'Nombre', 'Categoría', 'Precio neto', 'Coste neto', 'Margen unitario', 'Margen %', 'Tipo', 'Stock inicial', 'Stock mínimo', 'Stock disponible', 'Alerta', 'Estado'], [19, 32, 24, 19, 19, 20, 16, 18, 17, 17, 20, 18, 18], pe)
    ps['A2'] = 'Solo los productos físicos tienen stock. Registra cada salida en Inventario; Ventas no descuenta automáticamente.'
    format_body(ps, 5, pe, 13, {1, 2, 3, 4, 5, 8, 9, 10}, {4: EURO, 5: EURO, 6: EURO, 7: PERCENT, 9: '#,##0.00', 10: '#,##0.00', 11: '#,##0.00'})
    for r in range(5, pe + 1):
        ps.cell(r, 6, f'=IF(OR(A{r}="",D{r}="",E{r}=""),"",D{r}-E{r})')
        ps.cell(r, 7, f'=IF(OR(A{r}="",D{r}="",E{r}="",D{r}=0),"",F{r}/D{r})')
        ps.cell(r, 13, f'=IF(COUNTA(A{r}:E{r},H{r}:J{r})=0,"",IF(AND(A{r}<>"",COUNTIF($A$5:$A${pe},A{r})=1,B{r}<>"",ISNUMBER(D{r}),D{r}>=0,ISNUMBER(E{r}),E{r}>=0,OR(H{r}="",COUNTIF(TiposProducto,H{r})=1),IF(H{r}="Físico",AND(ISNUMBER(I{r}),I{r}>=0,ISNUMBER(J{r}),J{r}>=0),AND(I{r}="",J{r}=""))),"OK","REVISAR"))')
        ps.cell(r, 11, f'=IF(OR(M{r}<>"OK",H{r}<>"Físico"),"",I{r}+SUMIFS(Inventario!$H$5:$H${ie},Inventario!$C$5:$C${ie},A{r},Inventario!$I$5:$I${ie},"OK"))')
        ps.cell(r, 12, f'=IF(K{r}="","",IF(K{r}<0,"NEGATIVO",IF(K{r}<J{r},"BAJO","OK")))')
    add_decimal(ps, f'D5:E{pe}')
    add_decimal(ps, f'I5:J{pe}')
    add_list(ps, 'TiposProducto', f'H5:H{pe}')
    add_custom(ps, f'A5:A{pe}', f'OR(A5="",COUNTIF($A$5:$A${pe},A5)=1)', 'El ID de producto debe ser único.')
    flag(ps, f'AND($A5<>"",COUNTIF($A$5:$A${pe},$A5)>1)', f'A5:A{pe}')
    flag(ps, 'AND($A5<>"",OR($B5="",$D5="",$E5=""))', f'A5:E{pe}')
    flag(ps, '$M5="REVISAR"', f'A5:M{pe}')
    flag(ps, 'OR($L5="BAJO",$L5="NEGATIVO")', f'L5:L{pe}')

    style_sheet(cs, 'Clientes', ['ID cliente', 'Nombre / razón social', 'Email', 'Teléfono', 'Localidad', 'Fecha de alta', 'Observaciones', 'Situación', 'Estado', 'Facturación neta', 'Compras'], [19, 32, 30, 20, 23, 19, 35, 18, 18, 22, 15], ce)
    cs['A2'] = 'ID único y nombre obligatorios. El cliente es opcional en Ventas. Facturación y compras: ejercicio elegido, sin IVA.'
    format_body(cs, 5, ce, 11, set(range(1, 9)), {6: DATE, 10: EURO})
    for r in range(5, ce + 1):
        cs.cell(r, 9, f'=IF(COUNTA(A{r}:H{r})=0,"",IF(AND(A{r}<>"",COUNTIF($A$5:$A${ce},A{r})=1,B{r}<>"",COUNTIF(EstadosEntidad,H{r})=1,OR(C{r}="",COUNTIF(C{r},"?*@?*.?*")=1),OR(F{r}="",ISNUMBER(F{r}))),"OK","REVISAR"))')
        cs.cell(r, 10, f'=IF(I{r}<>"OK","",SUMIFS(Ventas!$H$5:$H${ve},Ventas!$D$5:$D${ve},A{r},Ventas!$O$5:$O${ve},"OK"))').number_format = EURO
        cs.cell(r, 11, f'=IF(I{r}<>"OK","",COUNTIFS(Ventas!$D$5:$D${ve},A{r},Ventas!$O$5:$O${ve},"OK"))')
    add_custom(cs, f'A5:A{ce}', f'OR(A5="",COUNTIF($A$5:$A${ce},A5)=1)', 'El ID de cliente debe ser único.')
    add_list(cs, 'EstadosEntidad', f'H5:H{ce}')
    add_date(cs, f'F5:F{ce}')
    flag(cs, '$I5="REVISAR"', f'A5:K{ce}')

    style_sheet(ss, 'Proveedores', ['ID proveedor', 'Nombre comercial', 'Contacto', 'Email', 'Teléfono', 'Categoría', 'Observaciones', 'Situación', 'Estado', 'Gasto neto'], [20, 32, 28, 30, 20, 24, 35, 18, 18, 22], se)
    ss['A2'] = 'ID único y nombre obligatorios. El proveedor es opcional en Gastos. Gasto neto: ejercicio elegido, sin IVA.'
    format_body(ss, 5, se, 10, set(range(1, 9)), {10: EURO})
    for r in range(5, se + 1):
        ss.cell(r, 9, f'=IF(COUNTA(A{r}:H{r})=0,"",IF(AND(A{r}<>"",COUNTIF($A$5:$A${se},A{r})=1,B{r}<>"",COUNTIF(EstadosEntidad,H{r})=1,OR(D{r}="",COUNTIF(D{r},"?*@?*.?*")=1)),"OK","REVISAR"))')
        ss.cell(r, 10, f'=IF(I{r}<>"OK","",SUMIFS(Gastos!$F$5:$F${ge},Gastos!$D$5:$D${ge},A{r},Gastos!$K$5:$K${ge},"OK"))').number_format = EURO
    add_custom(ss, f'A5:A{se}', f'OR(A5="",COUNTIF($A$5:$A${se},A5)=1)', 'El ID de proveedor debe ser único.')
    add_list(ss, 'EstadosEntidad', f'H5:H{se}')
    flag(ss, '$I5="REVISAR"', f'A5:J{se}')

    style_sheet(inv, 'Inventario · movimientos de stock', ['Fecha', 'ID movimiento', 'ID producto', 'Tipo', 'Cantidad', 'Referencia', 'Observaciones', 'Impacto stock', 'Estado'], [16, 21, 20, 18, 16, 24, 36, 20, 20], ie)
    inv['A2'] = 'Solo productos físicos. Una venta física requiere UNA salida manual; cada Salida exige referencia única.'
    format_body(inv, 5, ie, 9, {1, 2, 3, 4, 5, 6, 7}, {1: DATE, 5: '#,##0.00', 8: '#,##0.00'})
    for r in range(5, ie + 1):
        inv.cell(r, 9, f'=IF(COUNTA(A{r}:G{r})=0,"",IF(AND(ISNUMBER(A{r}),B{r}<>"",COUNTIF($B$5:$B${ie},B{r})=1,C{r}<>"",COUNTIF(IdsProducto,C{r})=1,COUNTIFS(Productos!$A$5:$A${pe},C{r},Productos!$H$5:$H${pe},"Físico",Productos!$M$5:$M${pe},"OK")=1,COUNTIF(TiposMovimiento,D{r})=1,ISNUMBER(E{r}),E{r}>0,IF(D{r}="Salida",AND(F{r}<>"",COUNTIFS($D$5:$D${ie},"Salida",$F$5:$F${ie},F{r})=1),TRUE)),"OK","REVISAR"))')
        inv.cell(r, 8, f'=IF(I{r}<>"OK","",IF(OR(D{r}="Entrada",D{r}="Ajuste +"),E{r},-E{r}))')
    add_custom(inv, f'B5:B{ie}', f'OR(B5="",COUNTIF($B$5:$B${ie},B5)=1)', 'El ID de movimiento debe ser único.')
    add_custom(inv, f'C5:C{ie}', 'OR(C5="",COUNTIF(IdsProducto,C5)=1)', 'El ID de producto debe existir una sola vez.')
    add_list(inv, 'TiposMovimiento', f'D5:D{ie}')
    add_decimal(inv, f'E5:E{ie}', '0.000001')
    add_date(inv, f'A5:A{ie}')
    flag(inv, '$I5="REVISAR"', f'A5:I{ie}')

    style_sheet(vs, 'Ventas · importes sin IVA', ['Fecha', 'ID venta', 'ID producto', 'Cliente ID', 'Cantidad', 'Precio unit. neto', 'Descuento %', 'Base neta', 'IVA %', 'IVA repercutido', 'Total con IVA', 'Coste unit. neto', 'Coste estimado', 'Referencia', 'Estado', 'Cobrado', 'Pendiente'], [16, 18, 18, 18, 13, 21, 16, 20, 13, 21, 21, 21, 20, 21, 21, 20, 20], ve)
    vs['A2'] = 'Amarillo: entrada · blanco/gris: cálculo · Estado REVISAR/FUERA: no entra en Dashboard. Precio y coste son históricos.'
    format_body(vs, 5, ve, 17, {1, 2, 3, 4, 5, 6, 7, 9, 12, 14}, {1: DATE, 6: EURO, 7: PERCENT, 8: EURO, 9: PERCENT, 10: EURO, 11: EURO, 12: EURO, 13: EURO, 16: EURO, 17: EURO})
    for r in range(5, ve + 1):
        vs.cell(r, 15, f'=IF(COUNTA(A{r}:G{r},I{r},L{r},N{r})=0,"",IF(AND(ISNUMBER(A{r}),ISNUMBER(E{r}),E{r}>0,ISNUMBER(F{r}),F{r}>=0,ISNUMBER(G{r}),G{r}>=0,G{r}<=1,ISNUMBER(I{r}),COUNTIF(TiposIVA,I{r})=1,ISNUMBER(L{r}),L{r}>=0,B{r}<>"",COUNTIF($B$5:$B${ve},B{r})=1,C{r}<>"",COUNTIF(IdsProducto,C{r})=1,COUNTIFS(Productos!$A$5:$A${pe},C{r},Productos!$M$5:$M${pe},"OK")=1,OR(D{r}="",COUNTIFS(Clientes!$A$5:$A${ce},D{r},Clientes!$I$5:$I${ce},"OK")=1)),IF(YEAR(A{r})=\'Configuración\'!$B$6,"OK","FUERA DE EJERCICIO"),"REVISAR"))')
        vs.cell(r, 8, f'=IF(AND(O{r}<>"OK",O{r}<>"FUERA DE EJERCICIO"),"",ROUND(E{r}*F{r}*(1-G{r}),2))')
        vs.cell(r, 10, f'=IF(H{r}="","",ROUND(H{r}*I{r},2))')
        vs.cell(r, 11, f'=IF(H{r}="","",H{r}+J{r})')
        vs.cell(r, 13, f'=IF(H{r}="","",ROUND(E{r}*L{r},2))')
        vs.cell(r, 16, f'=IF(K{r}="","",SUMIFS(Tesorería!$I$5:$I${te},Tesorería!$G$5:$G${te},B{r}))')
        vs.cell(r, 17, f'=IF(K{r}="","",K{r}-P{r})')
    add_decimal(vs, f'E5:E{ve}', '0.000001')
    add_decimal(vs, f'F5:F{ve}')
    add_decimal(vs, f'G5:G{ve}', '0', '1')
    add_decimal(vs, f'L5:L{ve}')
    add_list(vs, 'TiposIVA', f'I5:I{ve}', 'Selecciona un tipo de IVA de Configuración.')
    add_date(vs, f'A5:A{ve}')
    add_custom(vs, f'B5:B{ve}', f'OR(B5="",COUNTIF($B$5:$B${ve},B5)=1)', 'El ID de venta debe ser único.')
    add_custom(vs, f'C5:C{ve}', 'OR(C5="",COUNTIF(IdsProducto,C5)=1)', 'El ID debe existir una sola vez en Productos.')
    add_custom(vs, f'D5:D{ve}', 'OR(D5="",COUNTIF(IdsCliente,D5)=1)', 'El cliente es opcional; si se indica, debe existir una sola vez.')
    flag(vs, f'AND($B5<>"",COUNTIF($B$5:$B${ve},$B5)>1)', f'B5:B{ve}')
    flag(vs, f'AND($C5<>"",COUNTIF(Productos!$A$5:$A${pe},$C5)=0)', f'C5:C{ve}')
    flag(vs, '$O5="REVISAR"', f'A5:Q{ve}')
    flag(vs, '$O5="FUERA DE EJERCICIO"', f'O5:O{ve}')

    style_sheet(gs, 'Gastos · importes sin IVA', ['Fecha', 'ID gasto', 'Categoría', 'Proveedor ID', 'Concepto', 'Base neta', 'IVA %', 'IVA soportado', 'Total con IVA', 'Referencia', 'Estado', 'Pagado', 'Pendiente'], [16, 18, 22, 19, 34, 19, 13, 20, 20, 20, 21, 20, 20], ge)
    gs['A2'] = 'Amarillo: entrada · blanco/gris: cálculo · Estado REVISAR/FUERA: no entra en Dashboard.'
    format_body(gs, 5, ge, 13, {1, 2, 3, 4, 5, 6, 7, 10}, {1: DATE, 6: EURO, 7: PERCENT, 8: EURO, 9: EURO, 12: EURO, 13: EURO})
    for r in range(5, ge + 1):
        gs.cell(r, 11, f'=IF(COUNTA(A{r}:G{r},J{r})=0,"",IF(AND(ISNUMBER(A{r}),B{r}<>"",COUNTIF($B$5:$B${ge},B{r})=1,C{r}<>"",COUNTIF(CategoriasGasto,C{r})=1,ISNUMBER(F{r}),F{r}>=0,ISNUMBER(G{r}),COUNTIF(TiposIVA,G{r})=1,OR(D{r}="",COUNTIFS(Proveedores!$A$5:$A${se},D{r},Proveedores!$I$5:$I${se},"OK")=1)),IF(YEAR(A{r})=\'Configuración\'!$B$6,"OK","FUERA DE EJERCICIO"),"REVISAR"))')
        gs.cell(r, 8, f'=IF(AND(K{r}<>"OK",K{r}<>"FUERA DE EJERCICIO"),"",ROUND(F{r}*G{r},2))')
        gs.cell(r, 9, f'=IF(H{r}="","",F{r}+H{r})')
        gs.cell(r, 12, f'=IF(I{r}="","",SUMIFS(Tesorería!$J$5:$J${te},Tesorería!$G$5:$G${te},B{r}))')
        gs.cell(r, 13, f'=IF(I{r}="","",I{r}-L{r})')
    add_list(gs, 'CategoriasGasto', f'C5:C{ge}', 'Selecciona una categoría de Configuración.')
    add_list(gs, 'TiposIVA', f'G5:G{ge}', 'Selecciona un tipo de IVA de Configuración.')
    add_date(gs, f'A5:A{ge}')
    add_decimal(gs, f'F5:F{ge}')
    add_custom(gs, f'B5:B{ge}', f'OR(B5="",COUNTIF($B$5:$B${ge},B5)=1)', 'El ID de gasto debe ser único.')
    add_custom(gs, f'D5:D{ge}', 'OR(D5="",COUNTIF(IdsProveedor,D5)=1)', 'El proveedor es opcional; si se indica, debe existir una sola vez.')
    flag(gs, f'AND($B5<>"",COUNTIF($B$5:$B${ge},$B5)>1)', f'B5:B{ge}')
    flag(gs, '$K5="REVISAR"', f'A5:M{ge}')
    flag(gs, '$K5="FUERA DE EJERCICIO"', f'K5:K{ge}')

    style_sheet(cash, 'Tesorería · cobros y pagos reales', ['ID caja', 'Fecha efectiva', 'Tipo', 'Concepto', 'Importe con IVA', 'Cuenta', 'ID venta/gasto', 'Estado', 'Entrada', 'Salida'], [19, 18, 16, 32, 21, 20, 21, 21, 20, 20], te)
    cash['A2'] = 'Fecha efectiva de caja; importes con IVA. Referencia opcional: Cobro→ID venta, Pago→ID gasto. No afecta beneficio.'
    format_body(cash, 5, te, 10, {1, 2, 3, 4, 5, 6, 7}, {2: DATE, 5: EURO, 9: EURO, 10: EURO})
    for r in range(5, te + 1):
        sale_ok = f'COUNTIFS(Ventas!$B$5:$B${ve},G{r},Ventas!$O$5:$O${ve},"OK")+COUNTIFS(Ventas!$B$5:$B${ve},G{r},Ventas!$O$5:$O${ve},"FUERA DE EJERCICIO")=1'
        expense_ok = f'COUNTIFS(Gastos!$B$5:$B${ge},G{r},Gastos!$K$5:$K${ge},"OK")+COUNTIFS(Gastos!$B$5:$B${ge},G{r},Gastos!$K$5:$K${ge},"FUERA DE EJERCICIO")=1'
        sale_total = f'SUMIFS(Ventas!$K$5:$K${ve},Ventas!$B$5:$B${ve},G{r})'
        expense_total = f'SUMIFS(Gastos!$I$5:$I${ge},Gastos!$B$5:$B${ge},G{r})'
        linked_total = f'SUMIFS($E$5:$E${te},$G$5:$G${te},G{r},$C$5:$C${te},C{r})'
        cash.cell(r, 8, f'=IF(COUNTA(A{r}:G{r})=0,"",IF(AND(A{r}<>"",COUNTIF($A$5:$A${te},A{r})=1,ISNUMBER(B{r}),COUNTIF(TiposCaja,C{r})=1,D{r}<>"",ISNUMBER(E{r}),E{r}>0,COUNTIF(CuentasCaja,F{r})=1,COUNTIFS(\'Configuración\'!$H$5:$H$14,F{r},\'Configuración\'!$I$5:$I$14,"<>")=1,IF(G{r}="",TRUE,IF(C{r}="Cobro",AND({sale_ok},{linked_total}<={sale_total}),AND({expense_ok},{linked_total}<={expense_total})))),IF(YEAR(B{r})=\'Configuración\'!$B$6,"OK","FUERA DE EJERCICIO"),"REVISAR"))')
        cash.cell(r, 9, f'=IF(AND(OR(H{r}="OK",H{r}="FUERA DE EJERCICIO"),C{r}="Cobro"),E{r},"")')
        cash.cell(r, 10, f'=IF(AND(OR(H{r}="OK",H{r}="FUERA DE EJERCICIO"),C{r}="Pago"),E{r},"")')
    add_custom(cash, f'A5:A{te}', f'OR(A5="",COUNTIF($A$5:$A${te},A5)=1)', 'El ID de caja debe ser único.')
    add_date(cash, f'B5:B{te}')
    add_list(cash, 'TiposCaja', f'C5:C{te}')
    add_decimal(cash, f'E5:E{te}', '0.000001')
    add_list(cash, 'CuentasCaja', f'F5:F{te}')
    flag(cash, '$H5="REVISAR"', f'A5:J{te}')
    for col, label, width in [(12, 'Mes', 19), (13, 'Saldo inicial', 20), (14, 'Entradas', 20), (15, 'Salidas', 20), (16, 'Saldo final', 20)]:
        cell = cash.cell(4, col, label)
        cell.fill = PatternFill('solid', fgColor=NAVY)
        cell.font = Font(name='Aptos', bold=True, color=WHITE)
        cash.column_dimensions[get_column_letter(col)].width = width
    for row in range(5, 17):
        cash.cell(row, 12, f'=DATE(\'Configuración\'!$B$6,{row-4},1)').number_format = 'MMMM YYYY'
        cash.cell(row, 13, '=SUM(\'Configuración\'!$I$5:$I$14)' if row == 5 else f'=P{row-1}').number_format = EURO
        for col, source in [(14, 'I'), (15, 'J')]:
            cash.cell(row, col, f'=SUMIFS(${source}$5:${source}${te},$H$5:$H${te},"OK",$B$5:$B${te},">="&L{row},$B$5:$B${te},"<"&EDATE(L{row},1))').number_format = EURO
        cash.cell(row, 16, f'=M{row}+N{row}-O{row}').number_format = EURO

    style_sheet(goals, 'Objetivos · metas por ejercicio y mes', ['Ejercicio', 'Mes (1-12)', 'Meta ventas', 'Ventas reales', 'Cumpl. ventas', 'Desv. ventas', 'Meta beneficio', 'Beneficio real', 'Cumpl. beneficio', 'Desv. beneficio', 'Límite gastos', 'Gastos reales', 'Control gastos', 'Ahorro vs límite', 'Estado', 'Actividad'], [16, 16, 20, 20, 18, 20, 20, 20, 20, 20, 20, 20, 18, 20, 19, 19], oe)
    goals['A2'] = 'Una fila por ejercicio y mes. Metas sin IVA. Vacío = SIN META; meta 0 = porcentaje no aplicable.'
    format_body(goals, 5, oe, 16, {1, 2, 3, 7, 11}, {3: EURO, 4: EURO, 5: PERCENT, 6: EURO, 7: EURO, 8: EURO, 9: PERCENT, 10: EURO, 11: EURO, 12: EURO, 13: PERCENT, 14: EURO})
    for r in range(5, oe + 1):
        start = f'DATE(A{r},B{r},1)'
        end = f'EDATE({start},1)'
        sales_sum = f'SUMIFS(Ventas!$H$5:$H${ve},Ventas!$A$5:$A${ve},">="&{start},Ventas!$A$5:$A${ve},"<"&{end})'
        cost_sum = f'SUMIFS(Ventas!$M$5:$M${ve},Ventas!$A$5:$A${ve},">="&{start},Ventas!$A$5:$A${ve},"<"&{end})'
        expense_sum = '+'.join(f'SUMIFS(Gastos!$F$5:$F${ge},Gastos!$A$5:$A${ge},">="&{start},Gastos!$A$5:$A${ge},"<"&{end},Gastos!$K$5:$K${ge},"{status}")' for status in ('OK', 'FUERA DE EJERCICIO'))
        goals.cell(r, 15, f'=IF(COUNTA(A{r}:C{r},G{r},K{r})=0,"",IF(AND(ISNUMBER(A{r}),IFERROR(A{r}=INT(A{r}),FALSE),A{r}>=2000,A{r}<=2100,ISNUMBER(B{r}),IFERROR(B{r}=INT(B{r}),FALSE),B{r}>=1,B{r}<=12,COUNTIFS($A$5:$A${oe},A{r},$B$5:$B${oe},B{r})=1,OR(C{r}="",AND(ISNUMBER(C{r}),C{r}>=0)),OR(G{r}="",AND(ISNUMBER(G{r}),G{r}>=0)),OR(K{r}="",AND(ISNUMBER(K{r}),K{r}>=0))),IF(COUNT(C{r},G{r},K{r})=0,"SIN META","OK"),"REVISAR"))')
        activity = '+'.join(f'COUNTIFS({sheet}!$A$5:$A${limit},">="&{start},{sheet}!$A$5:$A${limit},"<"&{end},{sheet}!${status_col}$5:${status_col}${limit},"{status}")' for sheet, limit, status_col in [('Ventas', ve, 'O'), ('Gastos', ge, 'K')] for status in ('OK', 'FUERA DE EJERCICIO'))
        goals.cell(r, 16, f'=IF(OR(O{r}="OK",O{r}="SIN META"),IF({activity}=0,"SIN ACTIVIDAD","CON ACTIVIDAD"),"")')
        goals.cell(r, 4, f'=IF(OR(O{r}="OK",O{r}="SIN META"),{sales_sum},"")')
        goals.cell(r, 12, f'=IF(OR(O{r}="OK",O{r}="SIN META"),{expense_sum},"")')
        goals.cell(r, 8, f'=IF(D{r}="","",D{r}-{cost_sum}-L{r})')
        for col, target, actual, direction in [(5, 'C', 'D', 1), (9, 'G', 'H', 1), (13, 'K', 'L', -1)]:
            if direction == 1:
                expr = f'{actual}{r}/{target}{r}'
            else:
                expr = f'IF({actual}{r}=0,1,{target}{r}/{actual}{r})'
            goals.cell(r, col, f'=IF({target}{r}="","SIN META",IF(OR({actual}{r}="",P{r}="SIN ACTIVIDAD"),"SIN DATOS",IF({target}{r}=0,"n.a.",{expr})))')
        for col, target, actual, direction in [(6, 'C', 'D', 1), (10, 'G', 'H', 1), (14, 'K', 'L', -1)]:
            expr = f'{actual}{r}-{target}{r}' if direction == 1 else f'{target}{r}-{actual}{r}'
            goals.cell(r, col, f'=IF({target}{r}="","SIN META",IF(OR({actual}{r}="",P{r}="SIN ACTIVIDAD"),"SIN DATOS",{expr}))')
    add_custom(goals, f'A5:A{oe}', 'OR(A5="",AND(ISNUMBER(A5),IFERROR(A5=INT(A5),FALSE),A5>=2000,A5<=2100))', 'Introduce un ejercicio entre 2000 y 2100.')
    add_custom(goals, f'B5:B{oe}', 'OR(B5="",AND(ISNUMBER(B5),IFERROR(B5=INT(B5),FALSE),B5>=1,B5<=12))', 'Introduce un mes del 1 al 12.')
    for col in ('C', 'G', 'K'):
        add_decimal(goals, f'{col}5:{col}{oe}')
    flag(goals, '$O5="REVISAR"', f'A5:O{oe}')
    for col in ('E', 'I', 'M'):
        goals.conditional_formatting.add(f'{col}5:{col}{oe}', FormulaRule(formula=[f'AND(ISNUMBER({col}5),{col}5>=1)'], fill=PatternFill('solid', fgColor=GREEN)))

    style_sheet(reports, 'Informes · cifras de gestión sin IVA', ['Mes', 'Ventas netas', 'Coste producto', 'Gastos netos', 'Beneficio estimado', 'Margen'], [19, 21, 21, 21, 23, 17], 18)
    reports['A2'] = 'Gestión por fecha de operación. Tesorería usa fecha efectiva y total con IVA; no es resultado fiscal.'
    for row in range(5, 17):
        reports.cell(row, 1, f'=Dashboard!C{row}').number_format = 'MMMM YYYY'
        for col, source in [(2, 'D'), (3, 'E'), (4, 'F'), (5, 'G')]:
            reports.cell(row, col, f'=Dashboard!{source}{row}').number_format = EURO
        reports.cell(row, 6, f'=IF(B{row}=0,"",E{row}/B{row})').number_format = PERCENT
    reports['A18'] = 'TOTAL ANUAL'
    for col, source in [(2, 'B5'), (3, 'B6'), (4, 'B7'), (5, 'B8'), (6, 'B9')]:
        reports.cell(18, col, f'=Dashboard!{source}').number_format = PERCENT if col == 6 else EURO
    for col, label, width in [(8, 'Categoría gasto', 24), (9, 'Gasto neto', 20), (11, 'ID producto', 20), (12, 'Producto', 28), (13, 'Ventas netas', 20), (14, 'Coste estimado', 20), (15, 'Margen', 20)]:
        reports.cell(4, col, label)
        reports.cell(4, col).fill = PatternFill('solid', fgColor=NAVY)
        reports.cell(4, col).font = Font(name='Aptos', bold=True, color=WHITE)
        reports.column_dimensions[get_column_letter(col)].width = width
    for row in range(5, 12):
        reports.cell(row, 8, f'=Dashboard!I{row}')
        reports.cell(row, 9, f'=Dashboard!J{row}').number_format = EURO
    for row in range(5, pe + 1):
        reports.cell(row, 11, f'=IF(Productos!M{row}="OK",Productos!A{row},"")')
        reports.cell(row, 12, f'=IF(K{row}="","",Productos!B{row})')
        for col, amount in [(13, 'H'), (14, 'M')]:
            reports.cell(row, col, f'=IF(K{row}="","",SUMIFS(Ventas!${amount}$5:${amount}${ve},Ventas!$C$5:$C${ve},K{row},Ventas!$O$5:$O${ve},"OK"))').number_format = EURO
        reports.cell(row, 15, f'=IF(K{row}="","",M{row}-N{row})').number_format = EURO
    chart = BarChart()
    chart.title = 'Evolución de ventas y beneficio · sin IVA'
    for col in (2, 5):
        chart.add_data(Reference(reports, min_col=col, min_row=4, max_row=16), titles_from_data=True)
    chart.set_categories(Reference(reports, min_col=1, min_row=5, max_row=16))
    chart.width, chart.height = 20, 10
    reports.add_chart(chart, 'A21')
    reports.protection = SheetProtection(sheet=True, selectLockedCells=False, selectUnlockedCells=False, autoFilter=False)

    style_sheet(dash, 'NEGOCIO PRO · Dashboard', ['Indicador', 'Valor', 'Mes', 'Ventas netas', 'Coste producto', 'Gastos netos', 'Beneficio estimado'], [29, 22, 19, 20, 20, 20, 23], 30)
    dash['A2'] = 'Resultado por fecha de operación · sin IVA · estimación de gestión'
    dash['A5'] = 'Facturación neta'
    dash['A6'] = 'Coste de producto estimado'
    dash['A7'] = 'Gastos netos'
    dash['A8'] = 'Beneficio operativo estimado'
    dash['A9'] = 'Margen operativo estimado'
    dash['B5'] = f'=SUMIFS(Ventas!$H$5:$H${ve},Ventas!$O$5:$O${ve},"OK",Ventas!$A$5:$A${ve},">="&DATE(\'Configuración\'!$B$6,1,1),Ventas!$A$5:$A${ve},"<"&DATE(\'Configuración\'!$B$6+1,1,1))'
    dash['B6'] = f'=SUMIFS(Ventas!$M$5:$M${ve},Ventas!$O$5:$O${ve},"OK",Ventas!$A$5:$A${ve},">="&DATE(\'Configuración\'!$B$6,1,1),Ventas!$A$5:$A${ve},"<"&DATE(\'Configuración\'!$B$6+1,1,1))'
    dash['B7'] = f'=SUMIFS(Gastos!$F$5:$F${ge},Gastos!$K$5:$K${ge},"OK",Gastos!$A$5:$A${ge},">="&DATE(\'Configuración\'!$B$6,1,1),Gastos!$A$5:$A${ge},"<"&DATE(\'Configuración\'!$B$6+1,1,1))'
    dash['B8'] = '=B5-B6-B7'
    dash['B9'] = '=IF(B5=0,"",B8/B5)'
    for r in range(5, 10):
        dash.cell(r, 1).font = Font(name='Aptos', bold=True, color=NAVY)
        dash.cell(r, 2).font = Font(name='Aptos', bold=True, size=13, color=NAVY)
        dash.cell(r, 2).fill = PatternFill('solid', fgColor=GREEN)
        dash.cell(r, 2).number_format = PERCENT if r == 9 else EURO
    for r in range(5, 17):
        dash.cell(r, 3, f'=DATE(\'Configuración\'!$B$6,{r-4},1)').number_format = 'MMMM YYYY'
        for col, sheet, amount in [(4, 'Ventas', 'H'), (5, 'Ventas', 'M'), (6, 'Gastos', 'F')]:
            end = ve if sheet == 'Ventas' else ge
            status_col = 'O' if sheet == 'Ventas' else 'K'
            dash.cell(r, col, f'=SUMIFS({sheet}!${amount}$5:${amount}${end},{sheet}!${status_col}$5:${status_col}${end},"OK",{sheet}!$A$5:$A${end},">="&$C{r},{sheet}!$A$5:$A${end},"<"&EDATE($C{r},1))').number_format = EURO
        dash.cell(r, 7, f'=D{r}-E{r}-F{r}').number_format = EURO
    chart = BarChart()
    chart.title = 'Evolución mensual · sin IVA'
    chart.y_axis.title = '€'
    chart.x_axis.title = 'Mes'
    for col in (4, 7):
        chart.add_data(Reference(dash, min_col=col, min_row=4, max_row=16), titles_from_data=True)
    chart.set_categories(Reference(dash, min_col=3, min_row=5, max_row=16))
    chart.width, chart.height = 20, 10
    dash.add_chart(chart, 'A19')
    dash.freeze_panes = 'C5'
    dash.auto_filter.ref = 'C4:G16'
    dash.protection = SheetProtection(sheet=True, selectLockedCells=False, selectUnlockedCells=False)
    dash['A11'] = 'Filas para revisar'
    dash['B11'] = f'=COUNTIF(Ventas!$O$5:$O${ve},"REVISAR")+COUNTIF(Gastos!$K$5:$K${ge},"REVISAR")+COUNTIF(Clientes!$I$5:$I${ce},"REVISAR")+COUNTIF(Proveedores!$I$5:$I${se},"REVISAR")+COUNTIF(Productos!$M$5:$M${pe},"REVISAR")+COUNTIF(Inventario!$I$5:$I${ie},"REVISAR")+COUNTIF(Tesorería!$H$5:$H${te},"REVISAR")+COUNTIF(Objetivos!$O$5:$O${oe},"REVISAR")'
    dash['A12'] = 'Fuera de ejercicio'
    dash['B12'] = f'=COUNTIF(Ventas!$O$5:$O${ve},"FUERA DE EJERCICIO")+COUNTIF(Gastos!$K$5:$K${ge},"FUERA DE EJERCICIO")+COUNTIF(Tesorería!$H$5:$H${te},"FUERA DE EJERCICIO")'
    for row, label, formula in [
        (14, 'Clientes válidos', f'=COUNTIF(Clientes!$I$5:$I${ce},"OK")'),
        (15, 'Clientes con compras', f'=COUNTIF(Clientes!$K$5:$K${ce},">0")'),
        (16, 'Proveedores válidos', f'=COUNTIF(Proveedores!$I$5:$I${se},"OK")'),
        (17, 'Stock bajo', f'=COUNTIF(Productos!$L$5:$L${pe},"BAJO")'),
        (18, 'Stock negativo', f'=COUNTIF(Productos!$L$5:$L${pe},"NEGATIVO")'),
    ]:
        dash.cell(row, 1, label)
        dash.cell(row, 2, formula)
        dash.cell(row, 2).fill = PatternFill('solid', fgColor=GREEN)
    dash['I4'] = 'Categoría de gasto'
    dash['J4'] = 'Gasto neto'
    for pos in ('I4', 'J4'):
        dash[pos].fill = PatternFill('solid', fgColor=NAVY)
        dash[pos].font = Font(name='Aptos', bold=True, color=WHITE)
    dash.column_dimensions['I'].width = 26
    dash.column_dimensions['J'].width = 20
    for row in range(5, 12):
        dash.cell(row, 9, f"='Configuración'!C{row}")
        dash.cell(row, 10, f'=SUMIFS(Gastos!$F$5:$F${ge},Gastos!$C$5:$C${ge},I{row},Gastos!$K$5:$K${ge},"OK")').number_format = EURO
    dash['I14'] = 'Resultado sin IVA; no representa caja ni resultado fiscal.'
    dash['I15'] = 'Revisa las alertas antes de utilizar los indicadores.'
    dash['I18'] = 'Saldo final de caja'
    dash['J18'] = '=Tesorería!P16'
    dash['J18'].number_format = EURO
    target = f'SUMIFS(Objetivos!$C$5:$C${oe},Objetivos!$A$5:$A${oe},\'Configuración\'!$B$6,Objetivos!$O$5:$O${oe},"OK")'
    dash['I19'] = 'Cumplimiento ventas'
    dash['J19'] = f'=IF(COUNTIFS(Objetivos!$A$5:$A${oe},\'Configuración\'!$B$6,Objetivos!$O$5:$O${oe},"OK",Objetivos!$C$5:$C${oe},"<>")=0,"SIN META",IF({target}=0,"n.a.",B5/{target}))'
    dash['J19'].number_format = PERCENT
    for pos in ('J18', 'J19'):
        dash[pos].fill = PatternFill('solid', fgColor=GREEN)
    for pos in ('B11', 'B12'):
        dash[pos].fill = PatternFill('solid', fgColor=GREEN)
        dash[pos].font = Font(name='Aptos', bold=True, color=NAVY)
    for ws in wb:
        ws.sheet_properties.tabColor = GREEN if ws.title == 'Dashboard' else NAVY

    if demo:
        cfg['B5'] = 'Demo · negocio ficticio'
        cfg['B6'] = 2026
        cfg['I5'], cfg['I6'] = 1000, 100
        products = [('P-001', 'Servicio de diseño', 'Servicios', 400, 90, 'Servicio', None, None),
                    ('P-002', 'Kit de bienvenida', 'Productos', 80, 35, 'Físico', 10, 4)]
        for r, row in enumerate(products, 5):
            for col, value in zip((1, 2, 3, 4, 5, 8, 9, 10), row):
                if value is not None: ps.cell(r, col, value)
        for r, row in enumerate([
            ('C-001', 'Cliente ficticio A', 'demo-a@example.invalid', '', 'Madrid', date(2026, 1, 5), 'Datos ficticios', 'Activo'),
            ('C-002', 'Cliente ficticio B', 'demo-b@example.invalid', '', 'Sevilla', date(2026, 1, 9), 'Datos ficticios', 'Activo')], 5):
            for col, value in enumerate(row, 1): cs.cell(r, col, value)
        for r, row in enumerate([
            ('PR-001', 'Proveedor ficticio A', 'Contacto demo', 'proveedor-a@example.invalid', '', 'Marketing', 'Datos ficticios', 'Activo'),
            ('PR-002', 'Proveedor ficticio B', 'Contacto demo', 'proveedor-b@example.invalid', '', 'Servicios', 'Datos ficticios', 'Activo')], 5):
            for col, value in enumerate(row, 1): ss.cell(r, col, value)
        sales = [(date(2026, 1, 15), 'V-001', 'P-001', 'C-001', 2, 400, 0.10, 0.21, 90, 'F-2026-001'),
                 (date(2026, 2, 3), 'V-002', 'P-002', 'C-002', 3, 80, 0, 0.21, 35, 'F-2026-002')]
        for r, row in enumerate(sales, 5):
            for col, value in zip((1, 2, 3, 4, 5, 6, 7, 9, 12, 14), row): vs.cell(r, col, value)
        expenses = [(date(2026, 1, 20), 'G-001', 'Marketing', 'PR-001', 'Anuncios de prueba', 120, 0.21, 'FR-001'),
                    (date(2026, 2, 9), 'G-002', 'Servicios', 'PR-002', 'Suscripción ficticia', 35, 0.21, 'FR-002')]
        for r, row in enumerate(expenses, 5):
            for col, value in zip((1, 2, 3, 4, 5, 6, 7, 10), row): gs.cell(r, col, value)
        cash_rows = [
            ('T-001', date(2026, 1, 22), 'Cobro', 'Cobro parcial ficticio', 500, 'Banco', 'V-001'),
            ('T-002', date(2026, 2, 12), 'Cobro', 'Cobro total ficticio', 290.40, 'Banco', 'V-002'),
            ('T-003', date(2026, 1, 26), 'Pago', 'Pago parcial ficticio', 100, 'Banco', 'G-001'),
            ('T-004', date(2026, 2, 16), 'Pago', 'Pago total ficticio', 42.35, 'Efectivo', 'G-002'),
        ]
        for r, row in enumerate(cash_rows, 5):
            for col, value in enumerate(row, 1): cash.cell(r, col, value)
        for r, row in enumerate([(2026, 1, 700, 400, 130), (2026, 2, 300, 120, 40)], 5):
            for col, value in zip((1, 2, 3, 7, 11), row): goals.cell(r, col, value)
        for r, row in enumerate([
            (date(2026, 2, 3), 'M-001', 'P-002', 'Salida', 3, 'V-002', 'Salida por venta ficticia'),
            (date(2026, 2, 10), 'M-002', 'P-002', 'Entrada', 2, 'ENT-001', 'Reposición ficticia')], 5):
            for col, value in enumerate(row, 1): inv.cell(r, col, value)
    wb.save(output)
    return output


def main():
    parser = argparse.ArgumentParser(description='Genera NEGOCIO PRO limpio y demo')
    parser.add_argument('--dist', type=Path, default=Path('dist'))
    parser.add_argument('--ventas', type=int, default=500)
    parser.add_argument('--gastos', type=int, default=500)
    parser.add_argument('--productos', type=int, default=200)
    parser.add_argument('--clientes', type=int, default=200)
    parser.add_argument('--proveedores', type=int, default=200)
    parser.add_argument('--movimientos', type=int, default=500)
    parser.add_argument('--caja', type=int, default=500)
    parser.add_argument('--objetivos', type=int, default=120)
    args = parser.parse_args()
    for demo, name in [(False, 'negocio_pro_limpio.xlsx'), (True, 'negocio_pro_demo.xlsx')]:
        print(build(args.dist / name, demo, args.ventas, args.gastos, args.productos,
                    args.clientes, args.proveedores, args.movimientos, args.caja, args.objetivos))


if __name__ == '__main__':
    main()
