from decimal import Decimal
from pathlib import Path
import sys

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from negocio_pro import build


@pytest.fixture
def demo(tmp_path):
    return load_workbook(build(tmp_path / 'demo.xlsx', demo=True, ventas=14, gastos=14, productos=3, clientes=3, proveedores=3, movimientos=14, caja=14, objetivos=14))


def validation(ws, kind):
    return [v for v in ws.data_validations.dataValidation if v.type == kind]


def test_outputs_and_demo(tmp_path, demo):
    clean = load_workbook(build(tmp_path / 'clean.xlsx', ventas=14, gastos=14, productos=3, clientes=3, proveedores=3, movimientos=14, caja=14, objetivos=14))
    expected = ['Dashboard', 'Ventas', 'Gastos', 'Productos', 'Clientes', 'Proveedores', 'Inventario', 'Tesorería', 'Objetivos', 'Informes', 'Configuración']
    assert demo.sheetnames == clean.sheetnames == expected
    assert demo['Ventas']['B5'].value == 'V-001'
    assert clean['Ventas']['B5'].value is None
    assert len(demo['Dashboard']._charts) == 1
    assert demo.calculation.fullCalcOnLoad


def test_named_ranges_and_dropdowns(demo):
    assert demo.defined_names['CategoriasGasto'].attr_text == "'Configuración'!$C$5:$C$11"
    assert demo.defined_names['TiposIVA'].attr_text == "'Configuración'!$D$5:$D$8"
    assert demo.defined_names['IdsProducto'].attr_text.startswith('Productos!')
    assert {v.formula1 for v in validation(demo['Gastos'], 'list')} == {'=CategoriasGasto', '=TiposIVA'}
    assert {v.formula1 for v in validation(demo['Ventas'], 'list')} == {'=TiposIVA'}
    assert all('!' not in v.formula1 for ws in (demo['Ventas'], demo['Gastos']) for v in validation(ws, 'list'))


def test_protection_and_navigation(demo):
    for sheet, input_cell, output_cell in [('Ventas', 'F5', 'H5'), ('Gastos', 'F5', 'H5'), ('Productos', 'D5', 'F5')]:
        ws = demo[sheet]
        assert ws.protection.sheet and not ws.protection.autoFilter
        assert not ws.protection.selectLockedCells and not ws.protection.selectUnlockedCells
        assert not ws[input_cell].protection.locked and ws[output_cell].protection.locked
        assert ws.auto_filter.ref.startswith('A4:') and ws.freeze_panes == 'A5'
        assert ws[input_cell].fill.fgColor.rgb.endswith('FFF8E5')


def test_invalid_rows_excluded_and_flagged(demo):
    v, g, d = demo['Ventas'], demo['Gastos'], demo['Dashboard']
    assert 'COUNTIF($B$5:$B$18,B5)=1' in v['O5'].value  # ID duplicado
    assert 'COUNTIF(IdsProducto,C5)=1' in v['O5'].value  # producto inexistente
    assert 'COUNTIF(CategoriasGasto,C5)=1' in g['K5'].value  # categoría ausente
    assert 'COUNTIF(TiposIVA,G5)=1' in g['K5'].value
    assert 'O5<>"OK"' in v['H5'].value and 'K5<>"OK"' in g['H5'].value
    assert 'Gastos!$K$5:$K$18,"OK"' in d['B7'].value
    assert '"REVISAR"' in d['B11'].value and '"FUERA DE EJERCICIO"' in d['B12'].value
    assert len(v.conditional_formatting) >= 3 and len(g.conditional_formatting) >= 2


def test_boundaries(demo):
    v, g = demo['Ventas'], demo['Gastos']
    assert any('E5:E18' in str(x.sqref) and x.formula1 == '0.000001' for x in validation(v, 'decimal'))
    assert any('G5:G18' in str(x.sqref) and (x.formula1, x.formula2) == ('0', '1') for x in validation(v, 'decimal'))
    assert any('A5:A18' in str(x.sqref) for x in validation(v, 'date'))
    assert any('A5:A18' in str(x.sqref) for x in validation(g, 'date'))
    for fragment in ('ISNUMBER(G5)', 'G5<=1', 'F5>=0', 'YEAR(A5)', 'ISNUMBER(A5)'):
        assert fragment in v['O5'].value


def test_formula_separation_and_historical_price(demo):
    v, g, d = demo['Ventas'], demo['Gastos'], demo['Dashboard']
    assert 'E5*F5*(1-G5)' in v['H5'].value
    assert 'H5*I5' in v['J5'].value and 'E5*L5' in v['M5'].value
    assert 'F5*G5' in g['H5'].value
    assert d['B8'].value == '=B5-B6-B7'
    assert 'Ventas!$H$5' in d['B5'].value and 'Ventas!$M$5' in d['B6'].value
    assert 'Gastos!$F$5' in d['B7'].value
    assert all('Productos!' not in v[f'{col}5'].value for col in ('H', 'J', 'K', 'M'))
    old_price = v['F5'].value
    demo['Productos']['D5'] = 999
    assert v['F5'].value == old_price


def test_twelve_months(demo):
    d = demo['Dashboard']
    for row, month in zip(range(5, 17), range(1, 13)):
        assert f"'Configuración'!$B$6,{month},1" in d[f'C{row}'].value
        for col in 'DEF':
            assert f'$C{row}' in d[f'{col}{row}'].value and '"OK"' in d[f'{col}{row}'].value
        assert d[f'G{row}'].value == f'=D{row}-E{row}-F{row}'


def test_demo_expected_independent_arithmetic(demo):
    """Oracle Python; no demuestra el cálculo de fórmulas en Excel."""
    v, g = demo['Ventas'], demo['Gastos']
    revenue = sum(Decimal(str(v[f'E{r}'].value)) * Decimal(str(v[f'F{r}'].value)) *
                  (1 - Decimal(str(v[f'G{r}'].value))) for r in (5, 6))
    costs = sum(Decimal(str(v[f'E{r}'].value)) * Decimal(str(v[f'L{r}'].value)) for r in (5, 6))
    expenses = sum(Decimal(str(g[f'F{r}'].value)) for r in (5, 6))
    assert (revenue, costs, expenses) == (Decimal('960'), Decimal('285'), Decimal('155'))
    assert revenue - costs - expenses == Decimal('520')
    assert round((revenue - costs - expenses) / revenue * 100, 2) == Decimal('54.17')


def test_invalid_capacity(tmp_path):
    with pytest.raises(ValueError):
        build(tmp_path / 'x.xlsx', ventas=1)


def test_optional_and_invalid_entity_references(demo):
    v, g, c, s = (demo[x] for x in ('Ventas', 'Gastos', 'Clientes', 'Proveedores'))
    assert 'OR(D5="",COUNTIFS(Clientes!' in v['O5'].value
    assert 'OR(D5="",COUNTIFS(Proveedores!' in g['K5'].value
    assert 'Clientes!$I$5:$I$7,"OK"' in v['O5'].value
    assert 'Proveedores!$I$5:$I$7,"OK"' in g['K5'].value
    assert 'COUNTIF($A$5:$A$7,A5)=1' in c['I5'].value
    assert 'COUNTIF($A$5:$A$7,A5)=1' in s['I5'].value
    assert 'IdsCliente' in demo.defined_names and 'IdsProveedor' in demo.defined_names
    assert any('D5:D18' in str(x.sqref) and 'IdsCliente' in x.formula1 for x in validation(v, 'custom'))
    assert any('D5:D18' in str(x.sqref) and 'IdsProveedor' in x.formula1 for x in validation(g, 'custom'))
    assert 'SUMIFS(Ventas!' in c['J5'].value and 'COUNTIFS(Ventas!' in c['K5'].value
    assert 'SUMIFS(Gastos!' in s['J5'].value
    assert c['C5'].value.endswith('.invalid') and s['D5'].value.endswith('.invalid')


def test_stock_integrity_and_service_exclusion(demo):
    p, m, d = (demo[x] for x in ('Productos', 'Inventario', 'Dashboard'))
    assert p['H5'].value == 'Servicio' and p['H6'].value == 'Físico'
    assert p['I6'].value == 10 and p['J6'].value == 4
    assert m['D5'].value == 'Salida' and m['E5'].value == 3
    assert m['D6'].value == 'Entrada' and m['E6'].value == 2
    assert 10 - 3 + 2 == 9  # stock demo esperado, fuente independiente
    assert 'COUNTIF($B$5:$B$18,B5)=1' in m['I5'].value  # movimientos repetidos
    assert 'COUNTIFS($D$5:$D$18,"Salida",$F$5:$F$18,F5)=1' in m['I5'].value
    assert 'COUNTIF(IdsProducto,C5)=1' in m['I5'].value  # producto inexistente
    assert '"Físico"' in m['I5'].value  # servicios no admiten movimientos
    assert 'I5<>"OK"' in m['H5'].value  # movimiento inválido no altera stock
    assert 'SUMIFS(Inventario!$H$5:$H$18' in p['K6'].value
    assert 'K6<0,"NEGATIVO"' in p['L6'].value
    assert 'K6<J6,"BAJO"' in p['L6'].value
    assert '"NEGATIVO"' in d['B18'].value and '"BAJO"' in d['B17'].value
    assert all('Inventario!' not in d[x].value for x in ('B5', 'B6', 'B7', 'B8'))


def test_new_sheets_protection_and_dashboard(demo):
    for name, editable, calculated in [('Clientes', 'B5', 'I5'), ('Proveedores', 'B5', 'I5'), ('Inventario', 'E5', 'H5')]:
        ws = demo[name]
        assert ws.protection.sheet and not ws.protection.autoFilter
        assert not ws[editable].protection.locked and ws[calculated].protection.locked
        assert ws.freeze_panes == 'A5' and ws.auto_filter.ref.startswith('A4:')
    d = demo['Dashboard']
    assert 'Clientes!' in d['B14'].value and 'Clientes!' in d['B15'].value
    assert 'Proveedores!' in d['B16'].value
    assert 'Productos!' in d['B17'].value and 'Productos!' in d['B18'].value
    assert 'SUMIFS(Gastos!' in d['J5'].value


def test_treasury_partial_payments_and_references(demo):
    t, v, g, d = (demo[x] for x in ('Tesorería', 'Ventas', 'Gastos', 'Dashboard'))
    assert [(t[f'C{r}'].value, t[f'G{r}'].value) for r in range(5, 9)] == [
        ('Cobro', 'V-001'), ('Cobro', 'V-002'), ('Pago', 'G-001'), ('Pago', 'G-002')]
    assert 'COUNTIF($A$5:$A$18,A5)=1' in t['H5'].value  # ID duplicado
    assert 'COUNTIFS(Ventas!$B$5:$B$18,G5' in t['H5'].value  # venta inexistente
    assert 'COUNTIFS(Gastos!$B$5:$B$18,G5' in t['H5'].value  # gasto inexistente
    assert '<=SUMIFS(Ventas!$K$5:$K$18' in t['H5'].value  # no cobrar más que total
    assert '<=SUMIFS(Gastos!$I$5:$I$18' in t['H5'].value  # no pagar más que total
    assert 'SUMIFS($E$5:$E$18,$G$5:$G$18,G5,$C$5:$C$18,C5)' in t['H5'].value
    assert 'Tesorería!$I$5:$I$18' in v['P5'].value
    assert 'Tesorería!$J$5:$J$18' in g['L5'].value
    assert v['Q5'].value == '=IF(K5="","",K5-P5)'
    assert g['M5'].value == '=IF(I5="","",I5-L5)'
    assert all('Tesorería!' not in d[cell].value for cell in ('B5', 'B6', 'B7', 'B8'))
    assert not t['E5'].protection.locked and t['H5'].protection.locked
    assert {x.formula1 for x in validation(t, 'list')} == {'=TiposCaja', '=CuentasCaja'}


def test_treasury_demo_expected_independent_arithmetic(demo):
    t, v, g = (demo[x] for x in ('Tesorería', 'Ventas', 'Gastos'))
    sale_gross = Decimal(str(v['E5'].value)) * Decimal(str(v['F5'].value)) * (
        1 - Decimal(str(v['G5'].value))) * (1 + Decimal(str(v['I5'].value)))
    expense_gross = Decimal(str(g['F5'].value)) * (1 + Decimal(str(g['G5'].value)))
    assert sale_gross == Decimal('871.20')
    assert expense_gross == Decimal('145.20')
    assert sale_gross - Decimal(str(t['E5'].value)) == Decimal('371.20')
    assert expense_gross - Decimal(str(t['E7'].value)) == Decimal('45.20')
    inflows = sum(Decimal(str(t[f'E{r}'].value)) for r in (5, 6))
    outflows = sum(Decimal(str(t[f'E{r}'].value)) for r in (7, 8))
    assert (inflows, outflows) == (Decimal('790.4'), Decimal('142.35'))
    assert Decimal('1100') + inflows - outflows == Decimal('1748.05')
    assert t['P5'].value == '=M5+N5-O5'
    assert t['M6'].value == '=P5'
    assert 'Tesorería!' in demo['Dashboard']['J18'].value


def test_year_change_and_objectives_without_data(demo):
    t, o = demo['Tesorería'], demo['Objetivos']
    assert '"FUERA DE EJERCICIO"' in t['H5'].value
    assert '"FUERA DE EJERCICIO"' in demo['Ventas']['H5'].value
    assert '"FUERA DE EJERCICIO"' in demo['Gastos']['H5'].value
    assert "'Configuración'!$B$6" in t['L5'].value
    assert 'COUNTIFS($A$5:$A$18,A5,$B$5:$B$18,B5)=1' in o['O5'].value
    assert 'COUNT(C5,G5,K5)=0,"SIN META"' in o['O5'].value
    assert '"SIN ACTIVIDAD"' in o['P5'].value
    assert 'P5="SIN ACTIVIDAD"' in o['E5'].value
    assert 'C5=0,"n.a."' in o['E5'].value
    assert o['A5'].value == 2026 and o['B5'].value == 1
    assert o['C5'].value == 700 and o['G5'].value == 400 and o['K5'].value == 130


def test_reports_reconcile_with_dashboard(demo):
    d, r = demo['Dashboard'], demo['Informes']
    for row in range(5, 17):
        assert r[f'A{row}'].value == f'=Dashboard!C{row}'
        for col, source in [('B', 'D'), ('C', 'E'), ('D', 'F'), ('E', 'G')]:
            assert r[f'{col}{row}'].value == f'=Dashboard!{source}{row}'
    for col, source in [('B', 'B5'), ('C', 'B6'), ('D', 'B7'), ('E', 'B8'), ('F', 'B9')]:
        assert r[f'{col}18'].value == f'=Dashboard!{source}'
    assert r['I5'].value == '=Dashboard!J5'
    assert 'SUMIFS(Ventas!$H$5:$H$18' in r['M5'].value
    assert 'SUMIFS(Ventas!$M$5:$M$18' in r['N5'].value
    assert len(r._charts) == 1
