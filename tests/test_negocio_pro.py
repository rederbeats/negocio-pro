from decimal import Decimal
from pathlib import Path
import sys

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from negocio_pro import build


@pytest.fixture
def demo(tmp_path):
    return load_workbook(build(tmp_path / 'demo.xlsx', demo=True, ventas=14, gastos=14, productos=3))


def validation(ws, kind):
    return [v for v in ws.data_validations.dataValidation if v.type == kind]


def test_outputs_and_demo(tmp_path, demo):
    clean = load_workbook(build(tmp_path / 'clean.xlsx', ventas=14, gastos=14, productos=3))
    expected = ['Dashboard', 'Ventas', 'Gastos', 'Productos', 'Configuración']
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
