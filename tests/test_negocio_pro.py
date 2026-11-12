from pathlib import Path
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from negocio_pro import build


def test_structure_and_demo():
    with tempfile.TemporaryDirectory() as tmp:
        path = build(Path(tmp) / 'demo.xlsx', demo=True, ventas=3, gastos=3, productos=3)
        w = load_workbook(path)
        assert w.sheetnames == ['Dashboard', 'Ventas', 'Gastos', 'Productos', 'Configuración']
        assert w['Ventas']['B5'].value == 'V-001'
        assert w['Ventas']['H5'].value.startswith('=IF(')
        assert 'SUMIFS' in w['Dashboard']['B5'].value
        assert 'Ventas!' in w['Dashboard']['D5'].value
        assert len(w['Dashboard']._charts) == 1
        assert w['Ventas'].protection.sheet
        assert w['Ventas']['H5'].protection.locked
        assert not w['Ventas']['F5'].protection.locked
        assert len(w['Ventas'].data_validations.dataValidation) >= 5
        assert w['Configuración']['B5'].value.startswith('Demo')


def test_clean_and_capacity():
    with tempfile.TemporaryDirectory() as tmp:
        w = load_workbook(build(Path(tmp) / 'clean.xlsx', ventas=2, gastos=2, productos=2))
        assert w['Ventas']['B5'].value is None
        assert w['Productos']['A5'].value is None
        assert w['Ventas']['H6'].value.startswith('=IF(')
        assert w['Gastos']['I6'].value.startswith('=IF(')


def test_invalid_capacity():
    with tempfile.TemporaryDirectory() as tmp:
        try:
            build(Path(tmp) / 'x.xlsx', ventas=1)
        except ValueError:
            pass
        else:
            raise AssertionError('Debe rechazar capacidad menor que 2')


def test_financial_formula_contract():
    with tempfile.TemporaryDirectory() as tmp:
        w = load_workbook(build(Path(tmp) / 'x.xlsx', ventas=2, gastos=2, productos=2))
        v, g, d = w['Ventas'], w['Gastos'], w['Dashboard']
        assert 'E5*F5*(1-G5)' in v['H5'].value
        assert 'H5*I5' in v['J5'].value
        assert 'E5*L5' in v['M5'].value
        assert 'F5*G5' in g['H5'].value
        assert d['B8'].value == '=B5-B6-B7'
        assert 'Ventas!$H$5' in d['B5'].value
        assert 'Gastos!$F$5' in d['B7'].value
