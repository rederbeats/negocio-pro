"""Auditoría de distribución. LibreOffice se usa solo cuando está instalado."""
from datetime import date
from pathlib import Path
import shutil
import subprocess
import sys

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from negocio_pro import build


def engine():
    for candidate in (Path(r'C:\Program Files\LibreOffice\program\soffice.com'),
                      Path(r'C:\Program Files\LibreOffice\program\soffice.exe')):
        if candidate.exists():
            return str(candidate)
    return shutil.which('libreoffice') or shutil.which('soffice')


def small_book(path, demo=True):
    return build(path, demo=demo, ventas=8, gastos=8, productos=4, clientes=4,
                 proveedores=4, movimientos=8, caja=8, objetivos=8)


def recalculate(path, tmp_path):
    exe = engine()
    if not exe:
        pytest.skip('No hay motor LibreOffice disponible')
    outdir = tmp_path / 'recalc'
    outdir.mkdir()
    profile = (tmp_path / 'lo_profile').resolve().as_uri()
    result = subprocess.run([exe, f'-env:UserInstallation={profile}', '--headless',
                             '--convert-to', 'xlsx', '--outdir', str(outdir), str(path)],
                            capture_output=True, text=True, timeout=120)
    output = outdir / path.name
    assert result.returncode == 0, result.stdout + result.stderr
    assert output.exists(), result.stdout + result.stderr
    return output


def assert_no_formula_errors(book):
    errors = [(ws.title, cell.coordinate, cell.value) for ws in book
              for row in ws for cell in row if cell.data_type == 'e']
    assert errors == []


def test_clean_has_no_demo_records_or_personal_examples(tmp_path):
    book = load_workbook(small_book(tmp_path / 'clean.xlsx', demo=False))
    assert book.sheetnames[0] == 'Instrucciones'
    for name, cols in [('Ventas', 17), ('Gastos', 13), ('Productos', 13),
                       ('Clientes', 11), ('Proveedores', 10), ('Inventario', 9),
                       ('Tesorería', 10), ('Objetivos', 16)]:
        ws = book[name]
        editable = [c.value for row in ws.iter_rows(min_row=5, max_row=12, max_col=cols)
                    for c in row if not c.protection.locked]
        assert all(value is None for value in editable), (name, editable)
    assert book['Configuración']['B5'].value == 'Mi negocio'
    assert 'demo' not in book['Instrucciones']['B6'].value.lower()


def test_formula_tokens_and_links_are_intact(tmp_path):
    book = load_workbook(small_book(tmp_path / 'static.xlsx'))
    for ws in book:
        for row in ws:
            for cell in row:
                if cell.data_type == 'f':
                    assert all(bad not in cell.value for bad in ('#REF!', '#DIV/0!', '#VALUE!'))
                if cell.hyperlink and cell.hyperlink.location:
                    target = cell.hyperlink.location.split('!')[0].strip("'")
                    assert target in book.sheetnames
    assert book.properties.creator == 'NEXORA'
    assert book.properties.title == 'NEGOCIO PRO'


def test_engine_demo_values_and_preservation(tmp_path):
    original = load_workbook(small_book(tmp_path / 'demo.xlsx'))
    calculated_path = recalculate(tmp_path / 'demo.xlsx', tmp_path)
    calculated = load_workbook(calculated_path, data_only=True)
    roundtrip = load_workbook(calculated_path)
    assert_no_formula_errors(calculated)
    dash = calculated['Dashboard']
    assert [dash[c].value for c in ('B5', 'B6', 'B7', 'B8')] == [960, 285, 155, 520]
    assert dash['B9'].value == pytest.approx(520 / 960)
    assert dash['J18'].value == pytest.approx(1748.05)
    assert dash['J19'].value == pytest.approx(.96)
    assert calculated['Ventas']['Q5'].value == pytest.approx(371.2)
    assert calculated['Gastos']['M5'].value == pytest.approx(45.2)
    assert calculated['Productos']['K6'].value == 9
    assert calculated['Tesorería']['P16'].value == pytest.approx(1748.05)
    assert calculated['Informes']['E18'].value == 520
    assert roundtrip.sheetnames == original.sheetnames
    assert roundtrip.active.title == 'Instrucciones'
    assert set(roundtrip.defined_names) == set(original.defined_names)
    for name in original.sheetnames:
        before, after = original[name], roundtrip[name]
        assert len(before.data_validations.dataValidation) == len(after.data_validations.dataValidation)
        assert len(before.conditional_formatting) == len(after.conditional_formatting)
        assert len(before._charts) == len(after._charts)
        assert before.protection.sheet == after.protection.sheet
        assert before.auto_filter.ref == after.auto_filter.ref
        assert before.freeze_panes == after.freeze_panes
    for name, editable, formula in [('Ventas', 'F5', 'H5'), ('Gastos', 'F5', 'H5'),
                                    ('Tesorería', 'E5', 'H5'), ('Objetivos', 'C5', 'D5')]:
        assert not roundtrip[name][editable].protection.locked
        assert roundtrip[name][formula].protection.locked


def test_engine_year_change_and_excess_cash(tmp_path):
    path = small_book(tmp_path / 'scenario.xlsx')
    book = load_workbook(path)
    book['Configuración']['B6'] = 2027
    t = book['Tesorería']
    for col, value in enumerate(('T-005', date(2027, 1, 5), 'Cobro', 'Exceso', 400,
                                 'Banco', 'V-001'), 1):
        t.cell(9, col, value)
    book.save(path)
    calculated = load_workbook(recalculate(path, tmp_path), data_only=True)
    assert_no_formula_errors(calculated)
    assert calculated['Ventas']['O5'].value == 'FUERA DE EJERCICIO'
    assert calculated['Ventas']['K5'].value == pytest.approx(871.2)
    assert calculated['Tesorería']['H5'].value == 'FUERA DE EJERCICIO'
    assert calculated['Tesorería']['H9'].value == 'REVISAR'
    assert calculated['Ventas']['P5'].value == 500  # cobro previo permanece válido
    assert calculated['Ventas']['Q5'].value == pytest.approx(371.2)
    assert calculated['Dashboard']['J18'].value == 'REVISAR SALDO'
    assert calculated['Tesorería']['P16'].value in (None, '')
    assert calculated['Dashboard']['B8'].value == 0
    assert calculated['Objetivos']['D5'].value == 720  # objetivo 2026 sigue mostrando real


def test_engine_valid_partial_collection_next_year(tmp_path):
    path = small_book(tmp_path / 'next_year.xlsx')
    book = load_workbook(path)
    cfg = book['Configuración']
    cfg['B6'] = 2027
    cfg['B8'] = 2027
    cfg['I5'] = 1690.40  # cierre de Banco 2026 en la demo
    cfg['I6'] = 57.65    # cierre de Efectivo 2026 en la demo
    t = book['Tesorería']
    for col, value in enumerate(('T-005', date(2027, 1, 5), 'Cobro', 'Resto de V-001',
                                 371.2, 'Banco', 'V-001'), 1):
        t.cell(9, col, value)
    book.save(path)
    calculated = load_workbook(recalculate(path, tmp_path), data_only=True)
    assert_no_formula_errors(calculated)
    assert calculated['Tesorería']['H9'].value == 'OK'
    assert calculated['Tesorería']['I9'].value == pytest.approx(371.2)
    assert calculated['Ventas']['P5'].value == pytest.approx(871.2)
    assert calculated['Ventas']['Q5'].value == 0
    assert calculated['Dashboard']['B8'].value == 0  # sin nuevas operaciones de gestión
    assert calculated['Tesorería']['N5'].value == pytest.approx(371.2)
    assert calculated['Tesorería']['P16'].value == pytest.approx(2119.25)
