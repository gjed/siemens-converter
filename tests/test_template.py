import openpyxl
from importlib.resources import files


def _load_template():
    template_path = files("siemens_converter").joinpath("template.xlsx")
    return openpyxl.load_workbook(template_path)


def _load_static_template():
    template_path = files("siemens_converter").joinpath("static_data_template.xlsx")
    return openpyxl.load_workbook(template_path)


def test_template_has_three_sheets():
    wb = _load_template()
    assert len(wb.sheetnames) == 3


def test_template_ripartizione_sheet_exists():
    wb = _load_template()
    # First sheet name should contain "Ripartizione" or similar
    assert any("Ripartizione" in name or "Riparto" in name for name in wb.sheetnames)


def test_template_central_meter_labels_exist():
    wb = _load_template()
    ws = wb.worksheets[0]
    # Row 27 has the riscaldamento initial reading label
    assert ws["A27"].value is not None
    # Row 30 has the ACS initial reading label
    assert ws["A30"].value is not None


def test_template_data_cells_are_cleared():
    wb = _load_template()
    ws = wb.worksheets[0]
    # These should be None (cleared for writer to populate)
    assert ws["C28"].value is None  # Riscaldamento finale
    assert ws["C31"].value is None  # ACS finale
    assert ws["C78"].value is None  # App 01 heat
    assert ws["C130"].value is None  # App 01 water
    assert ws["C182"].value is None  # App 01 AFS


def test_template_formulas_preserved():
    wb = _load_template()
    ws = wb.worksheets[0]
    # Consumption formula rows should still have formulas
    # D78 = =+C78-B78 (consumption = finale - iniziale)
    d78 = ws["D78"].value
    assert d78 is not None and "=" in str(d78)


def test_template_tabelle_millesimali_structure():
    wb = _load_template()
    ws = wb["Tabelle millesimali"]
    assert ws["A1"].value == "Condomini"
    assert ws["C2"].value == "kWh"


def test_template_tabelle_millesimali_data_cleared():
    wb = _load_template()
    ws = wb["Tabelle millesimali"]
    for r in range(4, 14):
        # Names are formulas referencing Inquilini
        assert ws.cell(row=r, column=1).value is not None
        assert "Inquilini" in str(ws.cell(row=r, column=1).value)
        # Subalterno and energy values should be cleared
        assert ws.cell(row=r, column=2).value is None  # subalterno
        assert ws.cell(row=r, column=3).value is None  # energia risc.
        assert ws.cell(row=r, column=5).value is None  # energia ACS
    # Formulas in D and F should still be there
    assert ws["D4"].value is not None
    assert ws["F4"].value is not None


def test_template_apartment_cells_reference_inquilini():
    wb = _load_template()
    ws = wb.worksheets[0]
    # Summary: owner only (column A)
    assert ws["A3"].value == "=Inquilini!A2"  # apt 1 owner
    # Heat section 78-95: first row = apt 1 tenant (B), then pairs
    assert ws["A78"].value == "=Inquilini!B2"  # apt 1 tenant
    assert ws["A79"].value == "=Inquilini!A3"  # apt 2 owner
    assert ws["A80"].value == "=Inquilini!B3"  # apt 2 tenant
    # Detail sections: owner (A) then tenant (B) per apartment
    assert ws["A102"].value == "=Inquilini!A2"  # apt 1 owner
    assert ws["A103"].value == "=Inquilini!B2"  # apt 1 tenant
    assert ws["A129"].value == "=Inquilini!A2"  # water apt 1 owner
    assert ws["A130"].value == "=Inquilini!B2"  # water apt 1 tenant


def test_template_tabella_2026_eur_cleared():
    wb = _load_template()
    ws = wb["Tabella_2026"]
    # EUR cost values should be cleared
    for r in range(3, 9):
        assert ws.cell(row=r, column=5).value is None
    # Meter readings should be cleared
    for r in range(20, 23):
        assert ws.cell(row=r, column=6).value is None  # initial
        assert ws.cell(row=r, column=7).value is None  # final
    # Date should be stub
    assert ws["E18"].value is None
    assert ws["F18"].value is None


# -- Static data template tests --


def test_static_template_has_five_sheets():
    wb = _load_static_template()
    assert wb.sheetnames == [
        "Inquilini",
        "Millesimali",
        "Costi",
        "Contatori",
        "Letture_precedenti",
    ]


def test_static_template_inquilini_has_title_and_headers():
    wb = _load_static_template()
    ws = wb["Inquilini"]
    # Row 1: section title
    assert ws.cell(row=1, column=1).value is not None
    # Row 2: column headers
    headers = [ws.cell(row=2, column=c).value for c in range(1, 4)]
    assert any("App" in str(h) for h in headers if h)
    assert any("Proprietario" in str(h) or "Condomino" in str(h) for h in headers if h)
    assert any("Inquilino" in str(h) or "Conduttore" in str(h) for h in headers if h)


def test_static_template_inquilini_has_ten_data_rows():
    wb = _load_static_template()
    ws = wb["Inquilini"]
    # Data rows start at row 3 (title + header above)
    apt_nums = [ws.cell(row=r, column=1).value for r in range(3, 13)]
    assert apt_nums == list(range(1, 11))


def test_static_template_millesimali_has_title_and_headers():
    wb = _load_static_template()
    ws = wb["Millesimali"]
    assert ws.cell(row=1, column=1).value is not None  # title
    # Row 2 has column headers
    headers = [ws.cell(row=2, column=c).value for c in range(1, 7)]
    assert any(h for h in headers if h)


def test_static_template_millesimali_has_totale_and_data():
    wb = _load_static_template()
    ws = wb["Millesimali"]
    # Row 4 = TOTALE, rows 5-14 = apartments 1-10
    totale = ws.cell(row=4, column=1).value
    assert totale is not None
    apt_nums = [ws.cell(row=r, column=1).value for r in range(5, 15)]
    assert apt_nums == list(range(1, 11))


def test_static_template_millesimali_has_computed_columns():
    wb = _load_static_template()
    ws = wb["Millesimali"]
    # Column 4 (Mill risc.) and 6 (Mill ACS) have formulas
    d5 = ws.cell(row=5, column=4).value
    f5 = ws.cell(row=5, column=6).value
    assert d5 is not None and "=" in str(d5)
    assert f5 is not None and "=" in str(f5)


def test_static_template_costi_has_required_labels():
    wb = _load_static_template()
    ws = wb["Costi"]
    # Labels are in the data section (after title + header), check all cells col 1
    all_labels = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    assert "Energia elettrica" in all_labels
    assert "Gas metano" in all_labels
    assert "Acqua condominio" in all_labels
    assert "Conduzione e manutenzione" in all_labels
    assert "Contabilizzazione" in all_labels
    assert "Acqua sanitaria manutenzione" in all_labels


def test_static_template_costi_has_totale_formula():
    wb = _load_static_template()
    ws = wb["Costi"]
    # There should be a TOTALE row with a SUM formula
    total_row = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and "TOTALE" in str(v).upper():
            total_row = r
            break
    assert total_row is not None
    total_formula = ws.cell(row=total_row, column=2).value
    assert total_formula is not None and "SUM" in str(total_formula)


def test_static_template_contatori_has_required_meters():
    wb = _load_static_template()
    ws = wb["Contatori"]
    # Meter names are in column 1; check case-insensitively
    all_vals = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    all_upper = [str(v).upper() for v in all_vals if v]
    assert "ENERGIA ELETTRICA CT" in all_upper
    assert "GAS METANO CT" in all_upper
    assert "ACQUA GENERALE" in all_upper


def test_static_template_contatori_has_consumo_formula():
    wb = _load_static_template()
    ws = wb["Contatori"]
    # Column 5 should have consumption formulas
    formulas = [ws.cell(row=r, column=5).value for r in range(3, 6)]
    assert any(v and "=" in str(v) for v in formulas)


def test_static_template_letture_precedenti_has_ten_data_rows():
    wb = _load_static_template()
    ws = wb["Letture_precedenti"]
    # Find data rows (integer apartment numbers)
    apt_nums = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, int):
            apt_nums.append(v)
    assert apt_nums == list(range(1, 11))
