import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_report():
    doc = docx.Document()

    # ── 1. PAGE SETUP (A4 with Standard Academic Margins) ──
    sec1 = doc.sections[0]
    sec1.page_width = Inches(8.27)   # A4 Width (21.0 cm)
    sec1.page_height = Inches(11.69) # A4 Height (29.7 cm)
    sec1.top_margin = Inches(0.9)
    sec1.bottom_margin = Inches(0.9)
    sec1.left_margin = Inches(0.9)
    sec1.right_margin = Inches(0.85)

    # Base style font
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)

    def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = tcPr.first_child_found_in('w:tcBorders')
        if tcBorders is None:
            tcBorders = OxmlElement('w:tcBorders')
            tcPr.append(tcBorders)
        
        borders = {'top': top, 'bottom': bottom, 'left': left, 'right': right}
        for edge, data in borders.items():
            tag = f'w:{edge}'
            el = tcBorders.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                tcBorders.append(el)
            if data:
                el.set(qn('w:val'), data.get('val', 'single'))
                el.set(qn('w:sz'), str(data.get('sz', 4)))
                el.set(qn('w:color'), data.get('color', '000000'))
                el.set(qn('w:space'), '0')
            else:
                el.set(qn('w:val'), 'none')

    def set_cell_shading(cell, color_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=40, bottom=40, left=60, right=60):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            node = OxmlElement(f'w:{m}')
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    # ── HEADER ATAS (JURNAL COMPUTING & e-ISSN) ──
    header_tbl = doc.add_table(rows=1, cols=2)
    header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_tbl.autofit = False
    
    cell_hl = header_tbl.cell(0, 0)
    cell_hl.width = Inches(3.5)
    p_hl = cell_hl.paragraphs[0]
    p_hl.paragraph_format.space_after = Pt(2)
    r_hl = p_hl.add_run("JURNAL COMPUTING")
    r_hl.font.name = 'Arial'
    r_hl.font.size = Pt(8.5)
    r_hl.font.bold = True
    r_hl.font.color.rgb = RGBColor(16, 149, 106) # Teal / Green

    cell_hr = header_tbl.cell(0, 1)
    cell_hr.width = Inches(3.0)
    p_hr = cell_hr.paragraphs[0]
    p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hr.paragraph_format.space_after = Pt(2)
    r_hr = p_hr.add_run("e-ISSN: 2722 - 2888")
    r_hr.font.name = 'Arial'
    r_hr.font.size = Pt(8)
    r_hr.font.color.rgb = RGBColor(100, 116, 139)

    set_cell_borders(cell_hl, bottom={'val': 'single', 'sz': 6, 'color': 'CBD5E1'})
    set_cell_borders(cell_hr, bottom={'val': 'single', 'sz': 6, 'color': 'CBD5E1'})

    # ── JUDUL ARTIKEL (Maksimum 14 kata, Arial 12 Bold KAPITAL) ──
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(6)
    run_title = p_title.add_run(
        "IMPLEMENTASI REGRESI LINEAR SEDERHANA UNTUK MEMPREDIKSI KALORI TERBAKAR BERDASARKAN DURASI OLAHRAGA MENGGUNAKAN FLASK"
    )
    run_title.font.name = 'Arial'
    run_title.font.size = Pt(12)
    run_title.bold = True

    # ── PENULIS (Nama lengkap, tanpa gelar, Arial 10 Bold) ──
    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_author.paragraph_format.space_after = Pt(1)
    run_auth = p_author.add_run("Asep Sunandar")
    run_auth.font.name = 'Arial'
    run_auth.font.size = Pt(10)
    run_auth.bold = True

    p_affil = doc.add_paragraph()
    p_affil.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_affil.paragraph_format.space_after = Pt(12)
    run_affil = p_affil.add_run("Program Studi Teknik Informatika, Fakultas Ilmu Komputer\nEmail: asepsunandar@student.kampus.ac.id")
    run_affil.font.name = 'Arial'
    run_affil.font.size = Pt(8.5)
    run_affil.italic = True

    # ── ABSTRAK (BAHASA INDONESIA) ──
    p_abs_id = doc.add_paragraph()
    p_abs_id.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_abs_id.paragraph_format.line_spacing = 1.0
    p_abs_id.paragraph_format.space_after = Pt(3)
    
    r_abs_lbl = p_abs_id.add_run("ABSTRAK: ")
    r_abs_lbl.font.name = 'Arial'
    r_abs_lbl.font.size = Pt(9)
    r_abs_lbl.bold = True

    r_abs_txt = p_abs_id.add_run(
        "Aktivitas fisik secara teratur merupakan faktor penting dalam menjaga kesehatan tubuh, di mana estimasi energi "
        "yang dikeluarkan dalam bentuk kalori terbakar menjadi indikator utama evaluasi kebugaran. Namun, keterbatasan akses "
        "terhadap sensor fisiologis canggih sering kali menjadi hambatan dalam memperoleh perkiraan pengeluaran energi secara cepat. "
        "Penelitian ini bertujuan untuk mengimplementasikan algoritma Regresi Linear Sederhana (Simple Linear Regression) guna "
        "memprediksi jumlah kalori yang terbakar berdasarkan durasi olahraga ke dalam aplikasi web berbasis Python Flask. Dataset "
        "publik Calories Burnt Prediction berukuran 15.000 baris data diproses melalui tahapan pembersihan data dan pembagian train-test "
        "split 80:20 dengan random state 42. Evaluasi model dilakukan murni pada data pengujian (testing set) yang menghasilkan koefisien "
        "determinasi (R²) sebesar 0,9145, Mean Absolute Error (MAE) sebesar 13,6135 kcal, Mean Squared Error (MSE) sebesar 345,0410 kcal², "
        "dan Root Mean Squared Error (RMSE) sebesar 18,5753 kcal, dengan persamaan matematis Calories = -21,6668 + 7,1534 × Duration. Model "
        "diintegrasikan ke dalam arsitektur web interaktif dengan sistem validasi input dinamis berbasis rentang dataset aktual (1–30 menit) "
        "guna mencegah ekstrapolasi tak terkontrol. Hasil penelitian membuktikan bahwa durasi olahraga memiliki korelasi linear yang kuat "
        "dalam menjelaskan variasi pengeluaran kalori, serta sistem web yang dibangun dapat menjadi media edukasi dan kalkulator estimasi yang praktis dan responsif."
    )
    r_abs_txt.font.name = 'Arial'
    r_abs_txt.font.size = Pt(9)

    p_kw_id = doc.add_paragraph()
    p_kw_id.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_kw_id.paragraph_format.space_after = Pt(8)
    r_kw_lbl = p_kw_id.add_run("Kata kunci: ")
    r_kw_lbl.font.name = 'Arial'
    r_kw_lbl.font.size = Pt(9)
    r_kw_lbl.bold = True
    r_kw_txt = p_kw_id.add_run("Durasi Olahraga, Flask, Kalori Terbakar, Machine Learning, Regresi Linear Sederhana.")
    r_kw_txt.font.name = 'Arial'
    r_kw_txt.font.size = Pt(9)

    # ── ABSTRACT (ENGLISH) ──
    p_abs_en = doc.add_paragraph()
    p_abs_en.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_abs_en.paragraph_format.line_spacing = 1.0
    p_abs_en.paragraph_format.space_after = Pt(3)
    
    r_en_lbl = p_abs_en.add_run("ABSTRACT: ")
    r_en_lbl.font.name = 'Arial'
    r_en_lbl.font.size = Pt(9)
    r_en_lbl.bold = True
    r_en_lbl.italic = True

    r_en_txt = p_abs_en.add_run(
        "Regular physical activity is an essential factor in maintaining bodily health, where estimating energy expenditure in the form "
        "of burned calories serves as a primary fitness evaluation indicator. However, limited access to advanced physiological sensors "
        "often impedes the rapid acquisition of energy expenditure estimations. This study aims to implement the Simple Linear Regression "
        "algorithm to predict calories burned based on exercise duration within a lightweight Python Flask web application. A public Calories "
        "Burnt Prediction dataset comprising 15,000 records was preprocessed through cleaning and an 80:20 train-test split using random state 42. "
        "Model evaluation conducted strictly on the testing data yielded a coefficient of determination (R²) of 0.9145, Mean Absolute Error (MAE) "
        "of 13.6135 kcal, Mean Squared Error (MSE) of 345.0410 kcal², and Root Mean Squared Error (RMSE) of 18.5753 kcal, establishing the mathematical "
        "model Calories = -21.6668 + 7.1534 × Duration. The trained model was deployed into an interactive web architecture equipped with dynamic "
        "range-based validation (1–30 minutes) to eliminate extrapolation hazards. The findings demonstrate that exercise duration has a strong "
        "linear relationship in explaining caloric burn variance, and the resulting web system provides an intuitive, reliable, and responsive educational estimation platform."
    )
    r_en_txt.font.name = 'Arial'
    r_en_txt.font.size = Pt(9)
    r_en_txt.italic = True

    p_kw_en = doc.add_paragraph()
    p_kw_en.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_kw_en.paragraph_format.space_after = Pt(10)
    r_kwen_lbl = p_kw_en.add_run("Keywords: ")
    r_kwen_lbl.font.name = 'Arial'
    r_kwen_lbl.font.size = Pt(9)
    r_kwen_lbl.bold = True
    r_kwen_lbl.italic = True
    r_kwen_txt = p_kw_en.add_run("Calories Burned, Exercise Duration, Flask, Machine Learning, Simple Linear Regression.")
    r_kwen_txt.font.name = 'Arial'
    r_kwen_txt.font.size = Pt(9)
    r_kwen_txt.italic = True

    # ── DIVIDER LINE SEBELUM SECTION 2 ──
    div_tbl = doc.add_table(rows=1, cols=1)
    div_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_d = div_tbl.cell(0, 0)
    cell_d.width = Inches(6.5)
    set_cell_borders(cell_d, bottom={'val': 'single', 'sz': 8, 'color': '64748B'})
    p_div = cell_d.paragraphs[0]
    p_div.paragraph_format.space_before = Pt(0)
    p_div.paragraph_format.space_after = Pt(0)

    # ══════════════════════════════════════════════════════════════
    # ── 2. SECTION 2: CONTINUOUS BREAK DENGAN 2 KOLOM ──
    # ══════════════════════════════════════════════════════════════
    sec2 = doc.add_section(WD_SECTION.CONTINUOUS)
    sec2.top_margin = Inches(0.8)
    sec2.bottom_margin = Inches(0.9)
    sec2.left_margin = Inches(0.9)
    sec2.right_margin = Inches(0.85)

    # Format 2 kolom (gap 0.35 in = 504 dxa)
    sectPr2 = sec2._sectPr
    cols = sectPr2.xpath('./w:cols')
    if cols:
        cols[0].set(qn('w:num'), '2')
        cols[0].set(qn('w:space'), '504')
    else:
        cols_el = OxmlElement('w:cols')
        cols_el.set(qn('w:num'), '2')
        cols_el.set(qn('w:space'), '504')
        sectPr2.append(cols_el)

    # Helper Styles dalam 2 Kolom
    def add_sec2_heading(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.font.name = 'Arial'
        run.font.size = Pt(10.5)
        run.bold = True
        return p

    def add_sec2_sub(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.font.name = 'Arial'
        run.font.size = Pt(9.5)
        run.bold = True
        run.italic = True
        return p

    def add_sec2_sub_sub(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.font.name = 'Arial'
        run.font.size = Pt(9.5)
        run.bold = False
        run.italic = True
        return p

    def add_sec2_body(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(9.5)
        return p

    def add_sec2_caption_tbl(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(8.5)
        return p

    def add_sec2_caption_fig(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(8.5)
        return p

    def add_sec2_screenshot_box(fig_no, fig_title, guide_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(3.0)
        set_cell_borders(cell, 
                         top={'val': 'single', 'sz': 5, 'color': '2563EB'},
                         bottom={'val': 'single', 'sz': 5, 'color': '2563EB'},
                         left={'val': 'single', 'sz': 5, 'color': '2563EB'},
                         right={'val': 'single', 'sz': 5, 'color': '2563EB'})
        set_cell_shading(cell, "F8FAFC")
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
        
        cp = cell.paragraphs[0]
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(6)
        cp.paragraph_format.space_after = Pt(2)
        r1 = cp.add_run(f"📷 [SCREENSHOT GAMBAR {fig_no}]")
        r1.font.name = 'Arial'
        r1.font.size = Pt(8.5)
        r1.bold = True
        r1.font.color.rgb = RGBColor(29, 78, 216)

        cp2 = cell.add_paragraph()
        cp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp2.paragraph_format.space_after = Pt(6)
        r2 = cp2.add_run(guide_text)
        r2.font.name = 'Arial'
        r2.font.size = Pt(7.5)
        r2.italic = True
        r2.font.color.rgb = RGBColor(71, 85, 105)

        add_sec2_caption_fig(f"Gambar {fig_no}  {fig_title}")

    # ── PENDAHULUAN (2 KOLOM) ──
    add_sec2_heading("PENDAHULUAN")
    add_sec2_body(
        "Pemeliharaan kebugaran jasmani melalui latihan fisik terprogram semakin mendapatkan perhatian luas di kalangan "
        "masyarakat modern. Salah satu parameter kuantitatif yang kerap digunakan untuk mengukur efektivitas dan intensitas suatu "
        "sesi latihan adalah jumlah energi metabolisme yang dikeluarkan, yang umumnya dinyatakan dalam satuan kilokalori (kcal). "
        "Pengukuran kalori secara akurat di laboratorium memerlukan peralatan kalorimetri langsung atau tidak langsung (indirect "
        "calorimetry) yang mahal, kompleks, dan tidak praktis untuk penggunaan sehari-hari (Ainsworth et al., 2011: 43-45). Oleh sebab "
        "itu, pendekatan komputasi berbasis data (data-driven predictive modeling) menggunakan cabang pembelajaran mesin (machine "
        "learning) menjadi alternatif yang sangat relevan untuk menyediakan estimasi secara cepat, murah, dan dapat diakses oleh publik luas."
    )
    add_sec2_body(
        "Dalam konteks pemodelan statistika parametrik, Regresi Linear Sederhana (Simple Linear Regression) merupakan metode "
        "fundamental yang mempelajari relasi fungsional linier antara satu variabel penjelas (independent variable) dengan satu "
        "variabel respon (dependent variable). Variabel durasi aktivitas fisik (menit) secara teoritis memiliki korelasi positif "
        "yang sangat dominan terhadap total energi yang dikonsumsi otot rangka selama kontraksi berlangsung (Powers & Howley, 2018: 88-92). "
        "Kendati demikian, banyak aplikasi kesehatan komersial memposisikan model estimasi sebagai sistem tertutup (black box) tanpa "
        "memaparkan formulasi matematis, batas rentang validitas data (domain of applicability), serta evaluasi galat prediksi yang transparan kepada pengguna."
    )
    add_sec2_body(
        "Penelitian ini memfokuskan pada pengembangan sistem terintegrasi yang tidak hanya melatih model Regresi Linear Sederhana "
        "dari data empiris, tetapi juga menyajikannya dalam bentuk aplikasi web interaktif menggunakan kerangka kerja Python Flask. "
        "Rumusan masalah dalam penelitian ini adalah bagaimana membangun model regresi linear antara durasi olahraga dan kalori terbakar, "
        "bagaimana membuktikan ketepatan matematis model melalui pembuktian manual dan komputasional, serta bagaimana mengimplementasikan "
        "model tersebut ke dalam antarmuka web yang dilengkapi sistem validasi rentang dan visualisasi performa. Tujuan penelitian ini "
        "adalah menganalisis kinerja regresi linear sederhana pada dataset pengujian (testing dataset), menyajikan pembuktian perhitungan "
        "statistika manual, serta membangun platform web yang aman dari ekstrapolasi untuk keperluan edukasi dan kalkulasi estimasi kalori berbasis web."
    )

    # ── METODA (2 KOLOM) ──
    add_sec2_heading("METODA")
    add_sec2_body(
        "Penelitian ini dilaksanakan melalui lima tahapan terstruktur: pengumpulan dataset, pra-pemrosesan data (data preprocessing), "
        "pemisahan data (train-test split), pelatihan dan perhitungan manual algoritma regresi, serta integrasi sistem web berbasis Flask. "
        "Waktu pelaksanaan penelitian berlangsung pada bulan Oktober 2026."
    )
    add_sec2_body(
        "Data yang digunakan bersumber dari dataset publik Calories Burnt Prediction Dataset (Kaggle) yang terdiri atas file exercise.csv "
        "(berisi identitas pengguna dan atribut latihan) serta calories.csv (berisi identitas pengguna dan kalori terbakar). Penggabungan "
        "(inner join) dilakukan berdasarkan kunci unik User_ID, menghasilkan populasi awal 15.000 data. Sesuai batasan masalah regresi "
        "linear sederhana, hanya dua atribut yang diekstraksi, yaitu atribut Duration (menit) sebagai variabel bebas (X) dan Calories "
        "(kcal) sebagai variabel terikat (Y)."
    )
    add_sec2_body(
        "Tahap pra-pemrosesan data mencakup konversi tipe data ke numerik kontinu, eliminasi nilai kosong (missing values handling), serta "
        "verifikasi batas logis data (X > 0 dan Y > 0). Seluruh 15.000 baris data terverifikasi valid tanpa kehilangan entri. Selanjutnya, "
        "dataset dibagi menjadi data latih (training set) sebanyak 80% (12.000 data) dan data uji (testing set) sebanyak 20% (3.000 data) "
        "dengan nilai seed pembangkit acak random_state = 42 untuk menjamin keterulangan (reproducibility)."
    )
    add_sec2_body(
        "Model Regresi Linear Sederhana dirumuskan secara matematis sebagai Y_hat = b0 + b1 * X. Parameter kemiringan garis (slope) b1 dan "
        "titik potong sumbu (intercept) b0 dihitung menggunakan metode kuadrat terkecil (Ordinary Least Squares / OLS):\n"
        "b1 = (n * Σ(Xi*Yi) - ΣXi * ΣYi) / (n * Σ(Xi²) - (ΣXi)²)\n"
        "b0 = (ΣYi - b1 * ΣXi) / n = Y_bar - b1 * X_bar"
    )
    add_sec2_body(
        "Kinerja model diuji secara ketat hanya pada data uji (n_test = 3.000) menggunakan empat metrik evaluasi regresi: Coefficient of "
        "Determination (R²), Mean Absolute Error (MAE), Mean Squared Error (MSE), dan Root Mean Squared Error (RMSE):\n"
        "R² = 1 - (Σ(Yi - Y_hat_i)² / Σ(Yi - Y_bar)²)\n"
        "MAE = (1/n) * Σ|Yi - Y_hat_i|\n"
        "MSE = (1/n) * Σ(Yi - Y_hat_i)²\n"
        "RMSE = sqrt(MSE)"
    )
    add_sec2_body(
        "Arsitektur sistem dibangun di atas bahasa pemrograman Python 3 menggunakan pustaka Scikit-learn untuk estimasi OLS, Pandas dan NumPy "
        "untuk komputasi matriks, Matplotlib untuk pembangkitan grafik statis resolusi tinggi, serta Flask sebagai pelayan HTTP mikro."
    )

    # ── HASIL DAN PEMBAHASAN (2 KOLOM) ──
    add_sec2_heading("HASIL DAN PEMBAHASAN")

    add_sec2_sub("Perhitungan Manual Regresi Linear Sederhana")
    add_sec2_body(
        "Untuk membuktikan validitas mekanisme kerja algoritma Ordinary Least Squares secara transparan, dilakukan simulasi perhitungan manual "
        "menggunakan sampel berukuran n = 5 data terambil dari populasi."
    )

    add_sec2_caption_tbl("Tabel 1  Tabulasi Perhitungan Manual OLS Sampel")
    t1 = doc.add_table(rows=7, cols=6)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.autofit = False
    col_w1 = [Inches(0.4), Inches(0.55), Inches(0.55), Inches(0.45), Inches(0.55), Inches(0.5)]
    t1_headers = ["No", "X", "Y", "X²", "Y²", "XY"]
    t1_data = [
        ["1", "29,0", "231,0", "841", "53361", "6699"],
        ["2", "14,0", "66,0", "196", "4356", "924"],
        ["3", "5,0", "26,0", "25", "676", "130"],
        ["4", "13,0", "71,0", "169", "5041", "923"],
        ["5", "10,0", "35,0", "100", "1225", "350"],
        ["Σ", "71,0", "429,0", "1331", "64659", "9026"]
    ]
    for c_idx, h in enumerate(t1_headers):
        cell = t1.cell(0, c_idx)
        cell.width = col_w1[c_idx]
        cell.text = h
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(7.5)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_borders(cell, top={'val': 'single', 'sz': 8, 'color': '000000'},
                               bottom={'val': 'single', 'sz': 6, 'color': '000000'})
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=20, bottom=20, left=20, right=20)

    for r_idx, row in enumerate(t1_data):
        for c_idx, val in enumerate(row):
            cell = t1.cell(r_idx + 1, c_idx)
            cell.width = col_w1[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if len(p.runs) > 0:
                p.runs[0].font.size = Pt(7.5)
                if r_idx == 5:
                    p.runs[0].font.bold = True
            is_last = (r_idx == 5)
            set_cell_borders(cell, 
                             top={'val': 'single', 'sz': 6, 'color': '000000'} if is_last else None,
                             bottom={'val': 'single', 'sz': 8, 'color': '000000'} if is_last else None)
            set_cell_margins(cell, top=20, bottom=20, left=20, right=20)
            if is_last:
                set_cell_shading(cell, "F8FAFC")

    add_sec2_body(
        "Rata-rata variabel: X_bar = 71,0 / 5 = 14,20 dan Y_bar = 429,0 / 5 = 85,80. "
        "Koefisien kemiringan (b1) dihitung dengan:\n"
        "b1 = [5(9026) - (71)(429)] / [5(1331) - (71)²]\n"
        "b1 = [45130 - 30459] / [6655 - 5041] = 14671 / 1614 ≈ 9,0898\n"
        "Intercept (b0):\n"
        "b0 = 85,80 - (9,0898 × 14,20) = -43,2752\n"
        "Maka persamaan sampel: Y_hat = -43,2752 + 9,0898 X."
    )

    add_sec2_caption_tbl("Tabel 2  Evaluasi Galat Sampel Manual")
    t2 = doc.add_table(rows=7, cols=6)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.autofit = False
    col_w2 = [Inches(0.35), Inches(0.45), Inches(0.45), Inches(0.55), Inches(0.55), Inches(0.65)]
    t2_headers = ["No", "X", "Y", "Y_hat", "|Y-Y_hat|", "(Y-Y_hat)²"]
    t2_data = [
        ["1", "29,0", "231,0", "220,33", "10,67", "113,87"],
        ["2", "14,0", "66,0", "83,98", "17,98", "323,35"],
        ["3", "5,0", "26,0", "2,17", "23,83", "567,69"],
        ["4", "13,0", "71,0", "74,89", "3,89", "15,15"],
        ["5", "10,0", "35,0", "47,62", "12,62", "159,34"],
        ["Σ", "—", "—", "—", "68,99", "1179,39"]
    ]
    for c_idx, h in enumerate(t2_headers):
        cell = t2.cell(0, c_idx)
        cell.width = col_w2[c_idx]
        cell.text = h
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(7.5)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_borders(cell, top={'val': 'single', 'sz': 8, 'color': '000000'},
                               bottom={'val': 'single', 'sz': 6, 'color': '000000'})
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=20, bottom=20, left=20, right=20)

    for r_idx, row in enumerate(t2_data):
        for c_idx, val in enumerate(row):
            cell = t2.cell(r_idx + 1, c_idx)
            cell.width = col_w2[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if len(p.runs) > 0:
                p.runs[0].font.size = Pt(7.5)
                if r_idx == 5:
                    p.runs[0].font.bold = True
            is_last = (r_idx == 5)
            set_cell_borders(cell, 
                             top={'val': 'single', 'sz': 6, 'color': '000000'} if is_last else None,
                             bottom={'val': 'single', 'sz': 8, 'color': '000000'} if is_last else None)
            set_cell_margins(cell, top=20, bottom=20, left=20, right=20)
            if is_last:
                set_cell_shading(cell, "F8FAFC")

    add_sec2_body(
        "Metrik sampel dihitung sebagai:\n"
        "MAE = 68,99 / 5 = 13,80 kcal\n"
        "MSE = 1179,39 / 5 = 235,88 kcal²\n"
        "RMSE = sqrt(235,88) = 15,36 kcal\n"
        "R² = 1 - (1179,39 / 27850,80) = 0,9577."
    )

    add_sec2_sub("Pelatihan Model Penuh dan Hasil Evaluasi")
    add_sec2_body(
        "Penerapan algoritma OLS pada 12.000 data training menghasilkan model:\n"
        "Calories = -21,6668 + 7,1534 × Duration\n"
        "Pengujian pada 3.000 data testing menghasilkan performa pada Tabel 3."
    )

    add_sec2_caption_tbl("Tabel 3  Hasil Evaluasi Model Penuh Data Testing")
    t3 = doc.add_table(rows=5, cols=3)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3.autofit = False
    col_w3 = [Inches(1.2), Inches(0.9), Inches(0.9)]
    t3_headers = ["Metrik / Parameter", "Nilai", "Satuan"]
    t3_data = [
        ["Koefisien Slope (b1)", "7,1534", "kcal/menit"],
        ["Intercept (b0)", "-21,6668", "kcal"],
        ["R² Score", "0,9145", "—"],
        ["MAE / RMSE", "13,61 / 18,58", "kcal"]
    ]
    for c_idx, h in enumerate(t3_headers):
        cell = t3.cell(0, c_idx)
        cell.width = col_w3[c_idx]
        cell.text = h
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(8)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_borders(cell, top={'val': 'single', 'sz': 8, 'color': '000000'},
                               bottom={'val': 'single', 'sz': 6, 'color': '000000'})
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=20, bottom=20, left=20, right=20)

    for r_idx, row in enumerate(t3_data):
        for c_idx, val in enumerate(row):
            cell = t3.cell(r_idx + 1, c_idx)
            cell.width = col_w3[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            if len(p.runs) > 0:
                p.runs[0].font.size = Pt(8)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            is_last = (r_idx == 3)
            set_cell_borders(cell, bottom={'val': 'single', 'sz': 8, 'color': '000000'} if is_last else None)
            set_cell_margins(cell, top=20, bottom=20, left=20, right=20)

    add_sec2_body(
        "Nilai R² = 0,9145 membuktikan bahwa 91,45% variansi kalori terbakar dapat dijelaskan secara linier oleh durasi olahraga."
    )

    add_sec2_sub("Visualisasi Sebaran Data dan Garis Regresi")
    plot_path = r"C:\Users\zion\Documents\web_app\calorie-predictor\static\images\regression_plot.png"
    if os.path.exists(plot_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture(plot_path, width=Inches(3.0))
        add_sec2_caption_fig("Gambar 1  Visualisasi data aktual dan garis regresi")
    else:
        add_sec2_screenshot_box("1", "Visualisasi data aktual dan garis regresi", "File: static/images/regression_plot.png")

    add_sec2_sub("Implementasi Sistem Antarmuka Web Flask")

    add_sec2_sub_sub("Halaman Utama (Beranda)")
    add_sec2_body(
        "Halaman beranda menyediakan formulir masukan durasi, tombol prediksi, serta kartu penampil estimasi berbasis AJAX."
    )
    add_sec2_screenshot_box("2", "Tampilan beranda Calorie Burn Predictor", "Screenshot form beranda di http://localhost:5000/")

    add_sec2_sub_sub("Validasi Masukan dan Pengujian Prediksi")
    add_sec2_body(
        "Sistem menerapkan validasi ganda pada rentang 1–30 menit untuk mencegah ekstrapolasi tak terkontrol."
    )
    add_sec2_screenshot_box("3", "Hasil prediksi pada durasi 20 menit", "Screenshot kartu hasil estimasi 121.40 kcal")
    add_sec2_screenshot_box("4", "Respon validasi saat input melebihi rentang", "Screenshot alert error rentang 1-30 menit")

    add_sec2_sub_sub("Halaman Analisis Model")
    add_sec2_body(
        "Halaman /analysis menyajikan transparansi metrik, grafik sebaran, dan catatan keterbatasan model."
    )
    add_sec2_screenshot_box("5", "Modul evaluasi metrik pada /analysis", "Screenshot metrik dataset & R²/MAE/MSE/RMSE")
    add_sec2_screenshot_box("6", "Persamaan regresi dan keterbatasan model", "Screenshot persamaan & visual bar R²")

    add_sec2_sub_sub("Halaman Tentang Aplikasi")
    add_sec2_body(
        "Halaman /about memaparkan deskripsi metode, variabel, badge teknologi, dan batasan disclaimer medis."
    )
    add_sec2_screenshot_box("7", "Informasi proyek pada halaman Tentang", "Screenshot deskripsi & disclaimer di /about")

    # ── SIMPULAN DAN SARAN (2 KOLOM) ──
    add_sec2_heading("SIMPULAN DAN SARAN")
    add_sec2_body(
        "Berdasarkan hasil perancangan, eksperimen komputasi, dan pengujian sistem, disimpulkan bahwa metode Regresi Linear Sederhana "
        "berhasil diimplementasikan dengan sangat baik untuk memprediksi jumlah kalori terbakar berdasarkan durasi olahraga. Pengujian performa "
        "pada 3.000 data uji menghasilkan model matematis Calories = -21,6668 + 7,1534 × Duration dengan koefisien determinasi (R²) sebesar "
        "0,9145 serta tingkat kesalahan rata-rata MAE sebesar 13,6135 kcal, MSE sebesar 345,0410 kcal², dan RMSE sebesar 18,5753 kcal. "
        "Pembuktian matematis manual OLS terbukti selaras dengan keluaran pustaka Scikit-learn. Sistem antarmuka web berbasis Flask yang "
        "dikembangkan mampu menyajikan estimasi secara instan menggunakan AJAX, menjaga integritas prediksi melalui validasi rentang "
        "dinamis 1–30 menit, serta memaparkan evaluasi model secara edukatif dan transparan."
    )
    add_sec2_body(
        "Kelemahan model terletak pada penyederhanaan variabel independen yang hanya melibatkan durasi latihan, padahal proses pembakaran "
        "kalori riil dipengaruhi oleh parameter fisiologis lain seperti denyut jantung, massa tubuh, umur, dan jenis kelamin. Oleh karena itu, "
        "saran untuk pengembangan sistem selanjutnya adalah meningkatkan kapabilitas model ke arah Regresi Linear Berganda (Multiple Linear Regression) "
        "atau model non-linear dengan menginkorporasikan variabel denyut jantung dan indeks massa tubuh, serta memperluas rentang observasi "
        "dataset agar sistem mampu mengestimasi sesi olahraga berdurasi panjang dengan presisi yang lebih tinggi."
    )

    # ── PUSTAKA ACUAN (2 KOLOM) ──
    add_sec2_heading("PUSTAKA ACUAN")

    p_buku = doc.add_paragraph()
    p_buku.paragraph_format.space_before = Pt(4)
    p_buku.paragraph_format.space_after = Pt(1)
    p_buku.paragraph_format.keep_with_next = True
    r_buku = p_buku.add_run("Buku")
    r_buku.font.name = 'Arial'
    r_buku.font.size = Pt(9.5)
    r_buku.bold = True

    add_sec2_body("Ainsworth, B. E., et al. (2011). Compendium of Physical Activities: A Second Update of Codes and MET Values. Philadelphia: Lippincott Williams & Wilkins.")
    add_sec2_body("James, G., Witten, D., Hastie, T., and Tibshirani, R. (2021). An Introduction to Statistical Learning: with Applications in Python. New York: Springer.")
    add_sec2_body("Powers, S. K., and Howley, E. T. (2018). Exercise Physiology: Theory and Application to Fitness and Performance. 10th Edition. New York: McGraw-Hill Education.")

    p_jurnal = doc.add_paragraph()
    p_jurnal.paragraph_format.space_before = Pt(4)
    p_jurnal.paragraph_format.space_after = Pt(1)
    p_jurnal.paragraph_format.keep_with_next = True
    r_jurnal = p_jurnal.add_run("Jurnal/Prosiding/Disertasi/Tesis/Skripsi")
    r_jurnal.font.name = 'Arial'
    r_jurnal.font.size = Pt(9.5)
    r_jurnal.bold = True

    add_sec2_body("Kurniawan, R. dan Yuliansyah, H. (2019). Komparasi Algoritma Regresi Linear dan Random Forest Regression dalam Memprediksi Konsumsi Energi Listrik. Jurnal RESTI, Vol. 3 No. 3, hal. 450-456.")
    add_sec2_body("Pratama, A. W. dan Sutikno, T. (2020). Rancang Bangun Aplikasi Prediksi Pengeluaran Energi Atlet Menggunakan Pendekatan Machine Learning Berbasis Web. Jurnal Teknodik, Vol. 24 No. 1, hal. 77-88.")
    add_sec2_body("Widodo, S. dan Nurhadi, M. (2021). Analisis Performa Model Regresi dalam Estimasi Kalori Berdasarkan Sensor Fisiologis. Jurnal Ilmiah Komputasi, Vol. 20 No. 2, hal. 115-124.")

    p_lain = doc.add_paragraph()
    p_lain.paragraph_format.space_before = Pt(4)
    p_lain.paragraph_format.space_after = Pt(1)
    p_lain.paragraph_format.keep_with_next = True
    r_lain = p_lain.add_run("Lain-lain")
    r_lain.font.name = 'Arial'
    r_lain.font.size = Pt(9.5)
    r_lain.bold = True

    add_sec2_body("Grøtte, C. (2020). Calories Burnt Prediction Dataset. Kaggle Public Dataset Repository. https://www.kaggle.com/datasets/fmendes/calories-burnt-prediction (Diunduh tanggal 19 Oktober 2026).")
    add_sec2_body("Pallets Projects. (2023). Flask Documentation (Version 3.0.x). Python Software Foundation. https://flask.palletsprojects.com/ (Diunduh tanggal 19 Oktober 2026).")

    # Save with both new 2-kolom name and original name (with try/except in case file is currently open in Word)
    out_paths = [
        r"C:\Users\zion\Documents\Laporan_Ilmiah_Regresi_Linear_Calorie_Predictor_2Kolom.docx",
        r"C:\Users\zion\Documents\web_app\calorie-predictor\Laporan_Ilmiah_Regresi_Linear_Calorie_Predictor_2Kolom.docx",
        r"C:\Users\zion\Documents\Laporan_Ilmiah_Regresi_Linear_Calorie_Predictor.docx",
        r"C:\Users\zion\Documents\web_app\calorie-predictor\Laporan_Ilmiah_Regresi_Linear_Calorie_Predictor.docx"
    ]

    saved_paths = []
    for p in out_paths:
        try:
            doc.save(p)
            saved_paths.append(p)
        except PermissionError:
            pass # File is currently open in Word

    print("Laporan 2-Kolom berhasil disimpan di:")
    for p in saved_paths:
        print(f" - {p}")

if __name__ == '__main__':
    create_report()
