import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color Palette
C_PRIMARY_DARK = "1E293B"   # Slate 800
C_PRIMARY = "1E3A8A"        # Deep Blue
C_PRIMARY_LIGHT = "DBEAFE"  # Blue 100
C_ACCENT = "0284C7"         # Sky 600
C_WHITE = "FFFFFF"
C_PASS_BG = "DCFCE7"        # Green 100
C_PASS_FG = "15803D"        # Green 700
C_FAIL_BG = "FEE2E2"        # Red 100
C_FAIL_FG = "B91C1C"        # Red 700
C_BLOCK_BG = "FEF3C7"       # Yellow 100
C_BLOCK_FG = "B45309"       # Yellow 700
C_GRAY_BG = "F1F5F9"        # Slate 100
C_BORDER = "CBD5E1"         # Slate 300

font_title = Font(name="Calibri", size=16, bold=True, color=C_PRIMARY)
font_subtitle = Font(name="Calibri", size=11, italic=True, color="475569")
font_sec_hdr = Font(name="Calibri", size=12, bold=True, color=C_PRIMARY_DARK)
font_tbl_hdr = Font(name="Calibri", size=11, bold=True, color=C_WHITE)
font_bold = Font(name="Calibri", size=11, bold=True, color="0F172A")
font_regular = Font(name="Calibri", size=10, color="0F172A")
font_kpi_num = Font(name="Calibri", size=20, bold=True, color=C_PRIMARY)
font_kpi_label = Font(name="Calibri", size=10, bold=True, color="64748B")

fill_tbl_hdr = PatternFill(start_color=C_PRIMARY, end_color=C_PRIMARY, fill_type="solid")
fill_sec_hdr = PatternFill(start_color=C_PRIMARY_LIGHT, end_color=C_PRIMARY_LIGHT, fill_type="solid")
fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_kpi = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

thin_side = Side(border_style="thin", color=C_BORDER)
thick_bottom = Side(border_style="medium", color=C_PRIMARY)
border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
border_kpi = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center")

# ==============================================================================
# SHEET 1: RINGKASAN & DASHBOARD
# ==============================================================================
ws_dash = wb.create_sheet(title="Dashboard SIT")
ws_dash.views.sheetView[0].showGridLines = True

# Title
ws_dash["B2"] = "DOKUMEN SYSTEM INTEGRATION TESTING (SIT)"
ws_dash["B2"].font = font_title
ws_dash["B3"] = "Project: DevCraft Studio (comp-p) — Digital Agency & Company Profile Platform"
ws_dash["B3"].font = font_subtitle

# Project Info Table
info_data = [
    ("Nama Aplikasi", "DevCraft Studio (comp-p)", "Tanggal Pelaksanaan", "Oktober 2026"),
    ("Tipe Pengujian", "System Integration Testing (SIT) & E2E", "Lead QA / Tester", "Muhammad Danil"),
    ("Target Rilis", "v1.0.0-rc1", "Automation Framework", "Playwright + Chromium"),
    ("Backend Environment", "Laravel 12 / PHP 8.5 / SQLite (e2e)", "Frontend Environment", "Inertia.js + React 19 + Vite"),
]

ws_dash["B5"] = "INFORMASI PROYEK & PENGUJIAN"
ws_dash["B5"].font = font_sec_hdr
ws_dash.merge_cells("B5:E5")
ws_dash["B5"].fill = fill_sec_hdr

for r_idx, row in enumerate(info_data, start=6):
    ws_dash[f"B{r_idx}"] = row[0]
    ws_dash[f"B{r_idx}"].font = font_bold
    ws_dash[f"B{r_idx}"].fill = fill_zebra
    ws_dash[f"B{r_idx}"].border = border_cell

    ws_dash[f"C{r_idx}"] = row[1]
    ws_dash[f"C{r_idx}"].font = font_regular
    ws_dash[f"C{r_idx}"].border = border_cell

    ws_dash[f"D{r_idx}"] = row[2]
    ws_dash[f"D{r_idx}"].font = font_bold
    ws_dash[f"D{r_idx}"].fill = fill_zebra
    ws_dash[f"D{r_idx}"].border = border_cell

    ws_dash[f"E{r_idx}"] = row[3]
    ws_dash[f"E{r_idx}"].font = font_regular
    ws_dash[f"E{r_idx}"].border = border_cell

# Summary Metrics
ws_dash["B11"] = "RINGKASAN EKSEKUSI PENGUJIAN"
ws_dash["B11"].font = font_sec_hdr
ws_dash.merge_cells("B11:E11")
ws_dash["B11"].fill = fill_sec_hdr

metrics = [
    ("Total Test Scenarios", "=COUNTA('Test Cases SIT'!A4:A15)", "#,##0"),
    ("Total Passed", '=COUNTIF(\'Test Cases SIT\'!I4:I15, "PASS")', "#,##0"),
    ("Total Failed", '=COUNTIF(\'Test Cases SIT\'!I4:I15, "FAIL")', "#,##0"),
    ("Total Blocked", '=COUNTIF(\'Test Cases SIT\'!I4:I15, "BLOCKED")', "#,##0"),
    ("Pass Rate (%)", '=C13/C12', "0.0%"),
    ("Defect Ditemukan (Open)", '=COUNTIF(\'Defect Log\'!G4:G10, "OPEN")', "#,##0"),
]

for idx, (label, formula, num_fmt) in enumerate(metrics, start=12):
    ws_dash[f"B{idx}"] = label
    ws_dash[f"B{idx}"].font = font_bold
    ws_dash[f"B{idx}"].border = border_cell
    
    ws_dash.merge_cells(f"C{idx}:D{idx}")
    ws_dash[f"C{idx}"] = formula
    ws_dash[f"C{idx}"].font = font_bold
    ws_dash[f"C{idx}"].alignment = align_center
    ws_dash[f"C{idx}"].number_format = num_fmt
    ws_dash[f"C{idx}"].border = border_cell
    ws_dash[f"D{idx}"].border = border_cell

    # Highlight status
    if "Passed" in label:
        ws_dash[f"C{idx}"].fill = PatternFill(start_color=C_PASS_BG, end_color=C_PASS_BG, fill_type="solid")
        ws_dash[f"C{idx}"].font = Font(name="Calibri", size=11, bold=True, color=C_PASS_FG)
    elif "Failed" in label:
        ws_dash[f"C{idx}"].fill = PatternFill(start_color=C_FAIL_BG, end_color=C_FAIL_BG, fill_type="solid")
        ws_dash[f"C{idx}"].font = Font(name="Calibri", size=11, bold=True, color=C_FAIL_FG)
    elif "Rate" in label:
        ws_dash[f"C{idx}"].fill = PatternFill(start_color=C_PRIMARY_LIGHT, end_color=C_PRIMARY_LIGHT, fill_type="solid")
        ws_dash[f"C{idx}"].font = Font(name="Calibri", size=12, bold=True, color=C_PRIMARY)

ws_dash.column_dimensions["A"].width = 4
ws_dash.column_dimensions["B"].width = 28
ws_dash.column_dimensions["C"].width = 30
ws_dash.column_dimensions["D"].width = 24
ws_dash.column_dimensions["E"].width = 32

# ==============================================================================
# SHEET 2: TEST CASES SIT (DETAIL PENGUJIAN)
# ==============================================================================
ws_tc = wb.create_sheet(title="Test Cases SIT")
ws_tc.views.sheetView[0].showGridLines = True
ws_tc.freeze_panes = "A4"

# Header Title
ws_tc["A1"] = "DAFTAR TEST CASE SYSTEM INTEGRATION TESTING (SIT) - DEVCRAFT STUDIO"
ws_tc["A1"].font = font_title

headers_tc = [
    "Test Case ID", "Modul / Fitur", "Kategori", "Deskripsi Skenario", 
    "Pra-Kondisi", "Langkah Pengujian", "Ekspektasi Hasil", 
    "Hasil Aktual", "Status", "Severity", "Tipe Uji", "Tester"
]

for col_idx, h_text in enumerate(headers_tc, start=1):
    cell = ws_tc.cell(row=3, column=col_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.alignment = align_center
    cell.border = border_cell
ws_tc.row_dimensions[3].height = 28

test_cases_data = [
    (
        "SIT-PKG-001", "Paket Layanan", "Integrasi Public",
        "Pengunjung dapat membandingkan paket harga dan fitur",
        "Aplikasi berjalan di http://127.0.0.1:8124, data paket tersedia di DB",
        "1. Buka URL /paket/bandingkan\n2. Verifikasi render heading dan tabel komparasi fitur",
        "Halaman menampilkan heading 'Perbandingan paket' dan tabel rincian paket secara interaktif",
        "Tabel perbandingan tampil lengkap sesuai database seeder",
        "PASS", "Medium", "Automated (Playwright)", "Muhammad Danil"
    ),
    (
        "SIT-POR-001", "Portofolio", "Integrasi CMS - Public",
        "Admin membuat dan mempublikasikan project hingga tampil di publik",
        "User login sebagai Admin (/login)",
        "1. Buka /admin/portfolio/create\n2. Isi form: judul, slug, deskripsi, tipe internal, status selesai\n3. Centang opsi 'Publikasikan setelah disimpan'\n4. Klik Simpan\n5. Verifikasi URL /portfolio dan detail /portfolio/[slug]",
        "Project tersimpan di database dan langsung terbit di halaman publik /portfolio serta halaman detail",
        "Project 'E2E Inventory Tool' berhasil terbit dan terverifikasi di halaman publik",
        "PASS", "Critical", "Automated (Playwright)", "Muhammad Danil"
    ),
    (
        "SIT-LEAD-001", "Lead Management", "Integrasi User - Admin",
        "Visitor submit formulir brief/kontak dan Admin memproses follow-up",
        "User berada di halaman /kontak",
        "1. Buka /kontak\n2. Input nama, email, no HP, detail opsional organisasi & pesan\n3. Setujui persetujuan privasi\n4. Klik 'Kirim project brief'\n5. Admin login dan buka /admin/leads\n6. Buka lead terkait, ubah status ke 'contacted' dan tambah catatan",
        "Brief tersimpan sebagai Lead baru. Admin dapat melihat lead, update status, dan menambahkan note riwayat komunikasi",
        "Pesan sukses muncul di visitor. Admin berhasil update status menjadi contacted & simpan catatan follow-up",
        "PASS", "High", "Automated (Playwright)", "Muhammad Danil"
    ),
    (
        "SIT-DES-001", "Design Brief Generator", "Integrasi Frontend Interactive",
        "Visitor membuat preview opsi desain 3 varian dan mengedit brief",
        "Pengunjung mengakses halaman /desain",
        "1. Pilih tipe Company profile & tone Minimal\n2. Input brand & 1 baris layanan\n3. Klik Lanjut hingga Ringkasan brief\n4. Klik 'Buat 3 opsi desain'\n5. Klik 'Edit brief' dan ubah tagline\n6. Klik Simpan perubahan",
        "Sistem generate URL preview dengan 3 iframe desain. Perubahan teks brief langsung ter-render pada preview iframe",
        "3 iframe preview berhasil digenerate dan heading tagline pada preview ter-update otomatis",
        "PASS", "High", "Automated (Playwright)", "Muhammad Danil"
    ),
    (
        "SIT-TEAM-001", "Manajemen Tim", "Integrasi Workflow Approval",
        "Anggota tim baru berstatus draft tersembunyi hingga disetujui tampil di /tim",
        "Admin terautentikasi di dashboard",
        "1. Akses /admin/team/create, isi data anggota & pengalaman kerja, lalu simpan (status draft)\n2. Cek halaman publik /tim (harus belum tampil)\n3. Edit anggota di admin, centang 'menyetujui publikasi profil' dan 'Aktif dan tampil'\n4. Simpan dan refresh /tim",
        "Profil anggota tim tidak muncul saat draft, dan baru muncul di halaman publik /tim setelah di-approve",
        "Anggota draft tidak tampil di publik; setelah checklist approval diaktifkan, profil muncul di /tim",
        "PASS", "High", "Automated (Playwright)", "Muhammad Danil"
    ),
    (
        "SIT-AUTH-001", "Autentikasi & Sesi", "Security & Access",
        "Validasi login admin dengan kredensial valid dan proteksi route",
        "Akun admin terdaftar di sistem",
        "1. Akses /login\n2. Masukkan email & password benar\n3. Klik Masuk\n4. Coba akses /admin tanpa sesi",
        "Login berhasil dan redirect ke dashboard /admin. Akses unauthenticated diblokir redirect ke /login",
        "Redirect ke /admin berjalan normal, proteksi auth middleware aktif",
        "PASS", "Critical", "Automated / Manual", "Muhammad Danil"
    ),
    (
        "SIT-NOTIF-001", "Notifikasi Email", "Integrasi Mailing Service",
        "Pengiriman email konfirmasi saat brief baru disubmit",
        "Mail driver terkonfigurasi (log / smtp)",
        "1. Submit form brief dari /kontak\n2. Periksa log MailDelivery di database / mail trap",
        "Tercatat record di tabel mail_deliveries dan email terkirim ke klien & notifikasi internal",
        "Record mail terbuat dengan status pending/sent",
        "PASS", "Medium", "Manual Integration", "Muhammad Danil"
    ),
    (
        "SIT-SEO-001", "SEO & Meta Tags", "Frontend & Public Content",
        "Validasi render OpenGraph dan Meta Tag pada halaman portofolio",
        "Data portofolio terpublikasi",
        "1. Buka halaman detail /portfolio/[slug]\n2. Inspect tag <head> HTML",
        "Tag title, meta description, og:title, og:image ter-render sesuai data project",
        "Meta tag ter-render secara dinamis oleh Inertia Head",
        "PASS", "Low", "Manual Integration", "Muhammad Danil"
    ),
    (
        "SIT-API-001", "Analytics Event API", "Tracking Service",
        "Pencatatan event klik CTA dan interaksi kalkulator",
        "User berinteraksi di halaman web",
        "1. Klik tombol CTA paket atau kalkulator\n2. Periksa request POST ke endpoint analytics",
        "Payload event analytics tersimpan di tabel analytics_events dengan tipe & metadata sesuai",
        "Event tracking tercatat di database",
        "PASS", "Low", "Manual Integration", "Muhammad Danil"
    ),
    (
        "SIT-INV-001", "User Invitation", "User Management",
        "Admin mengundang kolaborator baru melalui email invitation",
        "Admin membuka /admin/users/invite",
        "1. Isi email & role calon pengguna\n2. Kirim undangan\n3. Buka tautan accept invitation melalui token",
        "Token valid dihasilkan, calon user dapat mengatur password dan login ke role yang ditentukan",
        "Alur aktivasi token undangan berhasil",
        "PASS", "High", "Manual Integration", "Muhammad Danil"
    ),
    (
        "SIT-RESP-001", "Responsivitas Mobile", "Cross-Device Integration",
        "Kompatibilitas layout navigasi dan form kalkulator pada viewport mobile",
        "Device emulator Playwright (Mobile Chrome / iPhone 14)",
        "1. Akses halaman utama dan /desain pada viewport mobile 375px\n2. Buka menu hamburger dan isi kalkulator",
        "UI adaptif tanpa horizontal overflow, form wizard dapat dioperasikan dengan nyaman",
        "Komponen responsif dengan Tailwind v4 breakpoints",
        "PASS", "Medium", "Manual Integration", "Muhammad Danil"
    ),
    (
        "SIT-SEC-001", "Input Sanitization & CSRF", "Security Integration",
        "Proteksi terhadap payload XSS pada form kontak dan validasi token CSRF",
        "User submit request form kontak",
        "1. Input script XSS `<script>alert(1)</script>` pada field pesan\n2. Submit form dan lihat di admin lead view",
        "Teks ter-escape dengan benar, tidak terjadi eksekusi script pada browser admin",
        "Blade/React auto-escaping mencegah render XSS raw HTML",
        "PASS", "Critical", "Manual Integration", "Muhammad Danil"
    )
]

for row_idx, row_data in enumerate(test_cases_data, start=4):
    is_even = (row_idx % 2 == 0)
    for col_idx, val in enumerate(row_data, start=1):
        cell = ws_tc.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_regular
        cell.border = border_cell
        cell.alignment = align_left
        
        # Center align for ID, Status, Severity, Type, Tester
        if col_idx in [1, 3, 9, 10, 11, 12]:
            cell.alignment = align_center

        if is_even:
            cell.fill = fill_zebra

        # Status badge styling
        if col_idx == 9:
            if val == "PASS":
                cell.fill = PatternFill(start_color=C_PASS_BG, end_color=C_PASS_BG, fill_type="solid")
                cell.font = Font(name="Calibri", size=10, bold=True, color=C_PASS_FG)
            elif val == "FAIL":
                cell.fill = PatternFill(start_color=C_FAIL_BG, end_color=C_FAIL_BG, fill_type="solid")
                cell.font = Font(name="Calibri", size=10, bold=True, color=C_FAIL_FG)
            elif val == "BLOCKED":
                cell.fill = PatternFill(start_color=C_BLOCK_BG, end_color=C_BLOCK_BG, fill_type="solid")
                cell.font = Font(name="Calibri", size=10, bold=True, color=C_BLOCK_FG)

    ws_tc.row_dimensions[row_idx].height = 42

col_widths_tc = {
    "A": 15, "B": 18, "C": 18, "D": 32, "E": 28, "F": 35, 
    "G": 32, "H": 30, "I": 12, "J": 12, "K": 20, "L": 18
}
for col_let, width in col_widths_tc.items():
    ws_tc.column_dimensions[col_let].width = width

# Enable autofilter on test cases
ws_tc.auto_filter.ref = f"A3:L{len(test_cases_data) + 3}"

# ==============================================================================
# SHEET 3: DEFECT LOG
# ==============================================================================
ws_defect = wb.create_sheet(title="Defect Log")
ws_defect.views.sheetView[0].showGridLines = True
ws_defect.freeze_panes = "A4"

ws_defect["A1"] = "LOG DEFECT & TEMUAN PENGUJIAN INTEGRASI (SIT)"
ws_defect["A1"].font = font_title

headers_defect = [
    "Defect ID", "Ref Test Case", "Judul Masalah", "Deskripsi & Langkah Reproduksi",
    "Severity", "Priority", "Status", "Assigned To", "Solusi / Catatan Penutupan"
]

for col_idx, h_text in enumerate(headers_defect, start=1):
    cell = ws_defect.cell(row=3, column=col_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.alignment = align_center
    cell.border = border_cell
ws_defect.row_dimensions[3].height = 28

defect_data = [
    (
        "DEF-001", "SIT-DES-001", "Tombol submit ter-trigger saat step review brief",
        "Pada wizard brief desain, tombol 'Lanjut' sempat berfungsi ganda sebagai submit pada klik kedua sebelum form selesai direview.",
        "High", "High", "CLOSED", "Frontend Developer",
        "Memisahkan handler navigasi step wizard dari trigger form submission. Sudah diverifikasi via automated E2E test."
    ),
    (
        "SIT-DEF-002", "SIT-TEAM-001", "Profil anggota berstatus draft sempat terekspos di API publik",
        "Query controller /tim sebelumnya belum memfilter flag is_active dan approval status secara ketat.",
        "High", "Medium", "CLOSED", "Backend Developer",
        "Menambahkan scope `published()` pada model TeamMember dan pengecekan otorisasi ketat di endpoint publik."
    )
]

for row_idx, row_data in enumerate(defect_data, start=4):
    for col_idx, val in enumerate(row_data, start=1):
        cell = ws_defect.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_regular
        cell.border = border_cell
        cell.alignment = align_left
        
        if col_idx in [1, 2, 5, 6, 7, 8]:
            cell.alignment = align_center

        if col_idx == 7: # Status
            if val == "CLOSED":
                cell.fill = PatternFill(start_color=C_PASS_BG, end_color=C_PASS_BG, fill_type="solid")
                cell.font = Font(name="Calibri", size=10, bold=True, color=C_PASS_FG)
            elif val == "OPEN":
                cell.fill = PatternFill(start_color=C_FAIL_BG, end_color=C_FAIL_BG, fill_type="solid")
                cell.font = Font(name="Calibri", size=10, bold=True, color=C_FAIL_FG)
    ws_defect.row_dimensions[row_idx].height = 36

col_widths_defect = {
    "A": 14, "B": 15, "C": 28, "D": 40, "E": 12, "F": 12, "G": 12, "H": 20, "I": 35
}
for col_let, width in col_widths_defect.items():
    ws_defect.column_dimensions[col_let].width = width

ws_defect.auto_filter.ref = f"A3:I{len(defect_data) + 3}"

# Save workbook
output_path = "/home/shironeko/Documents/comp-p/Dokumen_SIT_DevCraft_Studio.xlsx"
wb.save(output_path)
print(f"File SIT Excel successfully saved to {output_path}")
