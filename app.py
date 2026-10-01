import os
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches

# Konfigurasi Tampilan Halaman Streamlit
st.set_page_config(page_title="PRTG Word Notes Automator", page_icon="📝", layout="centered")

st.title("📝 PRTG Excel-to-Word Notes Automator")
st.markdown("Aplikasi web otomatis untuk memasukkan data Downtime/Uptime/Keterangan dari Excel ke bagian `Note :` dokumen Word secara presisi.")

st.markdown("---")

# 1. Upload File Word Mentah
st.subheader("1. Unggah Dokumen Word Mentah (.docx)")
word_file = st.file_uploader("Pilih file Word hasil generate PRTG", type=["docx"], key="word")

# 2. Upload File Excel Olahan
st.subheader("2. Unggah File Excel Olahan (.xlsx)")
excel_file = st.file_uploader("Pilih file Excel yang berisi data catatan (Note)", type=["xlsx", "xls"], key="excel")

st.markdown("---")

# Fungsi inti penggabungan Note
def process_word_notes(word_path, excel_path, output_path):
    try:
        excel_file_obj = pd.ExcelFile(excel_path)
        doc = Document(word_path)
    except Exception as e:
        return False, str(e)

    data_notes = {}
    for sheet_name in excel_file_obj.sheet_names:
        df_sheet = pd.read_excel(excel_file_obj, sheet_name=sheet_name)
        
        required_cols = ['LOKASI', 'Downtime', 'Uptime', 'Keterangan']
        if not all(col in df_sheet.columns for col in required_cols):
            continue
            
        data_notes[sheet_name] = {}
        current_lokasi = None
        
        for _, row in df_sheet.iterrows():
            lokasi_val = str(row['LOKASI']).strip()
            if lokasi_val and lokasi_val.lower() != "nan" and lokasi_val != "-":
                current_lokasi = lokasi_val
                
            if not current_lokasi:
                continue
                
            downtime = str(row['Downtime']).strip()
            uptime = str(row['Uptime']).strip()
            keterangan = str(row['Keterangan']).strip()
            
            if downtime and downtime != "-" and downtime.lower() != "nan":
                detail_teks = f"{downtime}     {uptime}"
                if keterangan and keterangan.lower() != "nan" and keterangan != "-":
                    detail_teks += f"     {keterangan}"
                
                if current_lokasi not in data_notes[sheet_name]:
                    data_notes[sheet_name][current_lokasi] = []
                data_notes[sheet_name][current_lokasi].append(detail_teks)
            elif keterangan and keterangan.lower() != "nan" and keterangan != "-":
                if current_lokasi not in data_notes[sheet_name]:
                    data_notes[sheet_name][current_lokasi] = []
                data_notes[sheet_name][current_lokasi].append(keterangan)

    current_sheet = None
    current_location = None

    for element in doc.element.body:
        if element.tag.endswith('p'):
            from docx.text.paragraph import Paragraph
            p = Paragraph(element, doc)
            text_p = p.text.strip()
            
            if p.style.name.startswith('Heading 1'):
                current_sheet = text_p.split(".", 1)[1].strip() if "." in text_p else text_p
                current_location = None
            elif p.style.name.startswith('Heading 2'):
                current_location = text_p
                
        elif element.tag.endswith('tbl'):
            from docx.table import Table
            table = Table(element, doc)
            table_text = "".join([cell.text for row in table.rows for cell in row.cells])
            
            if "Note" in table_text:
                matched_notes = []
                if current_sheet in data_notes:
                    for loc_key, notes_list in data_notes[current_sheet].items():
                        if current_location and (loc_key.lower() in current_location.lower() or current_location.lower() in loc_key.lower()):
                            matched_notes = notes_list
                            break
                
                if matched_notes:
                    for row in table.rows:
                        for cell in row.cells:
                            if "Note" in cell.text:
                                paragraphs_to_remove = []
                                found_note_section = False
                                for p in cell.paragraphs:
                                    if "Note" in p.text:
                                        found_note_section = True
                                    if found_note_section:
                                        paragraphs_to_remove.append(p)
                                        
                                for p in paragraphs_to_remove:
                                    p_elem = p._p
                                    p_elem.getparent().remove(p_elem)
                                
                                p_note_title = cell.add_paragraph()
                                p_note_title.paragraph_format.space_before = Pt(4)
                                p_note_title.paragraph_format.space_after = Pt(2)
                                run_title = p_note_title.add_run("Note :")
                                run_title.font.name = 'Times New Roman'
                                run_title.font.size = Pt(12)
                                run_title.font.bold = False
                                run_title.font.color.rgb = RGBColor(0, 0, 0)
                                
                                for note_text in matched_notes:
                                    p_bullet = cell.add_paragraph()
                                    p_bullet.paragraph_format.left_indent = Inches(0.35)
                                    p_bullet.paragraph_format.first_line_indent = Inches(-0.2)
                                    p_bullet.paragraph_format.space_before = Pt(2)
                                    p_bullet.paragraph_format.space_after = Pt(4)
                                    
                                    run_bullet = p_bullet.add_run(f"-  {note_text}")
                                    run_bullet.font.name = 'Times New Roman'
                                    run_bullet.font.size = Pt(12)
                                    run_bullet.font.color.rgb = RGBColor(0, 0, 0)

    doc.save(output_path)
    return True, "Berhasil"

# 3. Tombol Eksekusi
if st.button("🚀 Proses Penggabungan Note", type="primary"):
    if word_file is not None and excel_file is not None:
        with st.spinner("Sedang memproses dokumen Word dan mencocokkan data Excel..."):
            # Simpan file upload secara temporary
            temp_word_path = "temp_input.docx"
            temp_excel_path = "temp_input.xlsx"
            output_word_path = "Report_Final_Dengan_Note.docx"
            
            with open(temp_word_path, "wb") as f:
                f.write(word_file.getbuffer())
            with open(temp_excel_path, "wb") as f:
                f.write(excel_file.getbuffer())
                
            success, msg = process_word_notes(temp_word_path, temp_excel_path, output_word_path)
            
            if success:
                st.success("✅ Penggabungan berhasil dilakukan!")
                
                # Sediakan tombol download file hasil
                with open(output_word_path, "rb") as f:
                    st.download_button(
                        label="📥 Download Laporan Word Final",
                        data=f,
                        file_name="Report_PRTG_Final_Dengan_Note.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.error(f"❌ Terjadi kesalahan: {msg}")
    else:
        st.warning("⚠️ Mohon unggah file Word mentah dan file Excel olahan terlebih dahulu!")