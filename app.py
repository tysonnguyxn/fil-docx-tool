import io
import re
import zipfile
from copy import deepcopy
import pandas as pd
import streamlit as st
import docx
from docxtpl import DocxTemplate

st.set_page_config(
    page_title="Công cụ điền Word & Bảng tự động",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Tự động điền dữ liệu Word & Bảng danh sách")
st.markdown("""
* **Bảo toàn 100% định dạng chi tiết**: Từ khóa `{{key}}` được định dạng như thế nào (**Phông chữ**, **Cỡ chữ**, **In đậm**, *In nghiêng*, <u>Gạch chân</u>, Màu sắc...) thì giá trị thay thế vào sẽ **ăn theo y hệt**!
* **Giữ nguyên văn bản cố định trong ô**: Trong 1 ô có thể vừa có chữ cố định (màu đen) vừa có từ khóa thay thế (màu đỏ).
* **Bảng danh sách động**: Tự động nhân bản hàng mẫu của bảng theo danh sách Excel và tự động đánh số thứ tự 1, 2, 3...
""")

st.write("---")

# 1. Giao diện Upload File
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. File mẫu Word (.docx)")
    uploaded_docs = st.file_uploader(
        "Tải lên các file Word",
        type=["docx"],
        accept_multiple_files=True,
        help="Có thể chọn 1 hoặc nhiều file Word cùng lúc"
    )
    if uploaded_docs:
        st.success(f"Đã chọn {len(uploaded_docs)} file Word.")

with col2:
    st.subheader("2. Excel: Thông tin chung")
    uploaded_general = st.file_uploader(
        "File Excel Key-Value (Cột A: Key, Cột B: Giá trị)",
        type=["xlsx", "xls"],
        help="Chứa thông tin đơn lẻ như tên đơn vị, ngày tháng..."
    )
    if uploaded_general:
        st.success("Đã nạp file thông tin chung.")

with col3:
    st.subheader("3. Excel: Dữ liệu Bảng (Tùy chọn)")
    uploaded_table = st.file_uploader(
        "File Excel Danh sách (Hàng 1: Key, các hàng dưới: Dữ liệu)",
        type=["xlsx", "xls"],
        help="Ví dụ: danh sách nhân sự, hàng hóa..."
    )
    if uploaded_table:
        st.success("Đã nạp file dữ liệu bảng.")

# 2. Xử lý lưu Session State để hỗ trợ Chỉnh sửa dữ liệu trực tiếp (st.data_editor)
if uploaded_general:
    gen_file_id = f"{uploaded_general.name}_{uploaded_general.size}"
    if st.session_state.get("last_gen_file") != gen_file_id:
        try:
            df_gen_raw = pd.read_excel(uploaded_general, header=None)
            df_gen_raw = df_gen_raw.dropna(subset=[0])
            gen_rows = []
            for _, r in df_gen_raw.iterrows():
                k = str(r[0]).strip().replace("{{", "").replace("}}", "")
                v = "" if pd.isna(r[1]) else str(r[1]).strip()
                if k:
                    gen_rows.append({"Từ khóa (Key)": k, "Giá trị thay thế": v})
            st.session_state.df_general = pd.DataFrame(gen_rows)
            st.session_state.last_gen_file = gen_file_id
        except Exception as e:
            st.error(f"Lỗi đọc file Excel thông tin chung: {e}")
            st.session_state.df_general = pd.DataFrame(columns=["Từ khóa (Key)", "Giá trị thay thế"])

if uploaded_table:
    tbl_file_id = f"{uploaded_table.name}_{uploaded_table.size}"
    if st.session_state.get("last_tbl_file") != tbl_file_id:
        try:
            df_tbl_raw = pd.read_excel(uploaded_table)
            clean_cols = [str(c).strip().replace("{{", "").replace("}}", "") for c in df_tbl_raw.columns]
            df_tbl_raw.columns = clean_cols
            df_tbl_raw = df_tbl_raw.dropna(how="all").fillna("")
            df_tbl_raw = df_tbl_raw.astype(str)
            st.session_state.df_table = df_tbl_raw
            st.session_state.last_tbl_file = tbl_file_id
        except Exception as e:
            st.error(f"Lỗi đọc file Excel bảng dữ liệu: {e}")
            st.session_state.df_table = pd.DataFrame()

# 3. Khu vực Kiểm tra và Chỉnh sửa trực tiếp dữ liệu (Editable Data Tables)
has_general = "df_general" in st.session_state and not st.session_state.df_general.empty
has_table = "df_table" in st.session_state and not st.session_state.df_table.empty

edited_general_dict = {}
edited_table_list = []

if has_general or has_table:
    st.write("---")
    st.subheader("✏️ 3. Kiểm tra & Chỉnh sửa trực tiếp dữ liệu Excel")
    st.caption("💡 **Mẹo:** Bạn có thể **nhấp đúp chuột vào từng ô** để sửa trực tiếp, hoặc nhấn nút **(+)** ở cuối bảng để thêm dòng mới.")

    prev_col1, prev_col2 = st.columns(2)

    with prev_col1:
        if has_general:
            st.markdown("**Thông tin chung (Key - Value):**")
            edited_gen_df = st.data_editor(
                st.session_state.df_general,
                num_rows="dynamic",
                use_container_width=True,
                height=260,
                key="editor_general"
            )
            for _, r in edited_gen_df.iterrows():
                k = str(r.get("Từ khóa (Key)", "")).strip().replace("{{", "").replace("}}", "")
                v = str(r.get("Giá trị thay thế", "")).strip()
                if k:
                    edited_general_dict[k] = v
        else:
            st.info("Chưa tải file thông tin chung.")

    with prev_col2:
        if has_table:
            st.markdown(f"**Dữ liệu Bảng danh sách:**")
            edited_tbl_df = st.data_editor(
                st.session_state.df_table,
                num_rows="dynamic",
                use_container_width=True,
                height=260,
                key="editor_table"
            )
            clean_tbl_df = edited_tbl_df.dropna(how="all")
            for _, r in clean_tbl_df.iterrows():
                row_dict = {}
                for col in clean_tbl_df.columns:
                    row_dict[str(col).strip()] = str(r[col]).strip()
                edited_table_list.append(row_dict)
        else:
            st.info("Chưa tải file dữ liệu bảng.")

# ==============================================================================
# CORE FORMAT-PRESERVING ENGINE (CHUẨN ECMA-376 OPENXML)
# ==============================================================================

def enrich_rPr_from_pPr(r_elem, p_elem):
    """
    Kế thừa các thuộc tính định dạng (font chữ, cỡ chữ, màu sắc, bold...) từ cấp paragraph
    nếu run chưa khai báo cục bộ, đảm bảo giá trị điền vào ăn khớp 100% với giao diện Word.
    """
    pPr = p_elem.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    if pPr is None:
        return
    p_rPr = pPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
    if p_rPr is None:
        return

    rPr = r_elem.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
    if rPr is None:
        rPr = deepcopy(p_rPr)
        r_elem.insert(0, rPr)
        return

    # Nếu rPr đã có, bổ sung các thẻ còn thiếu từ p_rPr (font, sz, szCs, color...)
    for child in p_rPr:
        tag = child.tag
        if rPr.find(tag) is None:
            rPr.append(deepcopy(child))

def replace_placeholder_in_paragraph(paragraph, key, value):
    """
    Thay thế {{key}} bằng value trong 1 paragraph mà BẢO TOÀN 100% ĐỊNH DẠNG
    (Phông chữ, Cỡ chữ, Bold, Italic, Underline, Màu sắc, Mã màu hex...).
    Chuẩn hóa OpenXML: <w:rPr> luôn đứng đầu tiên trong <w:r> trước <w:t>.
    """
    placeholder = "{{" + str(key) + "}}"
    if placeholder not in paragraph.text:
        return False

    val_str = str(value)
    changed = False

    while placeholder in paragraph.text:
        # Trường hợp 1: Placeholder nằm trọn vẹn trong 1 Run duy nhất
        single_matched = False
        for r in paragraph.runs:
            if placeholder in r.text:
                r.text = r.text.replace(placeholder, val_str, 1)
                enrich_rPr_from_pPr(r._r, paragraph._p)
                single_matched = True
                changed = True
                break
        if single_matched:
            continue

        # Trường hợp 2: Word cắt rời placeholder qua nhiều Run (VD: Run1={{, Run2=key, Run3=}})
        runs = paragraph.runs
        full_text = ""
        run_indices = []
        for idx, r in enumerate(runs):
            start = len(full_text)
            full_text += r.text
            end = len(full_text)
            run_indices.append((idx, start, end))

        match = re.search(re.escape(placeholder), full_text)
        if not match:
            break

        p_start, p_end = match.span()
        matched_runs = []
        best_format_run = None

        for idx, start, end in run_indices:
            if start < p_end and end > p_start:
                matched_runs.append(idx)
                r = runs[idx]
                # Ưu tiên lấy định dạng của run chứa tên từ khóa
                if str(key) in r.text:
                    best_format_run = r
                elif best_format_run is None and r._r.rPr is not None and len(r._r.rPr) > 0:
                    best_format_run = r

        if not matched_runs:
            break

        if best_format_run is None:
            best_format_run = runs[matched_runs[0]]

        target_run = best_format_run
        rPr_copy = deepcopy(best_format_run._r.rPr) if best_format_run._r.rPr is not None else None

        first_idx = matched_runs[0]
        last_idx = matched_runs[-1]
        first_run = runs[first_idx]
        last_run = runs[last_idx]

        first_start = run_indices[first_idx][1]
        last_end = run_indices[last_idx][2]

        prefix = full_text[first_start:p_start]
        suffix = full_text[p_end:last_end]

        # Dọn dẹp các run khác ngoài target_run
        for m_idx in matched_runs:
            r = runs[m_idx]
            if r is target_run:
                continue
            if m_idx == first_idx and prefix:
                r.text = prefix
            elif m_idx == last_idx and suffix:
                r.text = suffix
            else:
                r.text = ""

        # Gán giá trị vào target_run
        if target_run is first_run and prefix:
            target_run.text = prefix + val_str
        elif target_run is last_run and suffix:
            target_run.text = val_str + suffix
        else:
            target_run.text = val_str

        # ĐẢM BẢO rPr LUÔN NẰM Ở ĐẦU TIÊN (Trước <w:t>) để Word không bao giờ bị mất định dạng
        if rPr_copy is not None:
            if target_run._r.rPr is not None:
                target_run._r.remove(target_run._r.rPr)
            target_run._r.insert(0, deepcopy(rPr_copy))

        # Kế thừa thêm font/size từ paragraph nếu run chưa có
        enrich_rPr_from_pPr(target_run._r, paragraph._p)

        changed = True

    return changed

def replace_in_all_paragraphs(doc, context_dict):
    """Quét toàn bộ tài liệu (paragraphs, header, footer, tables) để thay thế context_dict giữ nguyên 100% định dạng."""
    if not context_dict:
        return

    # 1. Các đoạn văn bản chính
    for p in doc.paragraphs:
        for k, v in context_dict.items():
            replace_placeholder_in_paragraph(p, k, v)

    # 2. Các ô trong tất cả các bảng
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in context_dict.items():
                        replace_placeholder_in_paragraph(p, k, v)

    # 3. Header và Footer của các section
    for section in doc.sections:
        for p in section.header.paragraphs:
            for k, v in context_dict.items():
                replace_placeholder_in_paragraph(p, k, v)
        for p in section.footer.paragraphs:
            for k, v in context_dict.items():
                replace_placeholder_in_paragraph(p, k, v)

def process_dynamic_tables(doc, items):
    """Nhân bản hàng mẫu của bảng theo danh sách items, bảo tồn 100% định dạng và các văn bản cố định trong ô."""
    if not items:
        return doc

    table_keys = list(items[0].keys())

    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            row_text = "".join([c.text for c in row.cells])
            # Hàng mẫu là hàng chứa ít nhất một placeholder trùng với cột Excel
            if any(f"{{{{{k}}}}}" in row_text for k in table_keys):
                template_tr = row._tr
                parent = template_tr.getparent()
                tr_idx = parent.index(template_tr)

                for index, raw_item in enumerate(items):
                    item = dict(raw_item)
                    # Tự động hỗ trợ điền số thứ tự nếu trong mẫu Word có {{tt}} hoặc {{stt}}
                    if "tt" not in item:
                        item["tt"] = str(index + 1)
                    if "stt" not in item:
                        item["stt"] = str(index + 1)

                    # Nhân bản nguyên vẹn hàng mẫu (giữ 100% viền, chiều cao, căn lề và văn bản cố định)
                    new_tr = deepcopy(template_tr)
                    row_obj = docx.table._Row(new_tr, table)
                    for cell in row_obj.cells:
                        for p in cell.paragraphs:
                            for k, v in item.items():
                                replace_placeholder_in_paragraph(p, k, v)
                    parent.insert(tr_idx, new_tr)
                    tr_idx += 1

                parent.remove(template_tr)
                break
    return doc

# ==============================================================================
# 4. Cấu hình đóng gói ZIP & Thực hiện xuất file
# ==============================================================================

if uploaded_docs and (edited_general_dict or edited_table_list):
    st.write("---")
    st.subheader("📦 4. Cấu hình xuất file & Tải về")

    cfg_col1, cfg_col2 = st.columns([2, 1])
    with cfg_col1:
        folder_in_zip = st.text_input(
            "📁 Tên thư mục nhóm các file bên trong file ZIP:",
            value="Ho_So_Xuat_Word",
            help="Tất cả các file Word sau khi điền sẽ được gom gọn gàng vào thư mục này bên trong file ZIP"
        ).strip().replace("\\", "/").strip("/")
        if not folder_in_zip:
            folder_in_zip = "Ket_Qua_Word"

    with cfg_col2:
        zip_file_name = st.text_input(
            "Tên file ZIP tải về:",
            value="ket_qua_tong_hop.zip"
        ).strip()
        if not zip_file_name.endswith(".zip"):
            zip_file_name += ".zip"

    if st.button("🚀 Bắt đầu thay thế & Đóng gói file ZIP", type="primary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            total_files = len(uploaded_docs)

            for index, doc_file in enumerate(uploaded_docs):
                status_text.text(f"Đang xử lý ({index + 1}/{total_files}): {doc_file.name}...")

                # Mở tài liệu Word
                doc_bytes = io.BytesIO(doc_file.getvalue())
                doc_obj = docx.Document(doc_bytes)

                # Bước 1: Nhân bản các hàng trong bảng danh sách (giữ 100% định dạng từng ô & văn bản cố định)
                if edited_table_list:
                    doc_obj = process_dynamic_tables(doc_obj, edited_table_list)

                # Bước 2: Thay thế toàn bộ từ khóa thông tin chung (giữ 100% bold, italic, font, color, size...)
                if edited_general_dict:
                    replace_in_all_paragraphs(doc_obj, edited_general_dict)

                # Bước 3: Nếu còn cú pháp Jinja2 đặc biệt (ví dụ {% ... %}), cho docxtpl xử lý nốt
                intermediate_stream = io.BytesIO()
                doc_obj.save(intermediate_stream)
                intermediate_stream.seek(0)

                raw_bytes = intermediate_stream.getvalue()
                if b"{%" in raw_bytes:
                    tpl = DocxTemplate(intermediate_stream)
                    merged_context = dict(edited_general_dict)
                    merged_context["items"] = edited_table_list
                    merged_context["danh_sach"] = edited_table_list
                    merged_context["table"] = edited_table_list
                    tpl.render(merged_context)
                    final_output = io.BytesIO()
                    tpl.save(final_output)
                    final_output.seek(0)
                else:
                    final_output = intermediate_stream

                # Đưa file vào thư mục nhóm trong ZIP
                zip_path = f"{folder_in_zip}/filled_{doc_file.name}"
                zip_file.writestr(zip_path, final_output.getvalue())

                progress_bar.progress((index + 1) / total_files)

        zip_buffer.seek(0)
        status_text.empty()
        progress_bar.empty()

        st.success(f"🎉 Đã xử lý xong! Các file được gom gọn vào thư mục **`{folder_in_zip}/`** trong file ZIP với đầy đủ 100% định dạng gốc.")

        st.download_button(
            label=f"📦 TẢI VỀ FILE ZIP ({zip_file_name})",
            data=zip_buffer,
            file_name=zip_file_name,
            mime="application/zip",
            use_container_width=True
        )
else:
    st.info("💡 Vui lòng tải lên file Word (.docx) và ít nhất 1 file Excel để bắt đầu.")
