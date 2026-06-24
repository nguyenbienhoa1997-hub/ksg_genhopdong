import os
import uuid
import json
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file, flash
from openpyxl import load_workbook

import database as db
import generator as gen

app = Flask(__name__)
app.secret_key = "hd-secret-2024"

UPLOAD_DIR    = os.path.join(os.path.dirname(__file__), "uploads")
TEMPLATES_DIR = os.path.join(UPLOAD_DIR, "templates")
PDFS_DIR      = os.path.join(UPLOAD_DIR, "pdfs")

db.init_db()


# ── Contracts ────────────────────────────────────────────────────────────────

@app.route("/")
def contracts():
    search        = request.args.get("q", "")
    template_code = request.args.get("template", "")
    items         = db.get_contracts(search=search, template_code=template_code)
    all_templates = db.get_templates()
    return render_template("contracts.html",
                           contracts=items,
                           templates=all_templates,
                           search=search,
                           selected_template=template_code)


@app.route("/contracts/<int:id>")
def contract_detail(id):
    contract = db.get_contract(id)
    if not contract:
        return redirect(url_for("contracts"))
    row_data = json.loads(contract["row_data"]) if contract["row_data"] else {}
    return render_template("contract_detail.html", contract=contract, row_data=row_data)


@app.route("/contracts/<int:id>/delete", methods=["POST"])
def delete_contract(id):
    contract = db.get_contract(id)
    if contract and contract["pdf_path"] and os.path.exists(contract["pdf_path"]):
        os.remove(contract["pdf_path"])
    db.delete_contract(id)
    flash("Đã xóa hợp đồng", "success")
    return redirect(url_for("contracts"))


@app.route("/contracts/<int:id>/download")
def download_contract(id):
    contract = db.get_contract(id)
    if not contract or not contract["pdf_path"] or not os.path.exists(contract["pdf_path"]):
        flash("File PDF không tồn tại", "error")
        return redirect(url_for("contract_detail", id=id))
    name = contract["name"] or f"HopDong_{id}"
    return send_file(contract["pdf_path"], as_attachment=True,
                     download_name=f"{gen.safe_filename(name)}.pdf")


# ── Generate ─────────────────────────────────────────────────────────────────

@app.route("/generate")
def generate_page():
    templates = db.get_templates()
    return render_template("generate.html", templates=templates)


@app.route("/api/upload-excel", methods=["POST"])
def upload_excel():
    f = request.files.get("excel")
    if not f or not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Vui lòng chọn file Excel (.xlsx)"}), 400

    temp_path = os.path.join(UPLOAD_DIR, f"tmp_{uuid.uuid4().hex}.xlsx")
    f.save(temp_path)

    try:
        wb   = load_workbook(temp_path, read_only=True, data_only=True)
        ws   = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception as e:
        os.remove(temp_path)
        return jsonify({"error": f"Lỗi đọc file: {e}"}), 400

    if not rows:
        os.remove(temp_path)
        return jsonify({"error": "File Excel trống"}), 400

    headers = [str(h) if h is not None else f"Col{i}" for i, h in enumerate(rows[0])]
    preview = [
        [str(v) if v is not None else "" for v in row]
        for row in rows[1:6]
    ]
    var_names = {h: gen.to_var_name(h) for h in headers}

    return jsonify({
        "temp_path":  temp_path,
        "headers":    headers,
        "var_names":  var_names,
        "preview":    preview,
        "total_rows": len(rows) - 1,
    })


@app.route("/api/generate", methods=["POST"])
def do_generate():
    data      = request.get_json(silent=True) or {}
    temp_path = data.get("temp_path", "")

    # Basic path traversal guard
    if not temp_path.startswith(UPLOAD_DIR) or not os.path.exists(temp_path):
        return jsonify({"error": "Session hết hạn, vui lòng upload lại"}), 400

    try:
        wb   = load_workbook(temp_path, data_only=True)
        ws   = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception as e:
        return jsonify({"error": f"Lỗi đọc file: {e}"}), 400

    headers  = [str(h) if h is not None else f"Col{i}" for i, h in enumerate(rows[0])]
    all_tpls = list(db.get_templates())

    # Auto-detect name column — dùng "Tên khách hàng" nếu có, không thì đánh số
    NAME_COL_CANDIDATES = ["Tên khách hàng", "Ten khach hang", "TÊN KHÁCH HÀNG"]
    name_column = next((h for h in headers if h in NAME_COL_CANDIDATES), "")

    if not all_tpls:
        return jsonify({"error": "Chưa có mẫu hợp đồng nào. Vui lòng upload mẫu trước."}), 400

    batch_id  = uuid.uuid4().hex
    batch_dir = os.path.join(PDFS_DIR, batch_id)
    os.makedirs(batch_dir, exist_ok=True)

    success, errors = 0, []

    for row_idx, row in enumerate(rows[1:], start=2):
        # Skip completely empty rows
        if all(v is None or str(v).strip() == "" for v in row):
            continue

        row_data = {headers[i]: (str(v) if v is not None else "")
                    for i, v in enumerate(row) if i < len(headers)}

        # Contract name from chosen column
        raw_name = row_data.get(name_column, "").strip() if name_column else ""
        name     = gen.safe_filename(raw_name) if raw_name else f"HopDong_{row_idx - 1:04d}"

        # Gen PDF for EVERY template
        for tpl in all_tpls:
            tpl     = dict(tpl)
            pdf_name = f"{name}_{tpl['code']}.pdf"
            pdf_path = os.path.join(batch_dir, pdf_name)
            try:
                gen.generate_pdf(tpl["file_path"], row_data, pdf_path)
                db.add_contract(
                    name          = f"{name} - {tpl['name']}",
                    template_id   = tpl["id"],
                    template_name = tpl["name"],
                    template_code = tpl["code"],
                    pdf_path      = pdf_path,
                    row_data      = json.dumps(row_data, ensure_ascii=False),
                    batch_id      = batch_id,
                )
                success += 1
            except Exception as e:
                errors.append(f"Dòng {row_idx} / Mẫu {tpl['code']}: {e}")

    db.add_batch(batch_id, os.path.basename(temp_path), success)
    os.remove(temp_path)

    return jsonify({"success": True, "count": success, "errors": errors, "batch_id": batch_id})


# ── Templates (Word mẫu) ──────────────────────────────────────────────────────

@app.route("/templates")
def templates_page():
    items = db.get_templates()
    return render_template("templates.html", templates=items)


@app.route("/templates/upload", methods=["POST"])
def upload_template():
    name = request.form.get("name", "").strip()
    code = request.form.get("code", "").strip().upper()
    f    = request.files.get("file")

    if not name or not code or not f:
        flash("Vui lòng điền đầy đủ thông tin", "error")
        return redirect(url_for("templates_page"))

    if not f.filename.lower().endswith(".docx"):
        flash("Chỉ nhận file .docx", "error")
        return redirect(url_for("templates_page"))

    if db.get_template_by_code(code):
        flash(f'Mã "{code}" đã tồn tại', "error")
        return redirect(url_for("templates_page"))

    filename  = f"{code}_{uuid.uuid4().hex[:8]}.docx"
    file_path = os.path.join(TEMPLATES_DIR, filename)
    f.save(file_path)

    db.add_template(name=name, code=code, filename=f.filename, file_path=file_path)
    flash(f'Đã thêm mẫu "{name}" (mã: {code})', "success")
    return redirect(url_for("templates_page"))


@app.route("/templates/<int:id>/delete", methods=["POST"])
def delete_template(id):
    tpl = db.get_template(id)
    if tpl and os.path.exists(tpl["file_path"]):
        os.remove(tpl["file_path"])
    db.delete_template(id)
    flash("Đã xóa mẫu hợp đồng", "success")
    return redirect(url_for("templates_page"))


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import webbrowser
    webbrowser.open("http://localhost:5000")
    app.run(host="0.0.0.0", port=5000, debug=False)
