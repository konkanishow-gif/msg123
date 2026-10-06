import os
import sys
import json
import tempfile
from flask import Flask, render_template, request, jsonify, send_file

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from pdf_parser import extract_pdf_data, load_master_suppliers
from excel_generator import generate_summary_excel
from supplier_pdf_generator import generate_all_supplier_pdfs_zip
import cloudinary_handler
import history_manager

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static')
)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/suppliers', methods=['GET'])
def get_suppliers():
    suppliers = load_master_suppliers()
    return jsonify({"success": True, "suppliers": suppliers})

@app.route('/api/parse', methods=['POST'])
def parse_pdf():
    if 'pdf_file' not in request.files:
        return jsonify({"success": False, "error": "No PDF file uploaded"}), 400
    
    file = request.files['pdf_file']
    if file.filename == '':
        return jsonify({"success": False, "error": "No selected file"}), 400

    if not file.filename.lower().endswith('.pdf'):
        return jsonify({"success": False, "error": "Uploaded file must be a PDF"}), 400

    temp_dir = tempfile.gettempdir()
    temp_pdf_path = os.path.join(temp_dir, file.filename)
    file.save(temp_pdf_path)

    try:
        extracted = extract_pdf_data(temp_pdf_path)
        
        # Upload original PDF to Cloudinary if configured
        pdf_url = None
        if cloudinary_handler.is_cloudinary_configured():
            try:
                pdf_url = cloudinary_handler.upload_file_to_cloudinary(
                    temp_pdf_path,
                    folder="raw_pdfs",
                    filename=f"SO_{file.filename}"
                )
            except Exception as upload_err:
                print(f"Cloudinary PDF upload error: {upload_err}")

        # Auto-save snapshot into history
        file_urls = {}
        if pdf_url:
            file_urls["pdf_url"] = pdf_url

        snapshot_record = history_manager.save_sheet_snapshot(extracted, file_urls=file_urls)

        return jsonify({
            "success": True,
            "data": extracted,
            "snapshot_id": snapshot_record.get("id"),
            "pdf_url": pdf_url
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        if os.path.exists(temp_pdf_path):
            try:
                os.remove(temp_pdf_path)
            except Exception:
                pass

@app.route('/api/history', methods=['GET'])
def get_history():
    query = request.args.get('q', '')
    sheets = history_manager.get_all_sheets(search_query=query)
    return jsonify({
        "success": True,
        "sheets": sheets,
        "cloudinary_configured": cloudinary_handler.is_cloudinary_configured()
    })

@app.route('/api/history/<sheet_id>', methods=['GET'])
def get_history_detail(sheet_id):
    sheet = history_manager.get_sheet_by_id(sheet_id)
    if not sheet:
        return jsonify({"success": False, "error": "Sheet record not found"}), 404
    return jsonify({"success": True, "sheet": sheet})

@app.route('/api/history/<sheet_id>', methods=['DELETE'])
def delete_history_item(sheet_id):
    success = history_manager.delete_sheet_snapshot(sheet_id)
    return jsonify({"success": success})

@app.route('/api/save-worksheet', methods=['POST'])
def save_worksheet():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No JSON payload provided"}), 400

        msg_ref = data.get("msg_ref", "SO").replace(" ", "_")
        po_num = data.get("po_number", "").replace(" ", "_")
        filter_machined_only = bool(data.get("filter_machined_only", True))

        file_urls = {}

        # If Cloudinary is configured, generate and upload Excel files & Supplier PDFs ZIP
        if cloudinary_handler.is_cloudinary_configured():
            try:
                # 1. Client Excel
                client_excel_stream = generate_summary_excel(data, is_internal=False, filter_machined_only=filter_machined_only)
                client_filename = f"Client_Summary_{msg_ref}_{po_num}.xlsx"
                c_url = cloudinary_handler.upload_file_to_cloudinary(client_excel_stream, folder="client_excels", filename=client_filename)
                if c_url:
                    file_urls["client_excel_url"] = c_url

                # 2. Internal Stores Excel
                stores_excel_stream = generate_summary_excel(data, is_internal=True, filter_machined_only=filter_machined_only)
                stores_filename = f"Stores_Summary_{msg_ref}_{po_num}.xlsx"
                s_url = cloudinary_handler.upload_file_to_cloudinary(stores_excel_stream, folder="stores_excels", filename=stores_filename)
                if s_url:
                    file_urls["stores_excel_url"] = s_url

                # 3. Supplier PDFs ZIP Package
                zip_stream = generate_all_supplier_pdfs_zip(data)
                zip_filename = f"Supplier_PO_Package_{msg_ref}.zip"
                z_url = cloudinary_handler.upload_file_to_cloudinary(zip_stream, folder="supplier_packages", filename=zip_filename)
                if z_url:
                    file_urls["supplier_pdf_zip_url"] = z_url
            except Exception as c_err:
                print(f"Cloudinary upload error during worksheet save: {c_err}")

        record = history_manager.save_sheet_snapshot(data, file_urls=file_urls)

        return jsonify({
            "success": True,
            "record": record,
            "cloudinary_configured": cloudinary_handler.is_cloudinary_configured()
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/cloudinary-config', methods=['GET', 'POST'])
def handle_cloudinary_config():
    if request.method == 'GET':
        config = cloudinary_handler.load_cloudinary_config()
        return jsonify({
            "success": True,
            "configured": cloudinary_handler.is_cloudinary_configured(),
            "cloud_name": config.get("cloud_name", ""),
            "has_api_key": bool(config.get("api_key")),
            "has_api_secret": bool(config.get("api_secret"))
        })
    else:
        body = request.json or {}
        cloud_name = body.get("cloud_name", "")
        api_key = body.get("api_key", "")
        api_secret = body.get("api_secret", "")

        success = cloudinary_handler.save_cloudinary_config(cloud_name, api_key, api_secret)
        return jsonify({
            "success": success,
            "configured": cloudinary_handler.is_cloudinary_configured()
        })

@app.route('/api/download', methods=['POST'])
def download_excel():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No JSON payload provided"}), 400

        msg_ref = data.get("msg_ref", "SUMMARY").replace(" ", "_")
        po_num = data.get("po_number", "").replace(" ", "_")
        is_internal = bool(data.get("is_internal", False))
        filter_machined_only = bool(data.get("filter_machined_only", True))
        
        prefix = "Internal_Stores_Summary" if is_internal else "Client_Summary"
        filename = f"{prefix}_{msg_ref}_{po_num}.xlsx" if po_num else f"{prefix}_{msg_ref}.xlsx"
        
        stream = generate_summary_excel(data, is_internal=is_internal, filter_machined_only=filter_machined_only)
        
        # Save snapshot asynchronously / in background
        try:
            history_manager.save_sheet_snapshot(data)
        except Exception:
            pass

        return send_file(
            stream,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/export-supplier-pdfs', methods=['POST'])
def export_supplier_pdfs():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No JSON payload provided"}), 400

        msg_ref = data.get("msg_ref", "SO").replace(" ", "_")
        zip_filename = f"Supplier_PO_Package_{msg_ref}.zip"
        
        zip_stream = generate_all_supplier_pdfs_zip(data)
        
        return send_file(
            zip_stream,
            as_attachment=True,
            download_name=zip_filename,
            mimetype="application/zip"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)
