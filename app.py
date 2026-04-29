#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import uuid
import shutil
from datetime import timedelta
from flask import Flask, request, jsonify, render_template, send_file, abort
import fitz
import re
import csv
import xlwt

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)


def detect_layout(page_width, page_height):
    if page_width > 500 or page_height > 500:
        return "grid"
    return "single"


def is_fedex(sub_no_str):
    if not sub_no_str:
        return False
    first_sub = sub_no_str.split('/')[0].strip()
    return bool(re.match(r'^\d{4}\s+\d{4}\s+\d{4}$', first_sub))


def extract_master_no(text):
    nums = re.findall(r'J\d{12}|ALS\d{11}', text, re.IGNORECASE)
    return nums[0].upper() if nums else None


def generate_excel_files(info_list, output_dir):
    sub_tpl_path = os.path.join(output_dir, '批量修改子单模板.xls')
    master_tpl_path = os.path.join(output_dir, '批量修改渠道转单主单号.xls')

    wb1 = xlwt.Workbook(encoding='utf-8')
    ws1 = wb1.add_sheet('Sheet1')
    for col, h in enumerate(['客户单号', '原子单号', '新子单号', '原渠道转单号', '新渠道转单号', '页数']):
        ws1.write(0, col, h)

    row = 1
    for item in info_list:
        master_no = item['主单号']
        sub_no_str = item['子单号']
        pages = item['页数']

        if not sub_no_str:
            ws1.write(row, 0, master_no)
            if pages:
                ws1.write(row, 5, pages)
            row += 1
            continue

        sub_nos = [s.strip() for s in sub_no_str.split('/')]
        fedex = is_fedex(sub_no_str)

        for i, sn in enumerate(sub_nos):
            ws1.write(row, 0, master_no)
            if fedex:
                ws1.write(row, 4, sn.replace(' ', ''))
            else:
                ws1.write(row, 4, sn.replace('(', '').replace(')', ''))
            if i == 0 and pages:
                ws1.write(row, 5, pages)
            row += 1

    wb1.save(sub_tpl_path)

    wb2 = xlwt.Workbook(encoding='utf-8')
    ws2 = wb2.add_sheet('Sheet1')
    for col, h in enumerate(['客户单号', '原渠道转单主单号', '新渠道转单主单号']):
        ws2.write(0, col, h)

    row = 1
    for item in info_list:
        master_no = item['主单号']
        sub_no_str = item['子单号']
        shipment_no = item['单号(Shipment No)']
        fedex = is_fedex(sub_no_str)

        ws2.write(row, 1, master_no)

        if fedex:
            sub_nos = [s.strip() for s in sub_no_str.split('/')] if sub_no_str else []
            if sub_nos:
                ws2.write(row, 2, sub_nos[0].replace(' ', ''))
        else:
            if shipment_no:
                try:
                    ws2.write(row, 2, int(shipment_no))
                except ValueError:
                    ws2.write(row, 2, shipment_no)
        row += 1

    wb2.save(master_tpl_path)
    return sub_tpl_path, master_tpl_path


def process_pdf(input_pdf, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    temp_files = {}

    doc = fitz.open(input_pdf)
    total_pages = len(doc)

    if total_pages == 0:
        raise Exception("PDF文件为空")

    page = doc[0]
    media_box = page.mediabox
    page_width = media_box.width
    page_height = media_box.height

    layout = detect_layout(page_width, page_height)

    if layout == "grid":
        half_width = page_width / 2
        half_height = page_height / 2
        regions = [
            ("top_left", fitz.Rect(0, half_height, half_width, page_height)),
            ("top_right", fitz.Rect(half_width, half_height, page_width, page_height)),
            ("bottom_left", fitz.Rect(0, 0, half_width, half_height)),
            ("bottom_right", fitz.Rect(half_width, 0, page_width, half_height)),
        ]
    else:
        regions = [("full_page", fitz.Rect(0, 0, page_width, page_height))]

    for page_idx in range(total_pages):
        for region_name, rect in regions:
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=page_idx, to_page=page_idx)
            page_obj = new_doc[0]

            try:
                page_obj.set_cropbox(rect)
            except ValueError:
                page_obj.set_mediabox(rect)
                page_obj.set_cropbox(rect)

            text = page_obj.get_text()
            master_no = extract_master_no(text)

            if not master_no:
                new_doc.close()
                continue

            shipment_no = ""
            shipment_match = re.search(r'Shipment\s*[Nn]o\s*[:.]?\s*(\d+)', text)
            if not shipment_match:
                shipment_match = re.search(r'Shipment\s*[Nn]o[\s\S]*?(\d{6,10})', text)
            if shipment_match:
                shipment_no = shipment_match.group(1)

            barcode_match = re.search(r'\(\d{2}\)\d{18,}', text)
            barcode = barcode_match.group(0) if barcode_match else ""

            sub_no_match = re.search(r'(?<!Mstr#\s)\d{4}\s+\d{4}\s+\d{4}', text)
            sub_no = sub_no_match.group(0) if sub_no_match else ""

            page_num = 1
            page_total = 1
            page_match = re.search(r'(\d+)/(\d+)', text)
            if page_match:
                page_num = int(page_match.group(1))
                page_total = int(page_match.group(2))

            tmp_path = os.path.join(output_dir, f"_tmp_{master_no}_{page_idx}_{region_name}.pdf")
            new_doc.save(tmp_path, garbage=4, deflate=True)
            new_doc.close()

            if master_no not in temp_files:
                temp_files[master_no] = {
                    "pages": [],
                    "master_no": master_no,
                    "sub_nos_with_page": [],
                    "shipment_no": shipment_no,
                    "barcodes_with_page": [],
                    "barcodes": [],
                }

            temp_files[master_no]["pages"].append({"path": tmp_path, "page_num": page_num, "page_total": page_total})
            if sub_no:
                temp_files[master_no]["sub_nos_with_page"].append({"page_num": page_num, "sub_no": sub_no})
            if not temp_files[master_no]["shipment_no"] and shipment_no:
                temp_files[master_no]["shipment_no"] = shipment_no
            if barcode:
                temp_files[master_no]["barcodes_with_page"].append({"page_num": page_num, "barcode": barcode})
            if barcode and barcode not in temp_files[master_no]["barcodes"]:
                temp_files[master_no]["barcodes"].append(barcode)

    doc.close()

    info_list = []
    for idx, (master_no, data) in enumerate(sorted(temp_files.items()), 1):
        pages_info = data["pages"]
        sub_nos_with_page = data["sub_nos_with_page"]
        barcodes_with_page = data["barcodes_with_page"]

        if sub_nos_with_page:
            sorted_sub_nos = sorted(sub_nos_with_page, key=lambda x: x["page_num"])
            sub_no = "/".join([item["sub_no"] for item in sorted_sub_nos])
        elif barcodes_with_page:
            sorted_barcodes = sorted(barcodes_with_page, key=lambda x: x["page_num"])
            sub_no = "/".join([item["barcode"] for item in sorted_barcodes])
        else:
            sub_no = ""

        shipment_no = data["shipment_no"]
        pages = len(pages_info)

        final_path = os.path.join(output_dir, f"{master_no}.pdf")
        if pages == 1:
            shutil.move(pages_info[0]["path"], final_path)
        else:
            sorted_pages = sorted(pages_info, key=lambda x: x["page_num"])
            merged = fitz.open()
            for page_info in sorted_pages:
                src_doc = fitz.open(page_info["path"])
                merged.insert_pdf(src_doc)
                src_doc.close()
            merged.save(final_path, garbage=4, deflate=True)
            merged.close()
            for page_info in pages_info:
                os.remove(page_info["path"])

        remark = "多页包裹" if pages > 1 else ""
        info_list.append({
            "序号": idx,
            "文件名": os.path.basename(final_path),
            "主单号": master_no,
            "子单号": sub_no,
            "单号(Shipment No)": shipment_no,
            "页数": pages,
            "备注": remark,
        })

    csv_path = os.path.join(output_dir, "面单信息汇总.csv")
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=["序号", "文件名", "主单号", "子单号", "单号(Shipment No)", "页数", "备注"])
        writer.writeheader()
        writer.writerows(info_list)

    sub_tpl_path, master_tpl_path = generate_excel_files(info_list, output_dir)

    return info_list, csv_path, sub_tpl_path, master_tpl_path


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '没有文件'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '没有选择文件'}), 400

    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'success': False, 'error': '请选择PDF文件'}), 400

    task_id = str(uuid.uuid4())
    upload_dir = os.path.join(app.config['UPLOAD_FOLDER'], task_id)
    output_dir = os.path.join(app.config['OUTPUT_FOLDER'], task_id)
    os.makedirs(upload_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    input_path = os.path.join(upload_dir, file.filename)
    file.save(input_path)

    try:
        info_list, csv_path, sub_tpl_path, master_tpl_path = process_pdf(input_path, output_dir)

        result = {
            'success': True,
            'task_id': task_id,
            'file_count': len(info_list),
            'files': [item['文件名'] for item in info_list],
            'summary': info_list,
            'csv_filename': os.path.basename(csv_path),
            'sub_template_filename': os.path.basename(sub_tpl_path),
            'master_template_filename': os.path.basename(master_tpl_path)
        }

        return jsonify(result)

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)
        if os.path.exists(upload_dir):
            shutil.rmtree(upload_dir)


@app.route('/api/download/<task_id>/<filename>')
def download_file(task_id, filename):
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], task_id, filename)
    if not os.path.exists(file_path):
        abort(404)
    return send_file(file_path, as_attachment=True)


@app.route('/api/download_all/<task_id>')
def download_all(task_id):
    output_dir = os.path.join(app.config['OUTPUT_FOLDER'], task_id)
    if not os.path.exists(output_dir):
        abort(404)

    zip_path = os.path.join(app.config['OUTPUT_FOLDER'], f'{task_id}_all.zip')
    shutil.make_archive(zip_path.replace('.zip', ''), 'zip', output_dir)

    return send_file(zip_path, as_attachment=True, download_name='split_labels.zip')


if __name__ == '__main__':
    import socket
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    print(f"\n{'='*60}")
    print(f"DHL面单分割工具 - Web版")
    print(f"{'='*60}")
    print(f"\n本地访问: http://127.0.0.1:5000")
    print(f"局域网访问: http://{local_ip}:5000")
    print(f"\n手机/其他设备访问: http://{local_ip}:5000")
    print(f"\n按 Ctrl+C 停止服务")
    print(f"{'='*60}\n")
    app.run(host='0.0.0.0', port=5000, debug=False)