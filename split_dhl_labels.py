#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DHL面单自动分割 + 信息提取工具 v4
功能：
  1. 自动检测面单格式（单页单张或2×2合版）
  2. 按单号分割PDF并提取信息
  3. 支持单号格式：J+12位数字 或 ALS+11位数字
  4. 支持子单号格式：4位数字 4位数字 4位数字
  5. 生成CSV汇总表
适用：DHL CMR V.2.5 DOMESTIC格式
依赖：PyMuPDF (fitz)
安装：pip install PyMuPDF
"""

import fitz
import re
import os
import shutil
import sys
import csv
import xlwt


def detect_layout(page_width, page_height, crop_width=None, crop_height=None):
    """
    自动检测面单布局格式
    返回: "single" (单页单张) 或 "grid" (2×2合版)
    如果CropBox已经小于MediaBox的一半，说明是预裁剪格式，按单页处理
    """
    if crop_width is not None and crop_height is not None:
        if crop_width < page_width * 0.6 and crop_height < page_height * 0.6:
            return "single"
    if page_width > 500 or page_height > 500:
        return "grid"
    else:
        return "single"


def extract_master_no(text):
    """
    提取主单号
    支持格式：J+12位数字 或 ALS+11位数字
    """
    nums = re.findall(r'J\d{12}|ALS\d{11}', text, re.IGNORECASE)
    if nums:
        return nums[0].upper()
    
    return None


def is_fedex(sub_no_str):
    """
    判断是否为联邦面单
    联邦面单子单号格式：XXXX XXXX XXXX
    DHL面单子单号格式：(00)XXXXXXXXXXXXXX
    """
    if not sub_no_str:
        return False
    first_sub = sub_no_str.split('/')[0].strip()
    return bool(re.match(r'^\d{4}\s+\d{4}\s+\d{4}$', first_sub))


def generate_excel_files(info_list, output_dir):
    """
    生成Excel模板文件
    1. 批量修改子单模板.xls
    2. 批量修改渠道转单主单号.xls
    """
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


def split_labels_and_extract(input_pdf, output_dir=None):
    """
    分割面单并提取关键信息
    """
    if output_dir is None:
        base = os.path.splitext(input_pdf)[0]
        output_dir = base + "_split"

    os.makedirs(output_dir, exist_ok=True)

    temp_files = {}
    doc = fitz.open(input_pdf)
    total_pages = len(doc)

    if total_pages == 0:
        print("❌ PDF文件为空")
        return []

    page = doc[0]
    media_box = page.mediabox
    crop_box = page.cropbox
    page_width = media_box.width
    page_height = media_box.height
    crop_width = crop_box.width
    crop_height = crop_box.height
    
    layout = detect_layout(page_width, page_height, crop_width, crop_height)
    pre_cropped = crop_width < page_width * 0.6 and crop_height < page_height * 0.6
    print(f"📐 页面尺寸: {page_width:.1f} × {page_height:.1f}")
    if pre_cropped:
        print(f"📐 可见区域: {crop_width:.1f} × {crop_height:.1f} (预裁剪格式)")
    print(f"📋 检测到布局: {'2×2合版' if layout == 'grid' else '单页单张'}")

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

    print(f"📄 输入文件: {input_pdf}")
    print(f"📊 总页数: {total_pages}")
    print(f"📁 输出目录: {output_dir}")
    print("=" * 60)

    for page_idx in range(total_pages):
        print(f"\n--- 处理第 {page_idx + 1}/{total_pages} 页 ---")
        for region_name, rect in regions:
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=page_idx, to_page=page_idx)
            page = new_doc[0]

            if pre_cropped:
                text = page.get_text()
            else:
                try:
                    page.set_cropbox(rect)
                except ValueError:
                    page.set_mediabox(rect)
                    page.set_cropbox(rect)
                text = page.get_text()
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

            print(f"  [{region_name}] 主单号:{master_no} | 子单号:{sub_no} | 单号:{shipment_no}")

            page_num = 1
            page_total = 1
            page_match = re.search(r'(\d+)/(\d+)', text)
            if page_match:
                page_num = int(page_match.group(1))
                page_total = int(page_match.group(2))

            actual_cropbox = page.cropbox
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
            temp_files[master_no]["pages"].append({"path": tmp_path, "page_num": page_num, "page_total": page_total, "cropbox": actual_cropbox})
            # 收集所有子单号（保留所有，不去重），记录对应的页码
            if sub_no:
                temp_files[master_no]["sub_nos_with_page"].append({"page_num": page_num, "sub_no": sub_no})
            # DHL面单使用条形码作为子单号，同样记录页码
            if barcode:
                temp_files[master_no]["barcodes_with_page"].append({"page_num": page_num, "barcode": barcode})
            if not temp_files[master_no]["shipment_no"] and shipment_no:
                temp_files[master_no]["shipment_no"] = shipment_no
            if barcode and barcode not in temp_files[master_no]["barcodes"]:
                temp_files[master_no]["barcodes"].append(barcode)

    doc.close()

    info_list = []
    for idx, (master_no, data) in enumerate(sorted(temp_files.items()), 1):
        pages_info = data["pages"]
        sub_nos_with_page = data["sub_nos_with_page"]
        barcodes_with_page = data["barcodes_with_page"]
        
        # 优先使用子单号（联邦面单），按页码排序
        if sub_nos_with_page:
            sorted_sub_nos = sorted(sub_nos_with_page, key=lambda x: x["page_num"])
            sub_no = "/".join([item["sub_no"] for item in sorted_sub_nos])
        elif barcodes_with_page:
            # DHL面单使用条形码作为子单号，按页码排序
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
                page_count_before = len(merged)
                merged.insert_pdf(src_doc)
                src_doc.close()
                merged[page_count_before].set_cropbox(page_info["cropbox"])
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
        writer = csv.DictWriter(
            f,
            fieldnames=["序号", "文件名", "主单号", "子单号", "单号(Shipment No)", "页数", "备注"]
        )
        writer.writeheader()
        writer.writerows(info_list)

    print("\n正在生成Excel模板文件...")
    sub_tpl_path, master_tpl_path = generate_excel_files(info_list, output_dir)

    print("\n" + "=" * 100)
    print(f"{'序号':<6} {'主单号':<18} {'子单号':<20} {'单号(Shipment No)':<15} {'页数':<6} {'备注':<10} {'类型':<8}")
    print("-" * 100)
    for r in info_list:
        carrier = "联邦" if is_fedex(r.get('子单号', '')) else "DHL"
        print(f"{r['序号']:<6} {r['主单号']:<18} {r['子单号']:<20} {r['单号(Shipment No)']:<15} {r['页数']:<6} {r['备注']:<10} {carrier:<8}")
    print("=" * 120)

    print(f"\n✅ 完成！")
    print(f"   • 独立PDF文件: {len(info_list)} 个")
    print(f"   • CSV汇总表: {csv_path}")
    print(f"   • 批量修改子单模板: {sub_tpl_path}")
    print(f"   • 批量修改渠道转单主单号: {master_tpl_path}")

    fedex_count = sum(1 for item in info_list if is_fedex(item.get('子单号', '')))
    if fedex_count:
        print(f"\n📮 联邦面单: {fedex_count} 票（渠道转单主单号已使用子单号）")

    return info_list


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python split_dhl_labels.py <输入PDF路径> [输出目录]")
        print("示例: python split_dhl_labels.py 15票label.pdf")
        sys.exit(1)

    input_file = sys.argv[1]
    output_folder = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(input_file):
        print(f"❌ 文件不存在: {input_file}")
        sys.exit(1)

    split_labels_and_extract(input_file, output_folder)
