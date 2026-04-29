import fitz
import re
import os
import shutil
import sys
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.font import Font
import xlwt


def detect_layout(page_width, page_height):
    if page_width > 500 or page_height > 500:
        return "grid"
    else:
        return "single"

def is_fedex(sub_no_str):
    if not sub_no_str:
        return False
    first_sub = sub_no_str.split('/')[0].strip()
    return bool(re.match(r'^\d{4}\s+\d{4}\s+\d{4}$', first_sub))

def split_labels_and_extract(input_pdf, output_dir=None, progress_callback=None):
    if output_dir is None:
        base = os.path.splitext(input_pdf)[0]
        output_dir = base + "_split"

    os.makedirs(output_dir, exist_ok=True)

    temp_files = {}

    try:
        doc = fitz.open(input_pdf)
    except Exception as e:
        raise Exception(f"无法打开PDF文件: {str(e)}")
    
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

    total_regions = total_pages * len(regions)

    if progress_callback:
        progress_callback(0, f"正在处理 {input_pdf} ({total_pages} 页)")

    processed_count = 0

    for page_idx in range(total_pages):
        if progress_callback:
            progress_callback(int((page_idx / total_pages) * 50), f"正在处理第 {page_idx + 1}/{total_pages} 页")
        
        for region_name, rect in regions:
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=page_idx, to_page=page_idx)
            page = new_doc[0]

            try:
                page.set_cropbox(rect)
            except ValueError:
                page.set_mediabox(rect)
                page.set_cropbox(rect)

            text = page.get_text()

            nums = re.findall(r'J\d{12}|ALS\d{11}', text, re.IGNORECASE)
            if not nums:
                new_doc.close()
                processed_count += 1
                continue

            master_no = nums[0]

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

            page_num = 1
            page_total = 1
            page_match = re.search(r'(\d+)/(\d+)', text)
            if page_match:
                page_num = int(page_match.group(1))
                page_total = int(page_match.group(2))

            temp_files[master_no]["pages"].append({"path": tmp_path, "page_num": page_num, "page_total": page_total, "cropbox": rect})
            if sub_no:
                temp_files[master_no]["sub_nos_with_page"].append({"page_num": page_num, "sub_no": sub_no})
            if not temp_files[master_no]["shipment_no"] and shipment_no:
                temp_files[master_no]["shipment_no"] = shipment_no
            if barcode:
                temp_files[master_no]["barcodes_with_page"].append({"page_num": page_num, "barcode": barcode})
            if barcode and barcode not in temp_files[master_no]["barcodes"]:
                temp_files[master_no]["barcodes"].append(barcode)

            processed_count += 1

    doc.close()

    if progress_callback:
        progress_callback(50, "正在合并面单...")

    info_list = []
    total_items = len(temp_files)
    for idx, (master_no, data) in enumerate(sorted(temp_files.items()), 1):
        if progress_callback:
            progress_callback(50 + int((idx / total_items) * 40), f"合并第 {idx}/{total_items} 个面单")
        
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

    if progress_callback:
        progress_callback(90, "正在生成CSV汇总表...")

    csv_path = os.path.join(output_dir, "面单信息汇总.csv")
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["序号", "文件名", "主单号", "子单号", "单号(Shipment No)", "页数", "备注"]
        )
        writer.writeheader()
        writer.writerows(info_list)

    if progress_callback:
        progress_callback(95, "正在生成Excel模板文件...")

    sub_tpl_path, master_tpl_path = generate_excel_files(info_list, output_dir)

    if progress_callback:
        progress_callback(100, "处理完成")

    return info_list, output_dir, csv_path, sub_tpl_path, master_tpl_path


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


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("DHL面单分割工具")
        self.root.geometry("600x500")
        self.root.resizable(False, False)

        try:
            self.root.iconbitmap(default=None)
        except:
            pass

        self.input_file = ""
        self.processing = False

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        title_font = Font(family="Microsoft YaHei", size=14, weight="bold")
        title_label = ttk.Label(main_frame, text="DHL面单自动分割工具", font=title_font)
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        ttk.Label(main_frame, text="选择PDF文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.file_entry = ttk.Entry(main_frame, width=50, state="readonly")
        self.file_entry.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=(0, 5))
        
        self.browse_btn = ttk.Button(main_frame, text="浏览...", command=self.browse_file)
        self.browse_btn.grid(row=2, column=2, pady=5, padx=(5, 0))

        self.process_btn = ttk.Button(main_frame, text="开始处理", command=self.start_process, state="disabled")
        self.process_btn.grid(row=3, column=0, columnspan=3, pady=20)

        self.progress_label = ttk.Label(main_frame, text="")
        self.progress_label.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=540, mode="determinate")
        self.progress_bar.grid(row=5, column=0, columnspan=3, pady=5)
        self.progress_bar["value"] = 0

        self.result_frame = ttk.LabelFrame(main_frame, text="处理结果")
        self.result_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        self.result_frame.grid_remove()

        self.result_text = tk.Text(self.result_frame, width=60, height=10, wrap=tk.WORD, state="disabled")
        self.result_text.grid(row=0, column=0, padx=10, pady=10)

        self.open_folder_btn = ttk.Button(main_frame, text="打开输出文件夹", command=self.open_output_folder, state="disabled")
        self.open_folder_btn.grid(row=7, column=0, columnspan=3, pady=10)

        main_frame.columnconfigure(0, weight=1)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if file_path:
            if not file_path.lower().endswith('.pdf'):
                messagebox.showwarning("警告", "请选择PDF格式的文件")
                return
            
            if not os.path.exists(file_path):
                messagebox.showerror("错误", "文件不存在")
                return
            
            try:
                with open(file_path, 'rb') as f:
                    header = f.read(4)
                    if header != b'%PDF':
                        messagebox.showwarning("警告", "文件不是有效的PDF格式")
                        return
            except Exception as e:
                messagebox.showerror("错误", f"无法读取文件: {str(e)}")
                return
            
            self.input_file = file_path
            self.file_entry.config(state="normal")
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.file_entry.config(state="readonly")
            self.process_btn.config(state="normal")
            self.result_frame.grid_remove()
            self.open_folder_btn.config(state="disabled")

    def start_process(self):
        if not self.input_file or self.processing:
            return
        
        self.processing = True
        self.process_btn.config(state="disabled")
        self.browse_btn.config(state="disabled")
        self.progress_bar["value"] = 0
        self.progress_label.config(text="准备处理...")
        self.result_frame.grid_remove()
        self.open_folder_btn.config(state="disabled")

        try:
            info_list, output_dir, csv_path, sub_tpl_path, master_tpl_path = split_labels_and_extract(
                self.input_file,
                progress_callback=self.update_progress
            )
            
            self.show_result(info_list, output_dir, sub_tpl_path, master_tpl_path)
            
        except Exception as e:
            messagebox.showerror("处理失败", f"发生错误: {str(e)}")
            self.progress_label.config(text="处理失败")
        finally:
            self.processing = False
            self.process_btn.config(state="normal")
            self.browse_btn.config(state="normal")

    def update_progress(self, value, message):
        self.progress_bar["value"] = value
        self.progress_label.config(text=message)
        self.root.update_idletasks()

    def show_result(self, info_list, output_dir, sub_tpl_path, master_tpl_path):
        self.result_frame.grid()
        self.open_folder_btn.config(state="normal")
        self.output_dir = output_dir
        
        self.result_text.config(state="normal")
        self.result_text.delete(1.0, tk.END)
        
        result_str = f"✅ 处理完成！\n\n"
        result_str += f"📄 输入文件: {os.path.basename(self.input_file)}\n"
        result_str += f"📁 输出目录: {output_dir}\n\n"
        result_str += f"📑 生成文件:\n"
        result_str += f"  • 独立PDF文件: {len(info_list)} 个\n"
        result_str += f"  • CSV汇总表: 面单信息汇总.csv\n"
        result_str += f"  • 批量修改子单模板: {os.path.basename(sub_tpl_path)}\n"
        result_str += f"  • 批量修改渠道转单主单号: {os.path.basename(master_tpl_path)}\n\n"
        
        fedex_count = sum(1 for item in info_list if is_fedex(item.get('子单号', '')))
        if fedex_count:
            result_str += f"📮 联邦面单: {fedex_count} 票（渠道转单主单号已使用子单号）\n\n"
        
        if info_list:
            result_str += "📊 面单列表:\n"
            for item in info_list[:5]:
                remark = f" ({item['备注']})" if item['备注'] else ""
                carrier = " [联邦]" if is_fedex(item.get('子单号', '')) else ""
                result_str += f"  - {item['主单号']}{carrier}{remark}\n"
            if len(info_list) > 5:
                result_str += f"  ... 还有 {len(info_list) - 5} 个面单\n"
        
        self.result_text.insert(tk.END, result_str)
        self.result_text.config(state="disabled")

    def open_output_folder(self):
        if hasattr(self, 'output_dir') and os.path.exists(self.output_dir):
            if sys.platform.startswith('win'):
                os.startfile(self.output_dir)
            else:
                import subprocess
                subprocess.run(['open', self.output_dir])


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
