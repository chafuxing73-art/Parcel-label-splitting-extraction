import fitz
import os
import shutil
import re

src = r'f:\项目\OX面单处理\6票label.pdf'
test_dir = r'f:\项目\OX面单处理\test_exact_flow'
if os.path.exists(test_dir):
    shutil.rmtree(test_dir)
os.makedirs(test_dir)

doc = fitz.open(src)
total_pages = len(doc)

page = doc[0]
media_box = page.mediabox
crop_box = page.cropbox
page_width = media_box.width
page_height = media_box.height
crop_width = crop_box.width
crop_height = crop_box.height

pre_cropped = crop_width < page_width * 0.6 and crop_height < page_height * 0.6
regions = [("full_page", fitz.Rect(0, 0, page_width, page_height))]

print(f"pre_cropped={pre_cropped}")
print(f"regions={regions}")

temp_files = {}

for page_idx in range(total_pages):
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

        nums = re.findall(r'J\d{12}|ALS\d{11}', text, re.IGNORECASE)
        if not nums:
            new_doc.close()
            continue

        master_no = nums[0].upper()

        page_num = 1
        page_total = 1
        page_match = re.search(r'(\d+)/(\d+)', text)
        if page_match:
            page_num = int(page_match.group(1))
            page_total = int(page_match.group(2))

        tmp_path = os.path.join(test_dir, f"_tmp_{master_no}_{page_idx}_{region_name}.pdf")
        
        cb_before_save = new_doc[0].cropbox
        print(f"Page {page_idx+1}: master={master_no}, page_num={page_num}/{page_total}, CropBox before save={cb_before_save}")
        
        new_doc.save(tmp_path, garbage=4, deflate=True)
        new_doc.close()
        
        verify = fitz.open(tmp_path)
        cb_after_save = verify[0].cropbox
        print(f"  After save: CropBox={cb_after_save}")
        verify.close()

        if master_no not in temp_files:
            temp_files[master_no] = {"pages": [], "master_no": master_no}
        temp_files[master_no]["pages"].append({"path": tmp_path, "page_num": page_num})

doc.close()

print("\n=== Merging multi-page labels ===")
for master_no, data in sorted(temp_files.items()):
    pages_info = data["pages"]
    if len(pages_info) <= 1:
        continue
    
    print(f"\nMerging {master_no} ({len(pages_info)} pages):")
    sorted_pages = sorted(pages_info, key=lambda x: x["page_num"])
    merged = fitz.open()
    for page_info in sorted_pages:
        src_doc = fitz.open(page_info["path"])
        src_cb = src_doc[0].cropbox
        print(f"  Inserting page {page_info['page_num']}: CropBox={src_cb}")
        merged.insert_pdf(src_doc)
        src_doc.close()
    
    final_path = os.path.join(test_dir, f"{master_no}.pdf")
    merged.save(final_path, garbage=4, deflate=True)
    merged.close()
    
    verify = fitz.open(final_path)
    print(f"  After merge & save:")
    for i in range(len(verify)):
        cb = verify[i].cropbox
        text_lines = [l.strip() for l in verify[i].get_text().split('\n') if l.strip()]
        first_line = text_lines[0] if text_lines else "(empty)"
        print(f"    Page {i+1}: CropBox={cb}, first_line={first_line}")
    verify.close()
