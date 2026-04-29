import fitz
import os

dir_v2 = r'f:\项目\OX面单处理\6票label_v2'
if not os.path.exists(dir_v2):
    print('6票label_v2 not found')
else:
    for fname in sorted(os.listdir(dir_v2)):
        if not fname.endswith('.pdf'):
            continue
        fpath = os.path.join(dir_v2, fname)
        doc = fitz.open(fpath)
        print(f'=== {fname} ===')
        print(f'  Pages: {len(doc)}')
        for i in range(len(doc)):
            p = doc[i]
            mb = p.mediabox
            cb = p.cropbox
            text_lines = [l.strip() for l in p.get_text().split('\n') if l.strip()]
            first_line = text_lines[0] if text_lines else "(empty)"
            print(f'  Page {i+1}: MediaBox={mb}, CropBox={cb}')
            print(f'    Visible: {cb.width:.1f} x {cb.height:.1f}')
            print(f'    First line: {first_line}')
        doc.close()
