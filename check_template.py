from docx import Document

print("=== 检查封面模板 ===")
doc = Document('模板文件/template_报告封面.docx')
for i, p in enumerate(doc.paragraphs):
    if '标题' in p.text or '项目名称' in p.text or len(p.text.strip()) > 15:
        print(f"段落{i}: runs={len(p.runs)}, 长度={len(p.text.strip())}, text=\"{p.text[:50]}\"")
        if len(p.runs) <= 3:
            for j, r in enumerate(p.runs):
                print(f"  run{j}: text=\"{r.text}\"")

