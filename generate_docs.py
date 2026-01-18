import os
import pandas as pd
from docx import Document

IN_PATH = '结果.xlsx'
OUT_DIR = '送审签'
TEMPLATE = None
TEMPLATE_CANDIDATES = ['模板.docx', '模板.doc']


def make_doc_default(row):
    doc = Document()
    table = doc.add_table(rows=7, cols=2)
    table.style = 'Table Grid'

    labels = ['稿件标题', '原发/转载', '类别', '时长（字数）', '交稿时间', '拟发时间', '作者说明']
    values = [
        str(row.get('网页文章标题', '')),
        '原发' if (not row.get('来源') or str(row.get('来源')).strip() == '无来源信息') else '转载',
        str(row.get('类型', '')),
        '',
        str(row.get('日期', '')),
        str(row.get('日期', '')),
        str(row.get('摘要', '')),
    ]

    text_words = row.get('文字字数', '')
    video_len = row.get('视频时长', '')
    times = f"文字字数: {text_words}"
    if video_len and str(video_len).strip() and str(video_len).strip() != '无视频':
        times += f"；视频时长: {video_len}"
    values[3] = times

    for i, lab in enumerate(labels):
        table.cell(i, 0).text = lab
        table.cell(i, 1).text = values[i]

    return doc


def fill_table_from_template(doc, row):
    # 尝试在 doc 的第一个表格中根据左列标签填入右列
    if not doc.tables:
        return False
    tbl = doc.tables[0]
    for r in tbl.rows:
        left = r.cells[0].text.strip()
        if '稿件标题' in left:
            r.cells[1].text = str(row.get('网页文章标题', ''))
        elif '原发' in left or '转载' in left or '原发/转载' in left:
            r.cells[1].text = '原发' if (not row.get('来源') or str(row.get('来源')).strip() == '无来源信息') else '转载'
        elif '类别' in left:
            r.cells[1].text = str(row.get('类型', ''))
        elif '时长' in left or '字数' in left:
            text_words = row.get('文字字数', '')
            video_len = row.get('视频时长', '')
            val = f"文字字数: {text_words}"
            if video_len and str(video_len).strip() and str(video_len).strip() != '无视频':
                val += f"；视频时长: {video_len}"
            r.cells[1].text = val
        elif '交稿' in left:
            r.cells[1].text = str(row.get('日期', ''))
        elif '拟发' in left:
            r.cells[1].text = str(row.get('日期', ''))
        elif '作者' in left:
            r.cells[1].text = str(row.get('摘要', ''))
    return True


def generate_word_docs(excel_path):
    """从Excel结果文件生成Word文档"""
    if not os.path.exists(excel_path):
        print(f"找不到输入文件: {excel_path}")
        return
    df = pd.read_excel(excel_path, dtype=str)
    os.makedirs(OUT_DIR, exist_ok=True)

    # 尝试加载模板（优先模板.docx）
    template_doc = None
    for t in TEMPLATE_CANDIDATES:
        if os.path.exists(t):
            try:
                template_doc = Document(t)
                print('已加载模板：', t)
                break
            except Exception as e:
                print(f'无法以 docx 模式读取 {t}，尝试下一个模板。', e)

    for _, row in df.iterrows():
        date = row.get('日期', '').replace('/', '-').replace('\\', '-')
        seq = row.get('序号', '')
        fname = f"{date}_{seq}.docx" if date else f"{seq}.docx"
        fname = fname.replace(' ', '_')
        out_path = os.path.join(OUT_DIR, fname)

        if template_doc is not None:
            # 复制模板内容到新文档对象
            new_doc = Document()
            for element in template_doc.element.body:
                new_doc.element.body.append(element)
            ok = False
            try:
                ok = fill_table_from_template(new_doc, row)
            except Exception as e:
                print('模板填充出错，回退默认样式：', e)
            if not ok:
                new_doc = make_doc_default(row)
            doc = new_doc
        else:
            doc = make_doc_default(row)

        doc.save(out_path)
        print(f"写入: {out_path}")


def main():
    """当直接运行此脚本时调用"""
    generate_word_docs(IN_PATH)


if __name__ == '__main__':
    main()
