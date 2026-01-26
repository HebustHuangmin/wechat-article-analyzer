import os
import sys
import pandas as pd
# 优先使用系统环境，缺少依赖时回退到本地 .venv site-packages
try:
    from web_analyzer import analyze_webpage
except ModuleNotFoundError:
    venv_site = os.path.join(os.path.dirname(__file__), '.venv', 'Lib', 'site-packages')
    if os.path.isdir(venv_site) and venv_site not in sys.path:
        sys.path.append(venv_site)
    from web_analyzer import analyze_webpage
from datetime import datetime, timedelta

# 读取KeyDept文件中的部门关键字
def load_department_keywords():
    """从KeyDept文件中读取部门名称作为关键字"""
    if os.path.exists('KeyDept'):
        with open('KeyDept', 'r', encoding='utf-8') as f:
            keywords = [line.strip() for line in f if line.strip()]
        return keywords
    return []

def find_department(title, keywords):
    """在标题中查找部门关键字，返回第一个匹配的部门名称"""
    if not title or not keywords:
        return ''
    title_str = str(title)
    for keyword in keywords:
        if keyword and keyword in title_str:
            return keyword
    return ''

DEPARTMENT_KEYWORDS = load_department_keywords()

# 读取Excel文件
df = pd.read_excel('高新发布.xlsx')

# 假设列名：'发布时间', '标题', '备注（链接）'
# 筛选最近一周（含今天）的行
today = datetime.now().date()
start_date = today - timedelta(days=6)
df['发布时间'] = pd.to_datetime(df['发布时间'], errors='coerce').dt.date
filtered_df = df[(df['发布时间'] >= start_date) & (df['发布时间'] <= today)]
print(f"匹配到近期记录条数: {len(filtered_df)}")

# 处理每个匹配的行
results = []
total = len(filtered_df)
for seq_idx, (idx, row) in enumerate(filtered_df.iterrows(), start=1):
    link = row['备注（链接）']
    print(f"[{seq_idx}/{total}] 分析: {link}")
    analysis = analyze_webpage(link)
    if 'error' in analysis:
        print(f"错误处理链接 {link}: {analysis['error']}")
        analysis = {'has_video': False, 'videos': [], 'word_count': 0, 'source': '无来源信息', 'summary': '无摘要', 'is_text': True}
    
    title = row['标题']  # 从Excel读取标题
    has_video = analysis['has_video']
    word_count = analysis['word_count']
    videos = analysis['videos']
    source = analysis.get('source', '无来源信息')
    summary = analysis.get('summary', '无摘要')
    # 精炼摘要
    if not analysis.get('is_text', True) and analysis.get('has_video', False):
        summary = f"视频内容：{title}"
    else:
        summary = summary[:50] + '...' if len(summary) > 50 else summary
    
    # 判断内容类型
    if has_video and word_count > 0:
        content_type = '文字，视频'
    elif has_video:
        content_type = '视频'
    else:
        content_type = '文字'
    
    # 判断原发/转载
    if source == '无来源信息':
        origin_type = '原发'
    else:
        origin_type = '转载'
    
    # 获取视频时长文本
    video_duration_text = '无视频'
    if videos:
        # 如果有多个视频，用逗号分隔时长
        durations = [v.get('duration', '无时长信息') for v in videos if v.get('duration')]
        if durations:
            video_duration_text = '、'.join(durations)
    
    # 构建字数时长列：有视频则两行显示，无视频只显示文字
    if video_duration_text == '无视频':
        char_duration_text = f"文字{word_count}字"
    else:
        char_duration_text = f"视频{video_duration_text}\n文字{word_count}字"
    
    # 查找部门：仅当原发时才查找关键字
    department = ''
    if origin_type == '原发':
        department = find_department(title, DEPARTMENT_KEYWORDS)
    
    # 链接处理：原发为空，转载保留链接
    article_link = '' if origin_type == '原发' else link
    
    results.append({
        '序号': seq_idx,
        '日期': str(row['发布时间']),
        '网页文章标题': title,
        '原发转载': origin_type,
        '类型': content_type,
        '文字字数': word_count,
        '视频时长': video_duration_text,
        '字数时长': char_duration_text,
        '来源': source,
        '摘要': summary,
        '部门': department,
        '链接': article_link
    })

# 输出结果
print("序号, 日期, 网页文章标题, 类型, 文字字数, 视频时长, 字数时长, 来源, 摘要")
for res in results:
    print(f"{res['序号']}, {res['日期']}, {res['网页文章标题']}, {res['类型']}, {res['文字字数']}, {res['视频时长']}, {res['字数时长']}, {res['来源']}, {res['摘要']}")

# 将结果保存到Excel和CSV文件
if results:
    out_df = pd.DataFrame(results)
    
    # 保存Excel文件
    out_xlsx = '结果.xlsx'
    try:
        out_df.to_excel(out_xlsx, index=False)
        print(f"✓ 已将结果保存到: {out_xlsx}")
    except Exception as e:
        print(f"✗ 保存Excel失败: {e}")
    
    # 保存CSV文件（用于Word文档生成）
    out_csv = '结果.csv'
    try:
        out_df.to_csv(out_csv, index=False, encoding='utf-8-sig')
        print(f"✓ 已将结果保存到: {out_csv}")
    except Exception as e:
        print(f"✗ 保存CSV失败: {e}")
    
    # 调用Word生成脚本
    print("\n开始生成Word文档...")
    try:
        from generate_docs import generate_word_docs
        generate_word_docs()  # 使用CSV作为数据源
        print("✓ Word文档生成完成!")
    except Exception as e:
        print(f"✗ 生成Word文档时出错: {e}")