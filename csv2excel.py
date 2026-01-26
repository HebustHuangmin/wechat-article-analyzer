import pandas as pd
import os

# 读取KeyDept文件中的部门关键字
def load_department_keywords():
    if os.path.exists('KeyDept'):
        with open('KeyDept', 'r', encoding='utf-8') as f:
            keywords = [line.strip() for line in f if line.strip()]
        return keywords
    return []

def find_department(title, keywords):
    if not title or not keywords:
        return ''
    title_str = str(title)
    for keyword in keywords:
        if keyword and keyword in title_str:
            return keyword
    return ''

DEPARTMENT_KEYWORDS = load_department_keywords()

df = pd.read_csv('结果.csv', encoding='utf-8')
df['字数时长'] = df.apply(
	lambda r: f"文字{r['文字字数']}字" if r['视频时长'] == '无视频' else f"视频{r['视频时长']}\n文字{r['文字字数']}字",
	axis=1
)
df['部门'] = df.apply(
	lambda r: find_department(r['网页文章标题'], DEPARTMENT_KEYWORDS) if r['原发转载'] == '原发' else '',
	axis=1
)
# 链接处理：原发为空，转载保留链接
df['链接'] = df.apply(
	lambda r: '' if r['原发转载'] == '原发' else r.get('链接', ''),
	axis=1
)
df.to_excel('结果.xlsx', index=False, engine='openpyxl')
print('OK')
