import pandas as pd
from datetime import datetime, timedelta

# 测试逻辑（不依赖web_analyzer）
# 创建模拟数据
data = {
    '发布时间': ['2026-01-18', '2026-01-17', '2026-01-16', '2025-01-11'],
    '标题': ['测试1', '测试2', '测试3', '旧数据'],
    '备注（链接）': ['http://example.com/1', 'http://example.com/2', 'http://example.com/3', 'http://example.com/4']
}

df = pd.DataFrame(data)
df['发布时间'] = pd.to_datetime(df['发布时间']).dt.date

# 假设分析结果
results = []
for idx, row in df.iterrows():
    # 模拟不同场景
    if idx == 0:
        word_count = 1000
        video_duration_text = '5:30'
    elif idx == 1:
        word_count = 500
        video_duration_text = '无视频'
    elif idx == 2:
        word_count = 1200
        video_duration_text = '10:00'
    else:
        word_count = 800
        video_duration_text = '无视频'
    
    # 构建字数时长列
    if video_duration_text == '无视频':
        char_duration_text = f"字数（{word_count}）"
    else:
        char_duration_text = f"字数（{word_count}）时长（{video_duration_text}）"
    
    results.append({
        '序号': idx + 1,
        '日期': str(row['发布时间']),
        '标题': row['标题'],
        '文字字数': word_count,
        '视频时长': video_duration_text,
        '字数时长': char_duration_text
    })

# 输出和保存
print("序号, 日期, 标题, 文字字数, 视频时长, 字数时长")
for res in results:
    print(f"{res['序号']}, {res['日期']}, {res['标题']}, {res['文字字数']}, {res['视频时长']}, {res['字数时长']}")

out_df = pd.DataFrame(results)
out_path = 'test_结果.xlsx'
try:
    out_df.to_excel(out_path, index=False)
    print(f"\n已将结果保存到: {out_path}")
except Exception as e:
    print(f"保存Excel失败: {e}")
