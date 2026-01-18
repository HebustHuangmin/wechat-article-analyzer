"""
视频检测测试脚本
用于验证requests和Selenium分别能否检测到视频div
"""

import requests
from bs4 import BeautifulSoup

def test_url_simple(url):
    """
    简单测试：用requests获取HTML，查找视频相关关键词
    """
    print(f"\n{'='*60}")
    print(f"测试URL: {url}")
    print(f"{'='*60}")
    
    # 使用requests获取
    session = requests.Session()
    session.trust_env = False  # 禁用代理
    
    try:
        print("\n[1] 使用 requests 获取HTML...")
        response = session.get(
            url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'},
            timeout=15
        )
        html = response.text
        
        print(f"  ✓ HTML长度: {len(html):,} 字符")
        print(f"  ✓ 状态码: {response.status_code}")
        
        # 搜索关键字符串
        keywords = {
            'full_screen_opr': 'full_screen_opr' in html,
            'wx_video_play_opr': 'wx_video_play_opr' in html,
            'video_length': 'video_length' in html,
            'data-v-': 'data-v-' in html,
            '<video': '<video' in html,
            'iframe': 'iframe' in html,
        }
        
        print(f"\n[2] 关键词检测:")
        for keyword, found in keywords.items():
            status = "✓ 找到" if found else "✗ 未找到"
            print(f"  {status}: {keyword}")
        
        # 使用BeautifulSoup解析
        print(f"\n[3] BeautifulSoup解析:")
        soup = BeautifulSoup(html, 'lxml')
        
        # 查找所有div
        all_divs = soup.find_all('div')
        print(f"  总div数量: {len(all_divs)}")
        
        # 查找包含特定class的div
        video_divs_method1 = soup.find_all('div', class_='full_screen_opr wx_video_play_opr')
        print(f"  方法1 - soup.find_all('div', class_='full_screen_opr wx_video_play_opr'): {len(video_divs_method1)}")
        
        # 手动遍历查找
        video_divs_method2 = []
        for div in all_divs:
            classes = div.get('class', [])
            if 'full_screen_opr' in classes and 'wx_video_play_opr' in classes:
                video_divs_method2.append(div)
        print(f"  方法2 - 手动遍历检查class列表: {len(video_divs_method2)}")
        
        # 显示找到的视频div详情
        if video_divs_method2:
            print(f"\n[4] 找到的视频div:")
            for i, div in enumerate(video_divs_method2, 1):
                print(f"\n  --- 视频 {i} ---")
                print(f"  class: {div.get('class')}")
                video_span = div.find('span', class_='video_length')
                if video_span:
                    print(f"  时长: {video_span.get_text(strip=True)}")
                else:
                    print(f"  时长: 未找到 video_length span")
                # 打印完整div HTML（截断显示）
                div_html = str(div)
                if len(div_html) > 300:
                    print(f"  HTML: {div_html[:300]}...")
                else:
                    print(f"  HTML: {div_html}")
        else:
            print(f"\n[4] 未找到视频div")
            # 搜索包含关键class的div
            print(f"\n  搜索包含'full_screen_opr'的div:")
            for div in all_divs:
                classes = div.get('class', [])
                if 'full_screen_opr' in classes:
                    print(f"    找到: class={classes}")
                    break
            else:
                print(f"    未找到任何包含'full_screen_opr'的div")
        
    except Exception as e:
        print(f"\n  ✗ 错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    # 从excel_analyzer.py获取一个URL进行测试
    print("请提供一个您确定包含视频的微信文章URL进行测试")
    print("示例: https://mp.weixin.qq.com/s/xxxxx")
    print()
    
    # 可以直接在这里填入URL测试
    test_url = input("请输入URL (或直接回车使用高新发布.xlsx中的第一个链接): ").strip()
    
    if not test_url:
        # 从Excel读取第一个链接
        import pandas as pd
        df = pd.read_excel('高新发布.xlsx')
        if '备注（链接）' in df.columns and len(df) > 0:
            test_url = df['备注（链接）'].iloc[0]
            print(f"使用Excel中的第一个链接: {test_url}")
        else:
            print("无法从Excel读取链接，请手动输入")
            exit(1)
    
    test_url_simple(test_url)
