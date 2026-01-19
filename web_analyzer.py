import requests
from bs4 import BeautifulSoup, Comment
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time

def analyze_webpage(url):
    """
    分析网页内容，提取视频、文字、来源等信息
    微信公众号文章使用Vue.js动态渲染，需要Selenium获取完整HTML
    """
    driver = None
    try:
        print(f"  开始分析: {url}")
        
        # 配置Selenium（微信文章是Vue.js动态渲染，必须用Selenium）
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--proxy-server="direct://"')
        chrome_options.add_argument('--proxy-bypass-list=*')
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        # 使用Selenium获取页面（视频div是JavaScript动态生成的）
        # 优先使用Selenium内置的driver管理，失败再回退到webdriver_manager
        try:
            driver = webdriver.Chrome(options=chrome_options)
        except Exception as e1:
            import os
            os.environ['WDM_LOG'] = '0'  # 禁用webdriver-manager日志
            try:
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=chrome_options)
            except Exception as e2:
                raise Exception(f"无法启动Chrome浏览器，请检查已安装的Chrome/ChromeDriver。原始错误: {e1} | {e2}")
        driver.set_page_load_timeout(20)
        driver.get(url)
        
        # 等待页面加载和Vue.js渲染完成
        print(f"  等待页面渲染...")
        time.sleep(4)  # 给足时间让Vue.js渲染视频div
        
        html_content = driver.page_source
        print(f"  ✓ 获取HTML成功，长度: {len(html_content):,} 字符")

        
        # 使用lxml解析器处理Vue.js属性
        soup = BeautifulSoup(html_content, 'lxml')
        
        # 提取视频信息 - 查找class同时包含"full_screen_opr"和"wx_video_play_opr"的div
        videos = []
        has_video = False
        
        all_divs = soup.find_all('div')
        video_divs = []
        
        for div in all_divs:
            classes = div.get('class', [])
            # 检查是否同时包含两个class
            if 'full_screen_opr' in classes and 'wx_video_play_opr' in classes:
                video_divs.append(div)
        
        if video_divs:
            print(f"  ✓ 检测到 {len(video_divs)} 个视频")
        else:
            print(f"  ✗ 未检测到视频")

        if video_divs:
            has_video = True
            for video_div in video_divs:
                video_info = {}
                # 查找class为"video_length"的span标签获取时长
                video_length_span = video_div.find('span', class_='video_length')
                if video_length_span:
                    duration_text = video_length_span.get_text(strip=True)
                    # 直接保存时长文本，不进行转换
                    video_info['duration'] = duration_text
                else:
                    video_info['duration'] = '无时长信息'
                
                videos.append(video_info)
        
        # 提取文字内容（获取所有文本内容）
        # 移除脚本和样式
        for script in soup(['script', 'style']):
            script.decompose()
        # 移除HTML注释
        for comment in soup.find_all(string=lambda text: isinstance(text, Comment)):
            comment.extract()

        # 获取全文文本（用于来源/摘要回退搜索）
        full_text = soup.get_text()
        # 去掉可能残留的JavaScript注释（多行 /*...*/ 和单行 //...）
        full_text = re.sub(r'/\*.*?\*/', '', full_text, flags=re.S)
        full_text = re.sub(r'//.*(?=\n)', '', full_text)
        full_text = re.sub(r'\s+', ' ', full_text).strip()

        # 优先查找 id 为 js_article 的 div，仅统计该 div 中的中文字符数
        article_div = soup.find('div', id='js_article')
        if article_div:
            # 从 article_div 中移除内部脚本/样式与注释，然后获取文本
            for s in article_div(['script', 'style']):
                s.decompose()
            for comment in article_div.find_all(string=lambda text: isinstance(text, Comment)):
                comment.extract()
            article_text = article_div.get_text()
            article_text = re.sub(r'/\*.*?\*/', '', article_text, flags=re.S)
            article_text = re.sub(r'//.*(?=\n)', '', article_text)
            article_text = re.sub(r'\s+', ' ', article_text).strip()
            chinese_chars = re.findall(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]', article_text)
            word_count = len(chinese_chars)
        else:
            # 回退到全文统计中文字符数
            chinese_chars = re.findall(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]', full_text)
            word_count = len(chinese_chars)
        
        # 提取来源信息：优先在文本中查找“来源：”或“来源:”并取冒号后的内容作为来源
        source = '无来源信息'
        # 在清理后的文本中查找“来源：”模式
        m = re.search(r'来源[:：]\s*([^\n\r]{1,200})', full_text)
        if m:
            candidate = m.group(1).strip()
            # 截断到常见结束符号
            candidate = re.split(r'[。；;，,\n\r]', candidate)[0].strip()
            if candidate:
                source = candidate
        else:
            # 回退到原有的标签查找方式
            source_elem = soup.find('span', class_='rich_media_meta rich_media_meta_text')
            if source_elem:
                source = source_elem.get_text(strip=True)
            else:
                source_elem = soup.find('div', class_='profile_info_inner')
                if source_elem:
                    source = source_elem.get_text(strip=True)
        
        # 提取来源后仅保留空格前的内容
        if source and isinstance(source, str):
            source = source.split()[0]

        # 提取摘要（获取文章描述或前段内容）
        summary = '无摘要'
        og_description = soup.find('meta', property='og:description')
        if og_description:
            summary = og_description.get('content', '无摘要')
        else:
            # 尝试获取文章描述
            desc_elem = soup.find('p', class_='js_profile_links')
            if desc_elem:
                summary = desc_elem.get_text(strip=True)
            else:
                # 默认使用开头文本
                summary = full_text[:200]
        
        
        return {
            'has_video': has_video,
            'videos': videos,
            'word_count': word_count,
            'source': source,
            'summary': summary,
            'is_text': word_count > 0
        }
    
    except Exception as e:
        print(f"  ✗ 错误: {str(e)}")
        return {
            'error': f'分析失败: {str(e)}',
            'has_video': False,
            'videos': [],
            'word_count': 0,
            'source': '无来源信息',
            'summary': '无摘要'
        }
    finally:
        # 确保关闭浏览器
        if driver:
            try:
                driver.quit()
            except:
                pass