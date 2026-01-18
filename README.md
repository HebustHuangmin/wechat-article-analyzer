# WeChat Article Analyzer

微信公众号文章内容分析系统 - 自动化提取、分析和处理WeChat公众号文章内容。

## 项目描述

本项目用于自动化分析石家庄高新区微信公众号发布的文章，提取以下信息：

- **视频检测**：使用Selenium识别Vue.js动态渲染的视频内容
- **中文字数统计**：仅统计中文字符，排除HTML标签
- **来源信息提取**：从文章内容中提取来源信息
- **内容分类**：判断文章类型（文字/视频/混合）
- **批量处理**：生成Excel结果和Word审核文档

## 功能特点

✓ 自动化文章分析与分类
✓ 视频检测和时长提取  
✓ 中文内容的精确字数统计
✓ 源信息与摘要提取
✓ 原发/转载自动分类
✓ 批量生成审核Word文档
✓ 完整的Excel数据导出

## 项目结构

```
WeChat/
├── excel_analyzer.py      # 主程序：读取Excel，分析文章，生成结果
├── web_analyzer.py        # 网页分析模块：Selenium爬取、内容解析
├── generate_docs.py       # Word文档生成模块：批量生成审核文档
├── create_excel.py        # Excel生成脚本：创建结果数据文件
├── 模板.docx              # Word模板文件
├── 高新发布.xlsx          # 输入：文章列表数据源
├── 结果.xlsx              # 输出：分析结果汇总
└── 送审签/                # 输出：生成的Word审核文档目录
```

## 核心模块

### excel_analyzer.py
主程序脚本，协调整个分析流程：
- 读取输入的高新发布.xlsx
- 筛选最近7天的文章
- 调用web_analyzer进行逐篇分析
- 输出结果到结果.xlsx
- 调用generate_docs生成Word文档

### web_analyzer.py
网页内容分析引擎（使用Selenium + BeautifulSoup）：

**特性**：
- **JavaScript渲染**：使用Selenium处理Vue.js动态渲染内容
- **视频检测**：查找class同时包含`full_screen_opr`和`wx_video_play_opr`的div元素
- **中文计数**：使用正则表达式仅统计Unicode CJK字符范围
- **信息提取**：提取来源、摘要等元数据

**技术细节**：
```
检测视频逻辑：
1. 获取所有div元素
2. 检查class属性同时包含两个标记
3. 提取video_length span中的时长信息

中文字数统计正则表达式：
[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]
覆盖CJK统一表意文字各个范围
```

### generate_docs.py
批量Word文档生成模块：
- 读取结果.xlsx数据
- 加载模板.docx并复制
- 填充每行的对应数据到表格
- 按日期_序号命名并保存

### create_excel.py
快速生成Excel结果文件的辅助脚本。

## 数据流程

```
高新发布.xlsx (输入)
    ↓
excel_analyzer.py (筛选最近7天)
    ↓
web_analyzer.py (逐篇分析)
    ├─ 视频检测
    ├─ 字数统计
    ├─ 来源提取
    └─ 摘要获取
    ↓
excel_analyzer.py (数据整合)
    ↓
结果.xlsx (输出数据)
    ↓
generate_docs.py (批量生成)
    ↓
送审签/*.docx (输出文档)
```

## 输出数据字段

| 字段 | 说明 |
|------|------|
| 序号 | 文章编号 |
| 日期 | 发布日期 |
| 网页文章标题 | WeChat文章标题 |
| 原发转载 | 根据来源字段自动判断 |
| 类型 | 文字/视频/文字，视频 |
| 文字字数 | 中文字符数 |
| 视频时长 | MM:SS格式 |
| 字数时长 | 综合显示 |
| 来源 | 文章来源信息 |
| 摘要 | 文章摘要 |
| 链接 | 微信公众号链接 |

## 技术栈

- **Python 3.14.0**
- **Selenium 4.x** - 浏览器自动化
- **BeautifulSoup4** - HTML解析
- **pandas** - 数据处理
- **openpyxl** - Excel操作
- **python-docx** - Word文档生成
- **webdriver-manager** - ChromeDriver自动管理

## 环境配置

### 依赖安装
```bash
pip install -r requirements.txt
```

### 虚拟环境
```bash
python -m venv .venv
.venv\Scripts\activate
```

## 使用方法

### 方式1：完整流程（推荐）
```bash
python excel_analyzer.py
```

此命令将：
1. 分析高新发布.xlsx中最近7天的文章
2. 生成结果.xlsx
3. 自动生成25个Word审核文档到送审签/目录

### 方式2：仅生成Excel
```bash
python create_excel.py
```

### 方式3：仅生成Word文档
```bash
python -c "from generate_docs import generate_word_docs; generate_word_docs('结果.xlsx')"
```

## 常见问题

### Q: 视频检测不准确？
A: 确保已安装最新的Selenium和ChromeDriver。脚本在获取HTML后会等待4秒让Vue.js完成渲染。

### Q: 字数统计包括英文吗？
A: 不包括。脚本使用正则表达式仅统计CJK字符（中文）。

### Q: 如何修改分析的日期范围？
A: 修改excel_analyzer.py第20行的timedelta(days=6)参数。

### Q: Word模板如何自定义？
A: 编辑模板.docx文件，保留表格结构即可。脚本会自动填充数据到表格单元格。

## 版本历史

### v1.0.0 (2026-01-18)
- ✓ 初始版本发布
- ✓ Selenium-based 视频检测
- ✓ 中文字符精确计数
- ✓ 批量Word文档生成
- ✓ 原发/转载自动分类

## 许可证

MIT License

## 作者

石家庄高新区内容管理团队

## 联系方式

如有问题或建议，请提交Issue或Pull Request。

---

**最后更新**: 2026年1月18日
