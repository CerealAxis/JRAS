# JRAS
~~苦逼大学生的python作业~~

基于**爬虫**与**AI技术**的京东商品评论自动化分析系统

****

## 项目概述

JRAS是一款专为京东商品评论深度挖掘而设计的自动化分析工具，运用爬虫技术和自然语言处理（NLP）来收集、清洗、分析用户评论，并通过可视化手段展示关键洞察。该系统能够高效地抓取商品评论数据，执行文本预处理、情感分析及关键词提取等工作。

### 项目结构

```plaintext
JRAS/
├── clean.py         # 数据清洗脚本
├── main.py          # 主程序入口
├── nlp.py           # 自然语言处理模块
├── qianwen.py       # 配置与调用通义千问API模块
├── read_excel.py    # Excel读写辅助脚本
├── segmented.py     # 文本分词处理脚本
└── spider.py        # 商品评论爬虫模块

├── data
│   ├── bg.png        # 可能用作报告背景图片
│   └── stopwords.txt # 停用词表

└── output
    ├── cleaned_data.xlsx  # 清理后整理的评论数据集
    ├── contents.txt      # 摘要和总结的评论内容
    ├── original_data.xlsx # 爬取得到的原始评论数据
    ├── report.docx       # 自动生成的综合分析报告
    ├── segmented_words.txt # 分词后的词汇列表
    ├── sentiment_analysis_pie_chart.png # 情感分析结果饼状图
    └── wordcloud.png        # 关键词词云图
```

### 系统需求

- **操作系统**：Windows 7或更高版本（支持64位架构）
- **处理器**：至少1 GHz或速度更快的64位CPU或SoC
- **内存**：至少4GB RAM
- **Python环境**：Python 3.11或更新版本

### 安装指南

#### Pandoc安装

为了生成Word文档报告，请先安装Pandoc，这是一个强大的文档格式转换工具。访问[Pandoc官网安装指南](https://pandoc.org/installing.html)以获取适合您操作系统的安装方法。

### 使用步骤

1. **克隆项目**：从GitHub或其他托管平台克隆本项目至本地计算机，然后使用PyCharm或其他兼容IDE打开该项目。

2. **安装依赖**：在项目根目录下执行以下命令安装所有必需的第三方库：
   ```shell
   pip install -r requirements.txt
   ```

3. **配置API密钥**：在`qianwen.py`文件中填写您的通义千问API密钥：
   ```python
   dashscope.api_key = "在此处输入你的API密钥"
   ```
4. **运行分析**：启动主程序`main.py`，根据提示输入待分析的京东商品链接，点击**开始分析**按钮，等待程序执行完毕。

![软件页面](https://i.072333.xyz/file/9ad0fa06251562efa158b.png)

5. **查看结果**：分析完成后，进入`output`目录，可找到包含如下内容的各类输出文件：

   - `cleaned_data.xlsx`：已清洗和整理过的评论数据
   - `contents.txt`：提炼总结的评论主要内容
   - `original_data.xlsx`：爬取得到的原始评论记录
   - `report.docx`：系统自动生成的分析报告（利用Pandoc转换生成）
   - `segmented_words.txt`：经过分词处理的词语列表
   - `sentiment_analysis_pie_chart.png`：情感倾向分析的饼状统计图
   - `wordcloud.png`：反映高频关键词的词云图像
   - 
****
### 本项目遵循MIT许可协议开源。