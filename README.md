# CNKI Journal Crawler

本项目使用 Python 和 Selenium 编写，用于自动爬取中国知网（CNKI）上指定期刊与文章标题对应的作者信息及其历史发表文献。

## 🧩 功能概述

- 按期刊和标题自动检索文献
- 抓取作者详情页中的所有历史论文（支持翻页）
- 提取并保存论文标题、作者、机构、期刊、页码、大小等信息到 CSV

## 🛠 使用方式

### 环境依赖

你需要先安装以下 Python 包：

```bash
pip install selenium pandas
