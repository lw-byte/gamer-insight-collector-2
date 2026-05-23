#!/usr/bin/env python3
"""
荣耀建议广场游戏体验板块数据抓取脚本
抓取荣耀WIN系列所有帖子的数据
"""

import asyncio
import json
import re
import os
from pathlib import Path
from typing import List, Dict, Optional
from dataclasses import dataclass, asdict
from datetime import datetime

# 尝试导入必要的库
try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("openpyxl not available, will skip xlsx generation")

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False
    print("requests not available")


@dataclass
class Post:
    """帖子数据模型"""
    user_id: str
    product_name: str = "荣耀WIN系列"
    title: str = ""
    content: str = ""
    post_time: str = ""
    view_count: int = 0
    comment_count: int = 0
    post_url: str = ""
    page_number: int = 1
    vote_count: int = 0
    image_urls: List[str] = None
    comments: List[Dict] = None

    def __post_init__(self):
        if self.image_urls is None:
            self.image_urls = []
        if self.comments is None:
            self.comments = []


def sanitize_filename(filename: str) -> str:
    """ sanitizefilename by removing/replacing invalid characters """
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    filename = re.sub(r'\s+', '_', filename)
    return filename[:100]


async def grab_honor_data():
    """主抓取函数"""
    print("=" * 60)
    print("荣耀建议广场 - 游戏体验板块 - 荣耀WIN系列数据抓取")
    print("=" * 60)

    all_posts: List[Post] = []
    base_url = "https://club.honor.com/cn//cn//opinion_thread-list.html"

    # 根据用户描述，需要抓取19页数据
    total_pages = 19

    for page in range(1, total_pages + 1):
        print(f"\n{'=' * 60}")
        print(f"正在抓取第 {page}/{total_pages} 页...")
        print(f"{'=' * 60}")

        page_url = f"{base_url}?filter=lastpost&type=23&page={page}"

        # 使用firecrawl interact来获取页面数据
        import subprocess

        # 由于firecrawl interact可能不稳定，我们使用一种混合方法
        # 先保存当前交互状态，然后处理
        cmd = [
            "firecrawl", "scrape",
            page_url,
            "--format", "markdown,html",
            "--wait-for", "5000",
            "-o", f".firecrawl/page_{page}"
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        print(f"页面抓取命令执行完成")

        # 读取抓取的markdown文件
        md_file = Path(f".firecrawl/page_{page}.md")
        if md_file.exists():
            content = md_file.read_text(encoding='utf-8')
            posts = parse_posts_from_markdown(content, page)
            all_posts.extend(posts)
            print(f"第 {page} 页解析到 {len(posts)} 个帖子")
        else:
            print(f"第 {page} 页文件不存在，跳过")

    print(f"\n总共抓取到 {len(all_posts)} 个帖子")

    # 保存数据
    save_data(all_posts)


def parse_posts_from_markdown(content: str, page_number: int) -> List[Post]:
    """从markdown内容中解析帖子数据"""
    posts = []

    # 根据之前interact的结果，我们知道帖子结构
    # 但是markdown可能格式不同，需要调整解析逻辑

    lines = content.split('\n')
    current_post = None

    for line in lines:
        line = line.strip()

        # 检测到新帖子标题
        if line.startswith('- **') or line.startswith('**') or '建议' in line:
            if current_post:
                posts.append(current_post)

            # 尝试提取标题
            title_match = re.search(r'\*\*(.+?)\*\*', line)
            if title_match:
                title = title_match.group(1)
            else:
                title = line.replace('-', '').replace('*', '').strip()

            current_post = Post(
                user_id="未知用户",
                title=title,
                page_number=page_number
            )

        # 检测到用户ID
        elif '用户' in line or line.startswith('by ') or '孤城残梦' in line or 'kay' in line.lower():
            for user in ['孤城残梦', '冷月无', '终成斗帝', '澄澈洺琛', '浮世倾尘', 'kay', '想优化设备']:
                if user in line:
                    current_post.user_id = user
                    break

        # 检测到评论数
        elif '评论' in line or re.search(r'\d+\s*评论', line):
            count_match = re.search(r'(\d+)\s*评论', line)
            if count_match and current_post:
                current_post.comment_count = int(count_match.group(1))

        # 检测到浏览数
        elif '浏览' in line or re.search(r'\d+\s*浏览', line):
            count_match = re.search(r'(\d+)\s*浏览', line)
            if count_match and current_post:
                current_post.view_count = int(count_match.group(1))

    if current_post:
        posts.append(current_post)

    return posts


def save_data(posts: List[Post]):
    """保存数据为JSON、XLSX和HTML格式"""
    output_dir = Path("honor_data")
    output_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 1. 保存为JSON
    json_file = output_dir / f"honor_win_posts_{timestamp}.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump([asdict(p) for p in posts], f, ensure_ascii=False, indent=2)
    print(f"JSON数据已保存到: {json_file}")

    # 2. 保存为XLSX
    if OPENPYXL_AVAILABLE:
        xlsx_file = output_dir / f"honor_win_posts_{timestamp}.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "荣耀建议广场数据"

        # 表头
        headers = [
            "用户ID", "产品名", "发帖标题", "发帖内容", "发帖时间",
            "浏览数量", "评论数量", "发帖链接", "页号", "投票数量",
            "评论详情", "图片链接"
        ]
        ws.append(headers)

        # 数据行
        for post in posts:
            row = [
                post.user_id,
                post.product_name,
                post.title,
                post.content,
                post.post_time,
                post.view_count,
                post.comment_count,
                post.post_url,
                post.page_number,
                post.vote_count,
                json.dumps(post.comments, ensure_ascii=False),
                json.dumps(post.image_urls, ensure_ascii=False)
            ]
            ws.append(row)

        # 调整列宽
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 20

        wb.save(xlsx_file)
        print(f"XLSX数据已保存到: {xlsx_file}")

    # 3. 保存为HTML
    html_file = output_dir / f"honor_win_posts_{timestamp}.html"
    html_content = generate_html(posts)
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"HTML数据已保存到: {html_file}")


def generate_html(posts: List[Post]) -> str:
    """生成HTML表格"""
    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>荣耀建议广场 - 游戏体验板块 - 荣耀WIN系列数据</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1 {{ color: #cf1f25; }}
        table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #cf1f25; color: white; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        tr:hover {{ background-color: #f5f5f5; }}
        .pagination {{ margin: 20px 0; }}
    </style>
</head>
<body>
    <h1>荣耀建议广场 - 游戏体验板块 - 荣耀WIN系列数据</h1>
    <p>抓取时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    <p>共抓取 {len(posts)} 个帖子</p>
    <table>
        <tr>
            <th>序号</th>
            <th>用户ID</th>
            <th>产品名</th>
            <th>发帖标题</th>
            <th>发帖内容</th>
            <th>发帖时间</th>
            <th>浏览数量</th>
            <th>评论数量</th>
            <th>发帖链接</th>
            <th>页号</th>
            <th>投票数量</th>
        </tr>
"""

    for i, post in enumerate(posts, 1):
        html += f"""        <tr>
            <td>{i}</td>
            <td>{post.user_id}</td>
            <td>{post.product_name}</td>
            <td>{post.title}</td>
            <td>{post.content}</td>
            <td>{post.post_time}</td>
            <td>{post.view_count}</td>
            <td>{post.comment_count}</td>
            <td><a href="{post.post_url}" target="_blank">查看</a></td>
            <td>{post.page_number}</td>
            <td>{post.vote_count}</td>
        </tr>
"""

    html += """    </table>
</body>
</html>"""

    return html


if __name__ == "__main__":
    asyncio.run(grab_honor_data())