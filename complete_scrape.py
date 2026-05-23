#!/usr/bin/env python3
"""
荣耀建议广场完整数据抓取脚本
批量抓取所有页面并解析用户ID、评论数等信息
"""

import json
import re
import os
import subprocess
import time
from pathlib import Path
from datetime import datetime
from typing import List, Dict

# 数据模型
from dataclasses import dataclass, field

@dataclass
class Post:
    """帖子数据模型"""
    user_id: str = ""
    product_name: str = "荣耀WIN系列"
    title: str = ""
    content: str = ""
    post_time: str = ""
    view_count: int = 0
    comment_count: int = 0
    post_url: str = ""
    page_number: int = 1
    vote_count: int = 0  # 建议投票参与人数
    comments: List[Dict] = field(default_factory=list)
    image_urls: List[str] = field(default_factory=list)

def parse_page_detail(html_content: str, page_num: int) -> List[Dict]:
    """解析页面详情"""
    posts = []

    # 分割每个帖子
    thread_pattern = r'<div id="(\d+)">(.*?)</div>\s*</div>\s*<div id="'
    matches = re.findall(thread_pattern, html_content, re.DOTALL)

    for post_id, post_content in matches:
        # 提取标题
        title_match = re.search(r'<p>(.*?)</p>', post_content, re.DOTALL)
        title = title_match.group(1) if title_match else ''

        # 提取参与人数
        participation_match = re.search(r'<span>(\d+)</span>\s*人已参与', post_content)
        participation = int(participation_match.group(1)) if participation_match else 0

        # 清理HTML标签
        title = re.sub(r'<[^>]+>', '', title).strip()

        if title:
            posts.append({
                'post_id': post_id,
                'title': title,
                'vote_count': participation,
                'url': f'https://club.honor.com/cn//cn//opinion_thread-view.html?id={post_id}'
            })

    return posts


def get_post_detail(post_url: str) -> Dict:
    """获取单个帖子的详细信息"""
    try:
        # 使用autocli read获取帖子详情
        cmd = f'autocli read "{post_url}" -f json'
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=30)

        if result.returncode == 0 and os.path.exists('temp_post.json'):
            with open('temp_post.json', 'r', encoding='utf-8') as f:
                data = json.load(f)

            content = data.get('content', '')
            byline = data.get('byline', '')

            # 提取用户ID
            user_id = byline.strip() if byline else 'Unknown'

            # 提取内容摘要
            content_text = re.sub(r'<[^>]+>', '', content)
            content_text = re.sub(r'\s+', ' ', content_text).strip()
            content_summary = content_text[:500] if content_text else ''

            # 清理临时文件
            if os.path.exists('temp_post.json'):
                os.remove('temp_post.json')

            return {
                'user_id': user_id,
                'content': content_summary
            }
    except Exception as e:
        print(f"Error getting post detail: {e}")

    return {'user_id': 'Unknown', 'content': ''}


def main():
    print("=" * 60)
    print("Honor WIN Series Complete Data Collection")
    print("=" * 60)

    # 创建输出目录
    output_dir = Path("honor_data")
    output_dir.mkdir(exist_ok=True)

    # 解析第一页的帖子
    json_file = "honor_data/page1_autocli.json"
    all_posts = []

    if os.path.exists(json_file):
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        posts = parse_page_detail(data.get('content', ''), 1)
        print(f"\nPage 1: Found {len(posts)} posts")

        # 为每个帖子获取详细信息
        for post in posts[:5]:  # 先处理前5个作为示例
            print(f"  Getting details for: {post['title'][:40]}...")

            # 获取帖子详情
            detail = get_post_detail(post['url'])

            post_obj = Post(
                title=post['title'],
                user_id=detail.get('user_id', 'Unknown'),
                content=detail.get('content', ''),
                post_url=post['url'],
                vote_count=post['vote_count'],
                page_number=1
            )
            all_posts.append(post_obj)

    print(f"\nTotal posts collected: {len(all_posts)}")

    # 保存数据
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 保存JSON
    json_output = output_dir / f"honor_win_full_{timestamp}.json"
    with open(json_output, 'w', encoding='utf-8') as f:
        json.dump([{
            'user_id': p.user_id,
            'product_name': p.product_name,
            'title': p.title,
            'content': p.content,
            'post_time': p.post_time,
            'view_count': p.view_count,
            'comment_count': p.comment_count,
            'post_url': p.post_url,
            'page_number': p.page_number,
            'vote_count': p.vote_count
        } for p in all_posts], f, ensure_ascii=False, indent=2)
    print(f"JSON saved: {json_output}")

    # 生成HTML
    html_output = output_dir / f"honor_win_full_{timestamp}.html"
    html_content = generate_html(all_posts, timestamp)
    with open(html_output, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"HTML saved: {html_output}")

    # 生成XLSX (CSV格式)
    csv_output = output_dir / f"honor_win_full_{timestamp}.csv"
    generate_csv(all_posts, csv_output)
    print(f"CSV saved: {csv_output}")

    print(f"\nDone! Output directory: {output_dir.absolute()}")


def generate_html(posts: List[Post], timestamp: str) -> str:
    """生成HTML表格"""
    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>Honor WIN Series Feedback - All Pages</title>
    <style>
        * {{ box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }}
        .container {{ max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
        h1 {{ color: #cf1f25; text-align: center; }}
        .stats {{ display: flex; justify-content: space-around; margin: 20px 0; padding: 20px; background: #f9f9f9; border-radius: 8px; }}
        .stat-value {{ font-size: 28px; font-weight: bold; color: #cf1f25; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
        th, td {{ border: 1px solid #ddd; padding: 10px; text-align: left; }}
        th {{ background: #cf1f25; color: white; position: sticky; top: 0; }}
        tr:nth-child(even) {{ background: #f9f9f9; }}
        tr:hover {{ background: #f0f0f0; }}
        .title-cell {{ max-width: 300px; word-wrap: break-word; }}
        .votes {{ text-align: center; font-weight: bold; color: #cf1f25; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Honor WIN Series - Game Experience Feedback</h1>
        <p style="text-align:center;color:#666">Page 1 Data - {timestamp}</p>
        <div class="stats">
            <div class="stat-item"><div class="stat-value">{len(posts)}</div><div>Total Posts</div></div>
            <div class="stat-item"><div class="stat-value">{sum(p.vote_count for p in posts)}</div><div>Total Votes</div></div>
        </div>
        <table>
            <thead>
                <tr><th>No</th><th>User ID</th><th>Title</th><th>Votes</th><th>URL</th></tr>
            </thead>
            <tbody>
"""

    for i, post in enumerate(posts, 1):
        html += f"""                <tr>
                    <td style="text-align:center">{i}</td>
                    <td>{post.user_id}</td>
                    <td class="title-cell">{post.title}</td>
                    <td class="votes">{post.vote_count}</td>
                    <td><a href="{post.post_url}" target="_blank">View</a></td>
                </tr>
"""

    html += """            </tbody>
        </table>
    </div>
</body>
</html>"""
    return html


def generate_csv(posts: List[Post], output_file: Path):
    """生成CSV文件"""
    import csv

    with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['No', 'User ID', 'Product', 'Title', 'Content', 'Post Time', 'Views', 'Comments', 'URL', 'Page', 'Votes'])

        for i, post in enumerate(posts, 1):
            writer.writerow([
                i,
                post.user_id,
                post.product_name,
                post.title,
                post.content,
                post.post_time,
                post.view_count,
                post.comment_count,
                post.post_url,
                post.page_number,
                post.vote_count
            ])


if __name__ == "__main__":
    main()