#!/usr/bin/env python3
"""
使用autocli批量抓取荣耀建议广场所有页面
"""

import json
import re
import time
from pathlib import Path
from datetime import datetime

def parse_autocli_content(json_file: str) -> list:
    """从autocli提取的JSON中解析帖子数据"""
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    content = data.get('content', '')

    posts = []
    # 分割每个帖子
    thread_pattern = r'<div id="(\d+)">(.*?)</div>\s*</div>\s*<div id="'
    matches = re.findall(thread_pattern, content, re.DOTALL)

    for post_id, post_content in matches:
        # 提取标题
        title_match = re.search(r'<p>(.*?)</p>', post_content)
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
                'participation': participation,
                'url': f'https://club.honor.com/cn//cn//opinion_thread-view.html?id={post_id}'
            })

    return posts


def main():
    print("=" * 60)
    print("Honor Suggestion Square - Page 1 Data Analysis")
    print("=" * 60)

    # 解析已抓取的数据
    json_file = "honor_data/page1_autocli.json"
    posts = parse_autocli_content(json_file)

    print(f"\nFound {len(posts)} posts from page 1:")
    for i, post in enumerate(posts[:10], 1):
        print(f"  {i}. {post['title'][:50]}...")
        print(f"     Participation: {post['participation']}")
        print(f"     URL: {post['url']}")

    # 保存解析结果
    output_file = "honor_data/posts_page1_parsed.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(posts, f, ensure_ascii=False, indent=2)
    print(f"\nParsed data saved to: {output_file}")


if __name__ == "__main__":
    main()