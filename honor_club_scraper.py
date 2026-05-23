#!/usr/bin/env python3
"""
荣耀俱乐部 WIN系列帖子数据抓取
抓取指定页面的所有帖子，支持批量抓取多页
"""

import requests
from bs4 import BeautifulSoup
import json
import re
import io
import sys
from datetime import datetime
from openpyxl import Workbook

# 设置UTF-8编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 目标URL - WIN系列板块 (fid=5570, type=23)
BASE_URL = "https://club.honor.com/cn/opinion_thread-list.html"

# 要抓取的页数配置
PAGES_TO_SCRAPE = [1, 2, 3, 4]

def get_page_content(page=1):
    """获取页面内容"""
    params = {
        'filter': 'lastpost',
        'fid': '5570',
        'type': '23',
        'page': page
    }

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Referer': 'https://club.honor.com/cn/'
    }

    try:
        response = requests.get(BASE_URL, params=params, headers=headers, timeout=30)
        response.encoding = 'utf-8'
        return response.text
    except Exception as e:
        print(f"获取页面 {page} 出错: {e}")
        return None

def extract_threads(html, page_num=1):
    """从HTML中提取帖子信息"""
    soup = BeautifulSoup(html, 'html.parser')
    threads = []

    # 查找所有帖子条目
    thread_divs = soup.find_all('div', class_='thread-child', id=re.compile(r'^\d+$'))

    for div in thread_divs:
        thread_id = div.get('id')
        content_div = div.find('div', class_='thread-child-con')
        if not content_div:
            continue

        full_text = content_div.get_text(separator=' ', strip=True)

        # 提取参与人数
        participants_match = re.search(r'(\d+)\s*人已参与', full_text)
        participants = int(participants_match.group(1)) if participants_match else 0

        # 提取产品ID
        product_match = re.search(r'HONOR(\d+)', full_text)
        product_id = product_match.group(1) if product_match else ''

        # 提取评论数
        comments_match = re.search(r'最新评论.*?(\d+)', full_text)
        comments = int(comments_match.group(1)) if comments_match else 0

        # 提取最新评论时间
        time_match = re.search(r'最新评论\s*(.+)', full_text)
        latest_comment_time = time_match.group(1).strip() if time_match else ''

        # 提取支持/反对比例
        support_pct = 0
        oppose_pct = 0
        pct_match = re.search(r'(\d+)%\s*(\d+)%', full_text)
        if pct_match:
            support_pct = int(pct_match.group(1))
            oppose_pct = int(pct_match.group(2))

        # 清理帖子内容
        content = full_text
        content = re.sub(r'\s*\d+\s*人已参与.*$', '', content)
        content = re.sub(r'\s*\d+%\s*\d+%.*$', '', content)
        content = re.sub(r'\s*HONOR\d+.*$', '', content)
        content = re.sub(r'\s*最新评论.*$', '', content)
        content = re.sub(r'\s*支持.*$', '', content)
        content = re.sub(r'\s*反对.*$', '', content)
        content = re.sub(r'\s+', ' ', content).strip()

        if not content:
            continue

        thread_url = f"https://club.honor.com/cn/thread-{thread_id}-1-1.html"

        thread_data = {
            'thread_id': thread_id,
            'content': content,
            'participants': participants,
            'support_pct': support_pct,
            'oppose_pct': oppose_pct,
            'product_id': product_id,
            'comments': comments,
            'latest_comment_time': latest_comment_time,
            'url': thread_url,
            'page': page_num
        }

        threads.append(thread_data)

    return threads

def get_product_name(product_id):
    """根据产品ID获取产品名称"""
    product_map = {
        '250': '荣耀250', '260': '荣耀260', '300': '荣耀300',
        '400': '荣耀400', '500': '荣耀500', 'GT': '荣耀GT',
        'WIN': '荣耀WIN', 'WINRT': '荣耀WIN RT',
    }
    return product_map.get(product_id, f'荣耀{product_id}')

def save_to_excel(threads, filename='honor_win_threads.xlsx'):
    """保存到Excel文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "荣耀WIN系列帖子"

    headers = ['页号', '帖子ID', '产品名', '发帖内容', '参与人数', '支持率', '评论数', '发帖时间', '帖子链接']
    ws.append(headers)

    for thread in threads:
        product_name = get_product_name(thread['product_id'])
        row = [
            thread['page'], thread['thread_id'], product_name, thread['content'],
            thread['participants'], f"{thread['support_pct']}%" if thread['support_pct'] else '',
            thread['comments'], thread['latest_comment_time'], thread['url']
        ]
        ws.append(row)

    wb.save(filename)
    print(f"Excel已保存: {filename}")

def save_to_html(threads, filename='honor_win_threads.html'):
    """保存到HTML文件"""
    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>荣耀WIN系列帖子</title>
    <style>
        body {{ font-family: "Microsoft YaHei", sans-serif; margin: 20px; }}
        h1 {{ color: #cf0a2c; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 10px; }}
        th {{ background: #f5f5f5; }}
        tr:nth-child(even) {{ background: #f9f9f9; }}
        a {{ color: #0066cc; text-decoration: none; }}
        .stats {{ margin: 20px 0; padding: 15px; background: #f0f0f0; border-radius: 5px; }}
    </style>
</head>
<body>
    <h1>荣耀俱乐部 WIN系列帖子</h1>
    <div class="stats">
        <p><strong>抓取时间:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        <p><strong>总帖子数:</strong> {len(threads)}</p>
    </div>
    <table>
        <tr>
            <th>序号</th><th>页号</th><th>帖子ID</th><th>产品</th><th>发帖内容</th>
            <th>参与人数</th><th>支持率</th><th>评论数</th><th>发帖时间</th><th>操作</th>
        </tr>
"""

    for i, thread in enumerate(threads, 1):
        product_name = get_product_name(thread['product_id'])
        content_short = thread['content'][:100] + '...' if len(thread['content']) > 100 else thread['content']
        html += f"""        <tr>
            <td>{i}</td>
            <td>第{thread['page']}页</td>
            <td>{thread['thread_id']}</td>
            <td>{product_name}</td>
            <td title="{thread['content']}">{content_short}</td>
            <td>{thread['participants']}</td>
            <td>{thread['support_pct']}%</td>
            <td>{thread['comments']}</td>
            <td>{thread['latest_comment_time']}</td>
            <td><a href="{thread['url']}" target="_blank">查看</a></td>
        </tr>
"""

    html += """    </table>
</body>
</html>"""

    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"HTML已保存: {filename}")

def main():
    print("=" * 60)
    print("荣耀俱乐部 WIN系列帖子数据抓取")
    print("=" * 60)
    print()

    all_threads = []

    for page in PAGES_TO_SCRAPE:
        print(f"正在获取第{page}页...")
        html = get_page_content(page)
        if html:
            threads = extract_threads(html, page)
            all_threads.extend(threads)
            print(f"第{page}页共找到 {len(threads)} 个帖子")
        else:
            print(f"第{page}页获取失败")

    print()
    print("=" * 60)
    print("抓取结果汇总")
    print("=" * 60)
    print(f"总计帖子数: {len(all_threads)}")

    # 保存JSON
    output = {
        'crawl_time': datetime.now().isoformat(),
        'source_url': BASE_URL,
        'pages_scraped': PAGES_TO_SCRAPE,
        'total_threads': len(all_threads),
        'all_threads': all_threads
    }

    with open('honor_club_data.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    if all_threads:
        save_to_excel(all_threads, 'honor_win_threads.xlsx')
        save_to_html(all_threads, 'honor_win_threads.html')

    print("\n生成文件:")
    print("  - honor_win_threads.xlsx")
    print("  - honor_win_threads.html")
    print("  - honor_club_data.json")

if __name__ == '__main__':
    main()