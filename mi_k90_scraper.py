#!/usr/bin/env python3
"""
小米社区 K90系列游戏相关帖子数据抓取
抓取K90 Max、K90/PM圈子的游戏相关帖子
"""

import requests
from bs4 import BeautifulSoup
import json
import re
import io
import sys
import time
from datetime import datetime
from openpyxl import Workbook

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# API配置
API_BASE = "https://api.vip.miui.com/mtop/planet/vip/home/discover"

# K90系列圈子配置
K90_CIRCLES = [
    {"name": "K90 Max", "circle": "K90 Max"},
    {"name": "K90/PM", "circle": "K90/PM"},
]

# 游戏关键词（用于过滤游戏相关帖子）
GAME_KEYWORDS = [
    "游戏", "手游", "打游戏", "帧率", "帧数", "掉帧", "卡顿", "发热",
    "王者", "峡谷", "吃鸡", "和平精英", "三角洲", "原神", "鸣潮",
    "无畏契约", "VALORANT", "异环", "绝区零", "崩铁", "明日方舟",
    "元梦", "蛋仔", "我的世界", "mc", "原版", "生存", "建造",
    "高负载", "流畅", "画质", "超分", "抗锯齿", "触控", "跟手",
    "散热", "温控", "功耗", "续航", "电池", "帧率稳定",
    "插帧", "稳帧", "游戏模式", "游戏助手", "肩键", "手柄",
    "120帧", "144帧", "165帧", "185帧", "高帧率",
    "K90max", "K90 Max", "K90max游戏", "K90Max"
]

# 要抓取的页数配置
PAGES_TO_SCRAPE = [1, 2, 3]

def get_circle_posts(circle_name, page=1, page_size=50):
    """获取K90圈子帖子"""
    params = {
        "ref": "",
        "pathname": "/",
        "version": "dev.undefined",
        "offset": page,
        "pageNum": page,
        "dotEntrance": "CIRCLE_DETAIL",
        "circle": circle_name,
        "pageSize": page_size,
        "instanceId": "k90_scraper_" + str(int(time.time()))
    }

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Referer": "https://web.vip.miui.com/",
    }

    try:
        response = requests.get(API_BASE, params=params, headers=headers, timeout=30)
        response.encoding = "utf-8"
        return response.json()
    except Exception as e:
        print(f"获取圈子 {circle_name} 第{page}页出错: {e}")
        return None


def extract_posts_from_response(data):
    """从API响应中提取帖子列表"""
    posts = []
    try:
        records = data.get("entity", {}).get("recommend", {}).get("records", [])
        for record in records:
            posts.append(record)
    except Exception as e:
        print(f"解析响应出错: {e}")
    return posts


def is_game_related(text_content, tags=""):
    """判断帖子是否与游戏相关"""
    text = (text_content + " " + tags).lower()
    for keyword in GAME_KEYWORDS:
        if keyword.lower() in text:
            return True
    return False


def get_topic_content(post_id):
    """通过autocli读取帖子详情"""
    try:
        from subprocess import run, PIPE
        result = run(
            ["autocli", "read",
             f"https://web.vip.miui.com/page/info/mio/mio/topic?topicId={post_id}",
             "-f", "json"],
            capture_output=True, text=True, timeout=30,
            encoding="utf-8"
        )
        if result.returncode == 0:
            try:
                data = json.loads(result.stdout)
                return data.get("textContent", "")
            except:
                return result.stdout
        return ""
    except Exception as e:
        return ""


def classify_by_feature(text):
    """功能特性分类"""
    text_lower = text.lower()

    if any(kw in text_lower for kw in ["帧率", "165帧", "185帧", "144帧", "120帧", "帧率增强", "插帧", "高帧"]):
        return "帧率适配"
    if any(kw in text_lower for kw in ["144hz", "165hz", "185hz", "刷新率", "高刷", "屏幕刷新"]):
        return "刷新率档位"
    if any(kw in text_lower for kw in ["画质", "抗锯齿", "超分", "超分辨率", "画质增强", "抗锯齿"]):
        return "画质优化"
    if any(kw in text_lower for kw in ["稳帧", "帧率稳定", "帧稳定", "稳帧技术"]):
        return "幻影稳帧"
    if any(kw in text_lower for kw in ["极客中心", "极客模式", "性能模式", "火力全开", "调度", "温控", "性能控制"]):
        return "极客中心"
    if any(kw in text_lower for kw in ["分身", "双开", "应用分身", "游戏分身", "多开"]):
        return "应用分身"
    if any(kw in text_lower for kw in ["适配", "游戏适配", "新游", "优化", "游戏优化", "适配游戏"]):
        return "游戏适配"
    if any(kw in text_lower for kw in ["温控", "续航", "发热", "降亮度", "电池", "功耗", "散热"]):
        return "温控续航"
    if any(kw in text_lower for kw in ["手柄", "肩键", "外设", "游戏手柄", "游戏辅助"]):
        return "外设支持"
    return "其他功能"


def classify_by_game(text):
    """游戏分类"""
    text_lower = text.lower()

    if any(kw in text_lower for kw in ["三角洲行动", "三角洲"]):
        return "三角洲行动"
    if any(kw in text_lower for kw in ["王者荣耀", "王者", "峡谷", "王者荣誉"]):
        return "王者荣耀"
    if any(kw in text_lower for kw in ["无畏契约", "瓦", "valorant"]):
        return "无畏契约"
    if any(kw in text_lower for kw in ["明日方舟", "方舟", "终末地"]):
        return "明日方舟"
    if any(kw in text_lower for kw in ["和平精英", "吃鸡"]):
        return "和平精英"
    if any(kw in text_lower for kw in ["原神"]):
        return "原神"
    if any(kw in text_lower for kw in ["鸣潮"]):
        return "鸣潮"
    if any(kw in text_lower for kw in ["异环"]):
        return "异环"
    if any(kw in text_lower for kw in ["绝区零", "绝区"]):
        return "绝区零"
    if any(kw in text_lower for kw in ["崩铁", "崩坏星穹铁道"]):
        return "崩坏星穹铁道"
    if any(kw in text_lower for kw in ["元梦", "元梦之星", "蛋仔"]):
        return "元梦之星/蛋仔派对"
    if any(kw in text_lower for kw in ["我的世界", "minecraft", "mc"]):
        return "我的世界"
    if any(kw in text_lower for kw in ["游戏", "打游戏", "手游", "打游戏"]):
        return "其他游戏"
    return "非游戏内容"


def save_to_json(posts, filename="mi_k90_data.json"):
    """保存到JSON文件"""
    output = {
        "crawl_time": datetime.now().isoformat(),
        "source": "小米社区 K90系列",
        "total_posts": len(posts),
        "posts": posts
    }
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"JSON已保存: {filename}")


def save_to_excel(posts, filename="mi_k90_threads.xlsx"):
    """保存到Excel文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "K90游戏帖子"

    headers = ["序号", "帖子ID", "圈子", "设备", "发帖内容", "评论数", "点赞数", "作者", "标签", "功能分类", "游戏分类"]
    ws.append(headers)

    for i, post in enumerate(posts, 1):
        row = [
            i,
            post.get("id", ""),
            post.get("circle", ""),
            post.get("deviceType", ""),
            post.get("textContent", "")[:200],
            post.get("commentCnt", 0),
            post.get("likeCnt", 0),
            post.get("author", {}).get("name", ""),
            post.get("tags", ""),
            post.get("feature_category", ""),
            post.get("game_category", "")
        ]
        ws.append(row)

    wb.save(filename)
    print(f"Excel已保存: {filename}")


def save_to_html(posts, filename="mi_k90_threads.html"):
    """保存到HTML文件"""
    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>小米社区 K90系列游戏帖子</title>
    <style>
        body {{ font-family: "Microsoft YaHei", sans-serif; margin: 20px; }}
        h1 {{ color: #ff6700; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; }}
        th {{ background: #f5f5f5; }}
        tr:nth-child(even) {{ background: #f9f9f9; }}
        a {{ color: #0066cc; text-decoration: none; }}
        .stats {{ margin: 20px 0; padding: 15px; background: #fff3e0; border-radius: 5px; }}
        .tag-game {{ background: #e3f2fd; padding: 2px 6px; border-radius: 3px; }}
    </style>
</head>
<body>
    <h1>小米社区 K90系列游戏相关帖子</h1>
    <div class="stats">
        <p><strong>抓取时间:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        <p><strong>总帖子数:</strong> {len(posts)}</p>
    </div>
    <table>
        <tr>
            <th>序号</th><th>帖子ID</th><th>圈子</th><th>设备</th><th>发帖内容</th>
            <th>评论数</th><th>点赞</th><th>作者</th><th>功能分类</th><th>游戏分类</th>
        </tr>
"""

    for i, post in enumerate(posts, 1):
        content_short = post.get("textContent", "")[:80] + "..." if len(post.get("textContent", "")) > 80 else post.get("textContent", "")
        post_id = post.get("id", "")
        html += f"""        <tr>
            <td>{i}</td>
            <td>{post_id}</td>
            <td>{post.get('circle', '')}</td>
            <td>{post.get('deviceType', '')}</td>
            <td title="{post.get('textContent', '')}">{content_short}</td>
            <td>{post.get('commentCnt', 0)}</td>
            <td>{post.get('likeCnt', 0)}</td>
            <td>{post.get('author', {}).get('name', '')}</td>
            <td><span class="tag-game">{post.get('feature_category', '')}</span></td>
            <td>{post.get('game_category', '')}</td>
        </tr>
"""

    html += """    </table>
</body>
</html>"""

    with open(filename, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"HTML已保存: {filename}")


def main():
    print("=" * 60)
    print("小米社区 K90系列游戏帖子数据抓取")
    print("=" * 60)
    print()

    all_posts = []

    for circle_info in K90_CIRCLES:
        circle_name = circle_info["name"]
        circle_key = circle_info["circle"]
        print(f"\n正在抓取圈子: {circle_name}")

        for page in PAGES_TO_SCRAPE:
            print(f"  获取第{page}页...")
            data = get_circle_posts(circle_key, page)

            if data:
                records = extract_posts_from_response(data)
                print(f"    第{page}页获取到 {len(records)} 条记录")

                for record in records:
                    text_content = record.get("textContent", "")
                    tags = record.get("tags", "")

                    if is_game_related(text_content, tags):
                        # 分类
                        record["circle"] = circle_name
                        record["feature_category"] = classify_by_feature(text_content)
                        record["game_category"] = classify_by_game(text_content)
                        record["url"] = f"https://web.vip.miui.com/page/info/mio/mio/topic?topicId={record.get('id')}"
                        all_posts.append(record)
                        print(f"    [游戏相关] {record.get('id')} - {text_content[:50]}...")
            else:
                print(f"    第{page}页获取失败")

            time.sleep(1)  # 避免请求过快

    print()
    print("=" * 60)
    print("抓取结果汇总")
    print("=" * 60)
    print(f"总计游戏相关帖子数: {len(all_posts)}")

    if all_posts:
        save_to_json(all_posts, "mi_k90_data.json")
        save_to_excel(all_posts, "mi_k90_threads.xlsx")
        save_to_html(all_posts, "mi_k90_threads.html")

        # 统计
        feature_stats = {}
        game_stats = {}
        for post in all_posts:
            fc = post.get("feature_category", "其他功能")
            gc = post.get("game_category", "其他游戏")
            feature_stats[fc] = feature_stats.get(fc, 0) + 1
            game_stats[gc] = game_stats.get(gc, 0) + 1

        print("\n功能特性分布:")
        for k, v in sorted(feature_stats.items(), key=lambda x: -x[1]):
            print(f"  {k}: {v} ({v*100/len(all_posts):.1f}%)")

        print("\n游戏分布:")
        for k, v in sorted(game_stats.items(), key=lambda x: -x[1]):
            print(f"  {k}: {v} ({v*100/len(all_posts):.1f}%)")

        # 生成报告
        report = f"""# 小米社区 K90系列游戏反馈分析报告

生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## 数据概览

- **抓取范围**: K90 Max、K90/PM 圈子
- **筛选条件**: 游戏相关内容（包含游戏、帧率、画质等关键词）
- **总帖子数**: {len(all_posts)}

## 功能特性分布

| 功能类别 | 数量 | 占比 |
|---------|------|------|
"""
        for k, v in sorted(feature_stats.items(), key=lambda x: -x[1]):
            report += f"| {k} | {v} | {v*100/len(all_posts):.1f}% |\n"

        report += """
## 游戏分布

| 游戏 | 数量 | 占比 |
|------|------|------|
"""
        for k, v in sorted(game_stats.items(), key=lambda x: -x[1]):
            report += f"| {k} | {v} | {v*100/len(all_posts):.1f}% |\n"

        report += """
## 典型帖子内容

"""
        for i, post in enumerate(all_posts[:10], 1):
            report += f"### {i}. {post.get('textContent', '')[:100]}...\n"
            report += f"- 圈子: {post.get('circle')} | 设备: {post.get('deviceType')} | 功能分类: {post.get('feature_category')} | 游戏分类: {post.get('game_category')}\n\n"

        with open("mi_k90_report.md", "w", encoding="utf-8") as f:
            f.write(report)
        print("\n分析报告已保存: mi_k90_report.md")

    print("\n生成文件:")
    print("  - mi_k90_data.json")
    print("  - mi_k90_threads.xlsx")
    print("  - mi_k90_threads.html")
    print("  - mi_k90_report.md")


if __name__ == "__main__":
    main()
