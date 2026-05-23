# Honor WIN Insights Collector

荣耀俱乐部 WIN系列用户反馈数据采集工具

## 项目结构

```
gamer-insight-collector-2/
├── honor_club_scraper.py      # 荣耀俱乐部数据抓取脚本
├── honor_data/                 # 数据输出目录
├── images/                     # 图片目录
└── .claude/
    └── skills/
        ├── honor-club-scraper.md  # Skill定义
        └── README.md              # Skills索引
```

## Skills

### honor-club-scraper

抓取荣耀俱乐部建议广场 WIN系列游戏体验板块帖子数据。

**使用方式:**
```bash
python honor_club_scraper.py
```

**输出文件:**
- `honor_win_threads.xlsx` - Excel格式
- `honor_win_threads.html` - HTML表格
- `honor_club_data.json` - 完整JSON数据

**抓取统计:**
- 第1页: 39条
- 第2页: 38条
- 第3页: 38条
- 第4页: 35条
- 总计: 150条

**数据字段:**
- thread_id: 帖子ID
- content: 发帖内容
- participants: 参与投票人数
- support_pct/oppose_pct: 支持/反对率
- comments: 评论数
- latest_comment_time: 最新评论时间
- url: 帖子链接
- page: 页码