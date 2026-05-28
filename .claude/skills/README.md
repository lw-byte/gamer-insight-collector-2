# Skills 索引

| Skill | 说明 |
|-------|------|
| honor-club-scraper | 抓取荣耀俱乐部 WIN系列帖子数据 |
| xiaomi-k90-scraper | 抓取小米社区 K90系列游戏相关帖子 |

---

## honor-club-scraper

抓取荣耀俱乐部建议广场 WIN系列游戏体验板块的帖子数据。

### 功能
- 批量抓取多个页面 (默认抓取1-4页)
- 提取帖子ID、内容、参与人数、评论数等
- 输出JSON/Excel/HTML格式

### 抓取页面
- 第1页: 39条帖子
- 第2页: 38条帖子
- 第3页: 38条帖子
- 第4页: 35条帖子
- 总计: 150条帖子

### 使用方式
```bash
python honor_club_scraper.py
```

### 输出文件
- `honor_win_threads.xlsx` - Excel格式
- `honor_win_threads.html` - HTML格式
- `honor_club_data.json` - 完整JSON数据

### 数据字段
- thread_id: 帖子ID
- content: 发帖内容
- participants: 参与人数
- support_pct: 支持率
- comments: 评论数
- url: 帖子链接
- page: 页码