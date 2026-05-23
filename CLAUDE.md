# Honor WIN Insights Collector

荣耀俱乐部 WIN系列用户反馈数据采集与分析工具

## 项目结构

```
gamer-insight-collector-2/
├── honor_club_scraper.py      # 数据抓取脚本
├── classify_posts.py          # 分类分析脚本
├── honor_club_data.json       # 原始抓取数据
├── classified_posts.json      # 分类后数据
├── classification_report.md  # 分析报告
├── classified_by_feature.xlsx # 按功能分类
├── classified_by_game.xlsx    # 按游戏分类
├── honor_data/                # 数据输出目录
└── .claude/
    └── skills/
        ├── honor-club-scraper.md  # Skill定义
        └── README.md              # Skills索引
```

## Skills

### honor-club-scraper

抓取荣耀俱乐部建议广场 WIN系列游戏体验板块帖子数据，并进行功能特性和游戏两个维度的分类分析。

**使用方式:**

1. 抓取数据:
```bash
python honor_club_scraper.py
```

2. 分类分析:
```bash
python classify_posts.py
```

## 分类体系

### 功能特性分类

| 类别 | 占比 | 典型需求 |
|------|------|---------|
| 帧率适配 | 36.7% | 165帧、185帧适配 |
| 游戏适配 | 12.0% | 新游戏优化 |
| 应用分身 | 8.0% | 游戏多开 |
| 极客中心 | 8.0% | 性能调度控制 |
| 画质优化 | 5.3% | 抗锯齿、超分 |
| 刷新率档位 | 4.7% | 增加144Hz/165Hz |
| 幻影稳帧 | 3.3% | 新游戏适配 |
| 温控续航 | 2.7% | 续航优化 |
| 外设支持 | 2.0% | 游戏手柄 |

### 游戏分类

| 游戏 | 帖子数 | 占比 |
|------|--------|------|
| 其他游戏 | 70 | 46.7% |
| 三角洲行动 | 28 | 18.7% |
| 王者荣耀 | 13 | 8.7% |
| 无畏契约 | 13 | 8.7% |
| 明日方舟 | 8 | 5.3% |
| 和平精英 | 6 | 4.0% |

## 输出文件

| 文件 | 说明 |
|------|------|
| `honor_club_data.json` | 原始抓取数据 (150条) |
| `classified_posts.json` | 分类后数据 |
| `classified_posts.xlsx` | 分类Excel |
| `classified_by_feature.xlsx` | 按功能分类 |
| `classified_by_game.xlsx` | 按游戏分类 |
| `classification_report.md` | 分析报告 |

## 关键洞察

1. **帧率适配是最大痛点** - 36.7%反馈与此相关
2. **三角洲行动最受关注** - 占游戏类反馈50%+
3. **WIN RT差异化对待** - WIN有插帧但RT没有，用户不满
4. **极客中心下放需求** - 希望Magic平板功能下放到WIN