#!/usr/bin/env node
/**
 * 荣耀建议广场数据抓取脚本
 * 使用firecrawl爬取所有页面
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const BASE_URL = "https://club.honor.com/cn//cn//opinion_thread-list.html";
const SCRAPE_ID = "019e5510-2b72-753c-9ed5-8f27b30fc9e0"; // 最新的scrape ID
const TOTAL_PAGES = 19;

const outputDir = path.join(__dirname, 'honor_data');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// 帖子数据存储
let allPosts = [];

// 执行firecrawl interact命令
function runInteract(prompt, scrapeId = SCRAPE_ID) {
    try {
        const cmd = `firecrawl interact -p "${prompt.replace(/"/g, '\\"')}" --scrape-id ${scrapeId}`;
        const result = execSync(cmd, { encoding: 'utf8', timeout: 120000 });
        return result;
    } catch (error) {
        console.error("命令执行失败:", error.message);
        return null;
    }
}

// 解析帖子数据
function parsePosts(jsonText) {
    try {
        // 提取JSON数组
        const match = jsonText.match(/\[[\s\S]*\]/);
        if (match) {
            return JSON.parse(match[0]);
        }
    } catch (e) {
        console.error("解析JSON失败:", e.message);
    }
    return [];
}

// 主函数
async function main() {
    console.log("=".repeat(60));
    console.log("荣耀建议广场 - 游戏体验板块 - 荣耀WIN系列数据抓取");
    console.log("=".repeat(60));
    console.log(`总共 ${TOTAL_PAGES} 页需要抓取`);

    for (let page = 1; page <= TOTAL_PAGES; page++) {
        console.log(`\n正在抓取第 ${page}/${TOTAL_PAGES} 页...`);

        // 如果不是第一页，需要点击翻页
        if (page > 1) {
            const pagePrompt = `点击第 ${page} 页的分页按钮，等待页面加载完成`;
            console.log(`执行翻页操作...`);
            runInteract(pagePrompt);
            await new Promise(r => setTimeout(r, 3000));
        }

        // 抓取当前页面的帖子数据
        const extractPrompt = `提取当前页面所有帖子的详细信息，返回JSON数组格式。每个帖子包含：
- user_id: 用户ID
- title: 发帖标题
- post_time: 发帖时间
- view_count: 浏览数量
- comment_count: 评论数量
- post_url: 发帖链接`;

        const result = runInteract(extractPrompt);
        if (result) {
            const posts = parsePosts(result);
            posts.forEach(p => p.page_number = page);
            allPosts.push(...posts);
            console.log(`第 ${page} 页抓取到 ${posts.length} 个帖子`);
        }

        await new Promise(r => setTimeout(r, 2000));
    }

    console.log(`\n总共抓取到 ${allPosts.length} 个帖子`);

    // 保存数据
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);

    // JSON
    fs.writeFileSync(
        path.join(outputDir, `honor_win_posts_${timestamp}.json`),
        JSON.stringify(allPosts, null, 2),
        'utf8'
    );
    console.log(`JSON数据已保存`);

    // HTML
    generateHTML(allPosts, timestamp);
    console.log(`HTML数据已保存`);

    // XLSX需要额外处理
    generateXLSX(allPosts, timestamp);
    console.log(`XLSX数据已保存`);

    console.log("\n数据抓取完成！");
}

function generateHTML(posts, timestamp) {
    let html = `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>荣耀建议广场 - 游戏体验板块 - 荣耀WIN系列数据</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #cf1f25; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #cf1f25; color: white; }
        tr:nth-child(even) { background-color: #f9f9f9; }
    </style>
</head>
<body>
    <h1>荣耀建议广场 - 游戏体验板块 - 荣耀WIN系列数据</h1>
    <p>抓取时间: ${new Date().toLocaleString('zh-CN')}</p>
    <p>共抓取 ${posts.length} 个帖子</p>
    <table>
        <tr>
            <th>序号</th><th>用户ID</th><th>产品名</th><th>发帖标题</th>
            <th>发帖时间</th><th>浏览数量</th><th>评论数量</th><th>页号</th>
        </tr>`;

    posts.forEach((p, i) => {
        html += `\n        <tr>
            <td>${i + 1}</td>
            <td>${p.user_id || ''}</td>
            <td>荣耀WIN系列</td>
            <td>${p.title || ''}</td>
            <td>${p.post_time || ''}</td>
            <td>${p.view_count || 0}</td>
            <td>${p.comment_count || 0}</td>
            <td>${p.page_number || 1}</td>
        </tr>`;
    });

    html += `\n    </table>
</body>
</html>`;

    fs.writeFileSync(
        path.join(outputDir, `honor_win_posts_${timestamp}.html`),
        html,
        'utf8'
    );
}

function generateXLSX(posts, timestamp) {
    // 使用csv格式代替xlsx（需要openpyxl才支持xlsx）
    const headers = ["序号", "用户ID", "产品名", "发帖标题", "发帖时间", "浏览数量", "评论数量", "页号"];
    let csv = headers.join(',') + '\n';

    posts.forEach((p, i) => {
        const row = [
            i + 1,
            `"${(p.user_id || '').replace(/"/g, '""')}"`,
            "荣耀WIN系列",
            `"${(p.title || '').replace(/"/g, '""')}"`,
            p.post_time || '',
            p.view_count || 0,
            p.comment_count || 0,
            p.page_number || 1
        ];
        csv += row.join(',') + '\n';
    });

    fs.writeFileSync(
        path.join(outputDir, `honor_win_posts_${timestamp}.csv`),
        '﻿' + csv, // UTF-8 BOM
        'utf8'
    );
}

main().catch(console.error);