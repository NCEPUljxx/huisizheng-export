# 慧思政题库导出

一个浏览器扩展（Chrome / Edge），用于一键导出[慧思政平台](https://pc.hdhuisizheng.com/)的题库，生成 Word 文档，支持三种导出模式。

## 功能

在慧思政平台的试卷分析页面自动注入一个悬浮面板，提供三个导出按钮：

| 模式 | 说明 |
| --- | --- |
| ✅ 导出 Word（带答案） | 题目 + 全部选项，正确答案用红色标注「（正确答案）」 |
| 📝 导出 Word（无答案） | 题目 + 全部选项，不标注答案 |
| 🔑 导出 Word（仅答案） | 只导出每道题的正确答案 |

导出结果为 `.doc` 文件（Word HTML 格式），包含：

- 自动识别**单选题**和**判断题**（按选项数量判断，2 个选项视为判断题）
- 分章节排版：「一、单选题」「二、判断题」
- 标题、题目、选项均有对应样式（宋体 / 黑体，A4 纸张）

## 文件结构

```
huisizheng-export/
├── manifest.json   # 扩展清单（Manifest V3）
├── content.js      # 核心脚本：题目提取、Word 生成、面板交互
├── content.css     # 悬浮面板样式
├── icons/          # 扩展图标（16 / 48 / 128）
└── README.md
```

## 安装

1. 打开 Chrome 或 Edge 的扩展管理页：
   - Chrome：地址栏输入 `chrome://extensions`
   - Edge：地址栏输入 `edge://extensions`
2. 打开右上角的「开发者模式」开关
3. 点击「加载已解压的扩展程序」，选择本文件夹（`huisizheng-export`）
4. 打开慧思政平台的试卷分析页面，即可在页面右侧看到「📋 慧思政导出」悬浮面板

## 使用

1. 进入慧思政平台的试卷**分析页面**（`/exam/analysis`）
2. 等待题目加载完成（面板右上角会显示题目数量）
3. 点击对应的导出按钮，即可下载 `.doc` 文件

## 工作原理

- 通过 `MutationObserver` 监听 DOM 变化，实时检测题目是否加载完成
- 通过 `setInterval` 轮询 URL 哈希值，兼容不触发 `hashchange` 事件的单页应用路由
- 试卷名称通过 `sessionStorage` 在结果页与分析页之间传递
- Word 文档由纯 HTML 生成（`application/msword`），无需任何第三方依赖，完全离线运行

## 权限说明

- `activeTab`：在用户访问目标页面时启用
- `host_permissions`：仅匹配 `*://pc.hdhuisizheng.com/*`，不影响其他网站
