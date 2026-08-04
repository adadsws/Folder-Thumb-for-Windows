# README Demo 区块设计

## 目标

在保持 README 紧凑、英文和终端用户导向的前提下，加入项目介绍视频与效果预览图，让访问者能够直接看到工具效果并进入 Bilibili 观看完整演示。

## 展示结构

在项目标题和一句话简介之后、`Getting Started` 之前新增 `Demo` 区块。区块使用标准 GitHub Markdown，不使用 HTML：

```markdown
## Demo

[![Folder thumbnail preview](docs/202604100646.mp4_20260726_161435.607.jpg)](https://www.bilibili.com/video/BV1heDdBNEVU)

[Watch the introduction video on Bilibili.](https://www.bilibili.com/video/BV1heDdBNEVU)
```

预览图本身和文字链接都指向同一视频。图片的英文替代文本用于图片无法加载和屏幕阅读器场景。

## 图片资产

- 使用用户提供的 `docs/202604100646.mp4_20260726_161435.607.jpg`。
- 保持原文件内容、尺寸和文件名，不压缩、不裁剪、不重新生成。
- README 使用仓库相对路径，确保 GitHub 能从仓库内容直接渲染图片。
- 图片属于普通必要文档资产，允许进入脱敏导出仓库。

## 变更范围

- 新增并本地提交预览图文件。
- 更新 `README.md`，新增英文 `Demo` 区块。
- 更新 `CHANGELOG.md`，记录介绍视频和预览图。
- 更新尚未 push 的 `output/github-export/` 导出提交，使远程最终只收到包含完整 README 更新、预览图和 CHANGELOG 的一个新提交。
- 不修改脚本、配置、测试或其他功能。

## 发布约束

- Bilibili URL 固定为 `https://www.bilibili.com/video/BV1heDdBNEVU`。
- 不推送 `~ref/`、`~archived/`、活动计划或其他本地专用路径。
- 导出仓库继续只保留 `main`，从远程 `main` fast-forward push，不 force push，不创建 PR。

## 验证

- 确认图片存在、可读取，尺寸为 760×673。
- 确认 README 仍全为英文，`Demo` 位于简介与 `Getting Started` 之间。
- 检查图片 Markdown 和文字链接均使用正确 URL。
- 在源仓库和导出仓库运行文档、合规及 pre-push 检查。
- push 后确认远程 `main` 包含图片和更新后的 README，且三个仓库均只有 `main`。
