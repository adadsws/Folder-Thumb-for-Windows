# README 预览图尺寸设计

## 目标

将 README `Demo` 区块中的预览图显示宽度缩小为 README 内容区域的 20%，同时保留点击图片打开 Bilibili 介绍视频的行为。

## 实现方式

GitHub Markdown 图片语法不能设置显示宽度，因此仅将当前图片链接替换为以下英文 HTML：

```html
<a href="https://www.bilibili.com/video/BV1heDdBNEVU"><img src="docs/202604100646.mp4_20260726_161435.607.jpg" alt="Folder thumbnail preview" width="20%"></a>
```

- `width="20%"` 使图片相对 README 内容宽度响应式缩放。
- `<a>` 保留点击图片跳转到原 Bilibili 视频的行为。
- `alt` 保持英文无障碍文本。
- 图片下方现有的英文文字链接保持不变。

## 变更范围

- 更新根目录 `README.md` 中的一行图片标记。
- 更新 `CHANGELOG.md`，记录响应式缩小预览图。
- 不修改、压缩或重新提交 JPEG 内容。
- 不修改脚本、配置、测试或 README 其他章节。

## 验证

- 确认 README 中只有一个 `width="20%"`。
- 确认图片 HTML 和文字链接仍指向 `https://www.bilibili.com/video/BV1heDdBNEVU`。
- 确认图片路径和英文替代文本保持正确。
- 运行仓库合规检查和 Markdown diff 检查。
- 只有在用户明确要求 push 时才更新脱敏导出仓库和远程 `main`。
