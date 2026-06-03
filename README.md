# Excel 按列拆分工具

这是一个可部署到 GitHub Pages 的静态网页工具。Excel 文件只在浏览器本地读取，不上传服务器。

## 功能

- 默认按 A 列值拆分
- 支持修改拆分列，例如 B、C、AA
- 首行标题可保留到每个子表
- 空值统一命名为 `未填写`
- 下载 ZIP，内部文件名为 `值_行数.xlsx`

## GitHub Pages 部署

1. 推送到 GitHub 仓库的 `main` 分支。
2. 打开仓库 `Settings -> Pages`。
3. Source 选择 `Deploy from a branch`。
4. Branch 选择 `main`，目录选择 `/root`。
5. 保存后等待 Pages 发布。
