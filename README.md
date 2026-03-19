# Txt2Excel

浏览器本地处理 TXT，组合导出为 CSV 或 Excel。

## 功能

- 通用逗号分列
- CHI 电化学数据筛选
- 原始文本末尾行保留
- 横向拼接、纵向合并、双表导出

## 本地使用

直接打开 [index.html](/Users/LeeChung/Txt2Excel/index.html)。

- 选择一个或多个 `txt` 文件
- 选择处理模式和输出布局
- 导出 `csv` 或 `xlsx`

如果网络受限导致 `SheetJS` CDN 无法加载，仍然可以导出 `csv`。

## GitHub Pages

这个项目是纯静态网页，部署到 GitHub Pages 后即可直接使用。

- 文件在浏览器本地读取
- 解析和合并在浏览器本地完成
- 不会上传文件内容到服务器
