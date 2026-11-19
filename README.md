# Txt2Excel

浏览器本地批量处理 CHI 电化学 TXT，整理成可直接复制到 Origin 的 Excel。

## 功能

- `Origin Data`：文件名、列名和完整数据按统一起始行对齐
- `Statistics`：每个文件最后 30 秒的所有测量列平均值
- 可选 `E-T Curves`：第一列为连续时间，之后每个 E 一列，按 T 顺序纵向拼接完整原始曲线
- 时间列保持数值类型并至少显示一位小数，便于直接导入 Origin
- 同时生成 `E-T Average`：按 E 列、T 行排列平均值
- 完整数据虚拟预览，可切换查看原始数据、统计、拼接曲线和平均矩阵
- 支持横向并排或纵向追加
- 可选生成按参数名对齐的 `Header Info` 工作表

## 本地使用

直接打开 [index.html](/Users/LeeChung/Txt2Excel/index.html)。

- 选择一个或多个 CHI 导出的 `txt` 文件
- 设置统计时长和数据排列方式
- 导出包含数据与统计结果的 `.xlsx`

如果网络受限导致 `SheetJS` CDN 无法加载，需要联网后刷新页面再导出。

## GitHub Pages

这个项目是纯静态网页，部署到 GitHub Pages 后即可直接使用。

- 文件在浏览器本地读取
- 解析和合并在浏览器本地完成
- 不会上传文件内容到服务器
