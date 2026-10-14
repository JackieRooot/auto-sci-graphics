# CLAUDE.md

本文件供 Claude Code 在本仓库工作时参考。

## 仓库定位

一套风格统一的科研绘图脚本（matplotlib），面向电催化方向的期刊单栏插图：

| 脚本 | 图型 | 输入 Excel 的工作表 |
|---|---|---|
| `FE_figure/FEline.py` | 法拉第效率折线图 | `FE`，可选 `SD` |
| `FE_figure/FEcolumn.py` | 法拉第效率柱状图（默认堆积，`-group` 为分组） | `FE`，可选 `SD` |
| `XRD_figure/XRD.py` | XRD 堆叠谱图 + PDF 卡片竖线 | `Sample`，`Standard` |

典型使用场景：用户把自己的原始数据交给你，你把数据整理成脚本要求的格式，再用脚本出图。**用户要的是“仓库这个风格的图”**，所以默认不要改动样式。

## 环境

```bash
pip install -r requirements.txt   # numpy, pandas, matplotlib, openpyxl
```

需要系统装有 Arial 字体（Windows/macOS 自带；Linux 缺字体时 matplotlib 会回退并告警）。下标通过 matplotlib mathtext（`$\mathregular{...}$`）渲染，不需要安装 LaTeX。

## 推荐工作流

1. **读用户的原始数据**（仪器导出的 txt/csv/xlsx 等），弄清列含义、单位、样品名。
2. **在仓库外的工作目录**生成符合下文格式的 xlsx。不要覆盖仓库自带的 `FEDataSourceForColumns.xlsx`、`XRDDatasource.xlsx`，它们是格式样例。
3. **出图前先把整理好的表给用户确认**：电位顺序、产物对应列、FE 是否为小数、样品命名。这一步最容易出错。
4. 用 `-i` 指定数据、`-o` 指定输出（以 `/` 结尾表示目录，不存在会自动创建）：
   ```bash
   python FE_figure/FEline.py   -i <data.xlsx> -o <out_dir>/ [-e] [-cb]
   python FE_figure/FEcolumn.py -i <data.xlsx> -o <out_dir>/ [-e] [-cb] [-group]
   python XRD_figure/XRD.py     -i <data.xlsx> -o <out_dir>/ [-cb] [--xmin 10 --xmax 90] [--spacing 1.0]
   ```
   `--spacing` 是相邻样品曲线的间距系数（相对于单条曲线归一化后的峰高）。样品多、图显得挤时调小（0.6–0.8）。调大到 1.5 左右时，最上面的曲线可能碰到右上角的图例，出图后要检查。PDF 卡片区的高度由 `FIGURE_CONFIG['card_height_ratio']` 控制。
5. 用 Read 查看生成的 PNG，检查图例遮挡、Y 轴范围、标签渲染等，再交给用户。

输出同名文件已存在时，脚本会自动加 `_1`、`_2` 后缀，不会覆盖。

Windows 下控制台输出中文可能乱码或报编码错误，可设置环境变量 `PYTHONIOENCODING=utf-8`。

## 数据格式约定

### FE 数据（`FEline.py` / `FEcolumn.py`）

- 工作表 `FE`：第 1 列为 X 轴分类标签（通常是电位，如 `-1.3`），其余每列一个产物。
- **FE 值写小数**（0.72 表示 72%），脚本会乘以 100。
- **列名用化学式键名**：`H2`、`CO`、`CH4`、`C2H4`、`CH3OH`、`HCOOH`、`C2H5OH`、`CH3COOH`。这些键在 `FE_figure/FE_substance_config.py` 中映射为带下标的图例；不在配置里的列名会原样显示。新增产物时往 `SUBSTANCE_CONFIG` 里加一项。
- 全空的列会被自动跳过，所以模板里可以保留未检测到的产物列。
- 产物按**从右到左**的列顺序绘制（堆积柱中最右列在最底层）。想调整堆叠顺序就调整列顺序。
- 工作表 `SD`（加 `-e` 时必需）：与 `FE` 的形状、列名、列顺序完全一致，填标准差（同样是小数）。
- 坐标轴标题写死为 `Potential (V vs. RHE)` 和 `Faradaic Efficiency (%)`。

### XRD 数据（`XRD.py`）

- 工作表 `Sample`：列成对出现，奇数列为角度，偶数列为强度。**角度列的表头必须包含 `2 Theta`**（如 `2 Theta (°)`），否则该对列会被静默忽略。强度列的表头即图例中的样品名。
- 工作表 `Standard`：同样成对，表头含 `2 Theta` 的角度列 + 卡片强度列。强度列表头即卡片标签，用 `_` 表示下标，例如 `TiO_2-rutile PDF#01-089-8300`。
- 强度会被自动归一化，原始计数即可，无需预处理。
- Excel 中重复的表头会被 pandas 加上 `.1` 后缀并显示在图中，所以样品名和卡片名要唯一。
- 样品按列顺序自下而上堆叠。

## 样式（默认不改）

三个脚本共用：3.2 × 2.56 英寸，300 DPI，Arial 加粗，轴标签 8.5 pt，刻度 7 pt，轴线 1.5 pt；`-cb` 切换为 Wong 色盲友好配色。

- FE 脚本的样式常量分散在 `generate_figure()` 中（画布尺寸、字号、边距）以及文件顶部的 `PALETTES`、`MARKERS`。
- XRD 的样式集中在 `XRD.py` 顶部的 `FIGURE_CONFIG` 字典。

用户只是要一次性调整（如换坐标轴标题、改 Y 轴上限）时，优先在工作目录里复制一份脚本修改，不要改动仓库内的脚本，除非用户明确要求修改仓库本身的默认风格。

## 修改仓库代码时

- 源文件是 UTF-8 编码；仓库中以 LF 换行存储（Windows 上检出时可能被 Git 转成 CRLF），不要提交仅有换行符差异的改动。
- 每个脚本顶部有详细的中文文档字符串（参数、样式、版本历史），改动行为时同步更新。
- 没有自动化测试；改完用仓库自带的示例 xlsx 跑一遍三个脚本，并查看输出图片。
