# auto-sci-graphics

用 Python 一条命令画出风格统一、可直接放进论文的电催化数据图。不需要打开 Origin，也不需要一张张手动调格式。

目前支持三种图：

| 法拉第效率折线图 | 法拉第效率堆积柱状图 | 法拉第效率分组柱状图 | XRD 谱图 |
|:---:|:---:|:---:|:---:|
| ![FE 折线图](docs/images/fe_line.png) | ![FE 堆积柱状图](docs/images/fe_stacked.png) | ![FE 分组柱状图](docs/images/fe_grouped.png) | ![XRD 谱图](docs/images/xrd.png) |
| `FEline.py` | `FEcolumn.py` | `FEcolumn.py -group` | `XRD.py` |

所有图的风格一致：

- 3.2 × 2.56 英寸，正好是期刊单栏宽度
- 300 DPI PNG
- Arial 加粗字体
- 化学式自动显示下标（H₂、C₂H₄、TiO₂）
- 可一键切换色盲友好配色

## 安装

需要 Python 3.9 或更高版本。

```bash
git clone https://github.com/JackieRooot/auto-sci-graphics
cd auto-sci-graphics
pip install -r requirements.txt
```

## 快速上手

仓库自带示例数据，可以直接运行看看效果：

```bash
python FE_figure/FEline.py -e          # 折线图，带误差棒
python FE_figure/FEcolumn.py -e        # 堆积柱状图，带误差线
python XRD_figure/XRD.py               # XRD 谱图
```

不加 `-o` 时，图片保存在数据文件所在的目录。

画你自己的数据：照着示例 Excel 的格式准备好文件，用 `-i` 指定它：

```bash
python FE_figure/FEline.py -i D:/data/my_fe.xlsx -e -o D:/figures/
python XRD_figure/XRD.py   -i D:/data/my_xrd.xlsx -o D:/figures/ --xmin 20 --xmax 80
```

## 准备数据

最省事的做法：复制仓库里的示例 Excel，然后把自己的数据填进去。

### 法拉第效率

示例文件：[`FE_figure/FEDataSourceForColumns.xlsx`](FE_figure/FEDataSourceForColumns.xlsx)

工作表 `FE`：

| Potential | H2 | CO | CH4 | C2H4 | … |
|---|---|---|---|---|---|
| -1.3 | 0.720 | | 0.227 | 0.053 | |
| -1.4 | 0.661 | | 0.278 | 0.061 | |

- 第一列是横坐标（通常是电位），其余每列一个产物。
- **法拉第效率填小数**，比如 0.72 表示 72%。
- 列名用化学式，比如 `H2`、`C2H4`，图例里会自动显示成下标形式。支持的产物见 [`FE_substance_config.py`](FE_figure/FE_substance_config.py)。要新增产物或改显示名称（比如把 `HCOOH` 显示成 `Formate`），就编辑这个文件。
- 没有数据的列留空即可，画图时会自动跳过。

工作表 `SD`（只在画误差棒时需要）：格式和 `FE` 完全一样，填标准差。

### XRD

示例文件：[`XRD_figure/XRDDatasource.xlsx`](XRD_figure/XRDDatasource.xlsx)

工作表 `Sample`（样品）和 `Standard`（标准 PDF 卡片）格式相同：每两列为一组，前一列是角度，后一列是强度。

| 2 Theta (°) | SnTi750 | 2 Theta (°) | SnTi800 | … |
|---|---|---|---|---|
| 5.00 | 2721 | 5.00 | 2779 | |

- **角度列的表头必须包含 `2 Theta`**，否则这一组数据会被忽略。
- 强度列的表头就是图例中的样品名或卡片名。卡片名里用 `_` 表示下标，比如 `TiO_2-rutile PDF#01-089-8300` 会显示为 TiO₂-rutile PDF#01-089-8300。
- 强度直接填原始计数，脚本会自动归一化。
- 每个样品名和卡片名都要不同，重复的名字会在图中多出 `.1` 后缀。

## 命令行参数

| 参数 | 适用脚本 | 作用 |
|---|---|---|
| `-i, --input` | 全部 | 输入的 Excel 文件。不填就使用脚本旁边的示例文件 |
| `-o, --output` | 全部 | 输出路径。以 `/` 结尾表示目录（不存在会自动创建），否则当作文件名 |
| `-cb` | 全部 | 使用色盲友好配色（Wong 配色方案） |
| `-e` | FE 两个脚本 | 画误差棒，需要 `SD` 工作表 |
| `-group` | `FEcolumn.py` | 画分组柱状图，不加就是堆积柱状图 |
| `--xmin`, `--xmax` | `XRD.py` | 横坐标范围，默认 10–90 |
| `--spacing` | `XRD.py` | 相邻样品曲线的间距，默认 1.0。调小（如 0.6）曲线更紧凑、峰更高；调大（如 1.4）曲线更松散 |

如果输出文件已存在，新图会自动加上 `_1`、`_2` 后缀，不会覆盖旧图。

## 搭配 Claude Code 使用

仓库里有一份 [`CLAUDE.md`](CLAUDE.md)，写明了数据格式、出图流程和注意事项。用 [Claude Code](https://claude.com/claude-code) 打开仓库后，可以直接把原始数据交给它，比如：

> 我的 GC 定量结果在 `D:\data\run3\`，帮我按仓库风格画一张带误差线的 FE 堆积柱状图。先把整理好的数据给我确认，再出图。

它会把数据整理成上面的格式，调用脚本出图并检查结果。你不需要自己调整数据格式，也不需要了解样式是怎么实现的。

## 自定义样式

- **FE 图**：配色和标记点形状在脚本顶部的 `PALETTES`、`MARKERS`；画布尺寸、字号、边距在 `generate_figure()` 函数里。
- **XRD 图**：样式参数都在 `XRD.py` 顶部的 `FIGURE_CONFIG` 里。

每个脚本开头都有详细的中文说明，列出了全部样式参数。

## License

[AGPL-3.0](LICENSE)
