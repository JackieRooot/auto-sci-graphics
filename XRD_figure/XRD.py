#!/usr/bin/env python3
# -*- coding: utf-8 -*-

r"""
XRD图谱自动绘制脚本
====================================

功能说明
--------
从结构化Excel文件中读取样品XRD数据和标准PDF卡片数据，自动生成专业级XRD对比图谱。
支持样品曲线堆叠显示、PDF卡片垂直线标注、自动归一化、多种配色方案等功能。

主要特性
--------
1. 自动读取Excel中的样品和标准卡片数据
2. 样品曲线自动堆叠，避免重叠
3. PDF卡片以垂直线形式绘制，支持基线偏移堆叠
4. 自动归一化强度，确保图谱美观
5. 支持色盲友好配色方案
6. 自动处理文件重名，避免覆盖
7. 支持LaTeX格式的PDF卡片标签
8. 高分辨率输出（300 DPI）

Excel文件格式要求
-----------------
文件必须包含两个工作表：

1. Sheet 'Sample' (样品数据)：
   - 奇数列：2 Theta 角度值
   - 偶数列：对应的强度值
   - 支持多个样品（多列数据）

2. Sheet 'Standard' (标准卡片数据)：
   - 奇数列：2 Theta 角度值
   - 偶数列：PDF卡片强度值
   - 支持多个PDF卡片（多列数据）

示例结构：
    Sample工作表：
    | 2 Theta | Sample1 | 2 Theta | Sample2 | ...
    | 10.0    | 100     | 10.0    | 150     | ...
    | 10.1    | 120     | 10.1    | 160     | ...

使用方法
--------
基本用法：
    python XRD.py -i input.xlsx

指定输出路径：
    python XRD.py -i input.xlsx -o output.png
    python XRD.py -i input.xlsx -o D:/results/          # 保存到目录

使用色盲友好配色：
    python XRD.py -i input.xlsx -cb

自定义X轴范围和高度：
    python XRD.py -i input.xlsx --xmin 10 --xmax 80 --height 150

完整参数示例：
    python XRD.py -i data.xlsx -o ./plots/ -cb --xmin 5 --xmax 90 --height 120

命令行参数
----------
-i, --input      : 输入Excel文件路径（默认：脚本所在目录/XRDDatasource.xlsx）
-o, --output     : 输出图片路径（支持文件路径或目录路径）
-cb              : 使用色盲友好配色方案（Wong配色）
--xmin, --xmax   : X轴范围（默认：10-90）
--height         : 样品曲线基础高度（默认：100）
-h, --help       : 显示帮助信息

输出说明
--------
- 输出格式：PNG（300 DPI高分辨率）
- 图片尺寸：3.2 × 2.56 英寸（适合论文插入）
- 字体：Arial（加粗）
- 自动重名处理：如文件已存在，自动添加 _1, _2 等后缀

配置文件说明
------------
脚本内置 FIGURE_CONFIG 字典，可自定义以下参数：
- figsize          : 图片尺寸（英寸）
- dpi              : 分辨率
- font_family      : 字体类型
- label_fontsize   : 轴标签字体大小
- tick_fontsize    : 刻度字体大小
- legend_fontsize  : 图例字体大小
- line_width       : 线条宽度
- stem_width       : PDF垂直线宽度
- card_height_ratio: PDF卡片高度比例
- left/right/bottom/top_margin : 边距设置

依赖库
------
- numpy
- pandas
- matplotlib
- argparse
- pathlib

版本历史
--------
v2.1 (2026-06-07)
- 改进：默认输入路径改为脚本所在目录，增强可移植性
- 改进：使用 vlines/hlines 替代 stem 绘图，提高兼容性
- 新增：支持目录路径作为 -o 参数
- 新增：自动处理文件重名，避免覆盖

v2.0 (2026-06-07)
- 新增：PDF卡片基线仅在必要时显示
- 优化：标签格式化和配色方案

作者
----
XRD Data Processing Tool
最后更新：2026-06-07
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import argparse
from pathlib import Path
import re
import os

# ── 配色方案字典 ──
PALETTES = {
    'default': [  # Matplotlib 默认配色
        '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
        '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'
    ],
    'cb': [       # 色盲友好配色 (Wong 8色方案)
        '#000000',  # 黑色
        '#E69F00',  # 橙色
        '#56B4E9',  # 天蓝
        '#009E73',  # 绿色
        '#F0E442',  # 黄色
        '#0072B2',  # 蓝色
        '#D55E00',  # 朱红
        '#CC79A7'   # 紫红
    ]
}

# ── 图形格式参数 ──
FIGURE_CONFIG = {
    'figsize': (3.2, 2.56),
    'dpi': 300,
    'font_family': 'Arial',
    'label_fontsize': 8.5,
    'tick_fontsize': 7,
    'legend_fontsize': 6,
    'line_width': 1.5,
    'tick_width': 1.5,
    'stem_width': 1.5,
    'card_height_ratio': 0.7,  # 卡片高度比例（相对于样品线高度）
    # 边距配置 (相对坐标，范围0-1)
    'left_margin': 0.18,    # 左边界
    'right_margin': 0.95,   # 右边界
    'bottom_margin': 0.18,  # 下边界
    'top_margin': 0.95,     # 上边界
}

def setup_plot_style():
    """设置全局绘图样式"""
    plt.rcParams['font.family'] = FIGURE_CONFIG['font_family']
    plt.rcParams['font.weight'] = 'bold'
    plt.rcParams['axes.linewidth'] = FIGURE_CONFIG['line_width']
    plt.rcParams['xtick.major.width'] = FIGURE_CONFIG['tick_width']
    plt.rcParams['ytick.major.width'] = FIGURE_CONFIG['tick_width']
    plt.rcParams['xtick.labelsize'] = FIGURE_CONFIG['tick_fontsize']
    plt.rcParams['ytick.labelsize'] = FIGURE_CONFIG['tick_fontsize']
    plt.rcParams['axes.labelsize'] = FIGURE_CONFIG['label_fontsize']
    plt.rcParams['axes.labelweight'] = 'bold'

def get_unique_output_path(output_path):
    """
    生成唯一的输出路径，如果文件已存在则自动添加序号
    
    参数:
        output_path: 原始输出路径 (Path 对象或字符串)
    
    返回:
        唯一的输出路径 (Path 对象)
    """
    output_path = Path(output_path)
    
    # 如果文件不存在，直接返回原路径
    if not output_path.exists():
        return output_path
    
    # 分离文件名的各个部分
    parent_dir = output_path.parent
    stem = output_path.stem  # 文件名（不含扩展名）
    suffix = output_path.suffix  # 扩展名
    
    # 查找可用的序号
    counter = 1
    while True:
        new_stem = f"{stem}_{counter}"
        new_path = parent_dir / f"{new_stem}{suffix}"
        if not new_path.exists():
            print(f"文件 {output_path.name} 已存在，将保存为 {new_path.name}")
            return new_path
        counter += 1

def read_xrd_data(excel_path):
    """
    从Excel文件中读取XRD数据
    
    参数:
        excel_path: Excel文件路径
    
    返回:
        samples: 样品数据字典 {样品名: (角度数组, 强度数组)}
        standards: 标准卡片数据字典 {卡片名: (角度数组, 强度数组)}
    """
    xl = pd.ExcelFile(excel_path)
    
    # 读取样品数据 (工作表1)
    df_samples = pd.read_excel(xl, sheet_name='Sample')
    samples = {}
    sample_columns = df_samples.columns.tolist()
    
    # 奇数列为角度，偶数列为强度
    for i in range(1, len(sample_columns), 2):
        angle_col = sample_columns[i-1]
        intensity_col = sample_columns[i]
        
        if '2 Theta' in angle_col:
            angles = df_samples[angle_col].dropna().values
            intensities = df_samples[intensity_col].dropna().values
            samples[intensity_col] = (angles, intensities)
    
    # 读取标准卡片数据 (工作表2)
    df_standards = pd.read_excel(xl, sheet_name='Standard')
    standards = {}
    standard_columns = df_standards.columns.tolist()
    
    for i in range(1, len(standard_columns), 2):
        angle_col = standard_columns[i-1]
        card_col = standard_columns[i]
        
        if '2 Theta' in angle_col:
            angles = df_standards[angle_col].dropna().values
            intensities = df_standards[card_col].dropna().values
            standards[card_col] = (angles, intensities)
    
    return samples, standards

def normalize_intensity(intensities, target_max=100):
    """将强度归一化到指定范围"""
    max_int = np.max(intensities)
    if max_int > 0:
        return intensities / max_int * target_max
    return intensities

def format_pdf_label(text):
    r"""
    格式化 PDF 卡片标签为 LaTeX 格式
    例如: 'TiO_2-Rutile PDF#01-089-8300' -> '$\mathregular{TiO_2-Rutile PDF\#01-089-8300}$'
    """
    # 转义特殊字符
    # # 需要转义为 \#
    escaped_text = text.replace('#', r'\#')
    # 其他可能需要转义的字符
    escaped_text = escaped_text.replace('%', r'\%')
    escaped_text = escaped_text.replace('&', r'\&')
    
    # 替换 - 为 \text{-}
    escaped_text = escaped_text.replace('-', r'\text{-}')
    # 将空格替换为 \ （反斜杠+空格）
    escaped_text = escaped_text.replace(' ', r'\ ')
    # escaped_text = escaped_text.replace('_', r'\_')  # 保留下划线功能？如果要保留下划线功能，不要转义
    
    # 用 \mathregular{} 包裹，保持正体字
    return f'$\\mathregular{{{escaped_text}}}$'

def plot_xrd(samples, standards, output_path, use_cb_palette=False, 
             x_range=(10, 90), normalize=True, base_height=100):
    """
    绘制XRD图谱
    
    参数:
        samples: 样品数据字典
        standards: 标准卡片数据字典
        output_path: 输出文件路径
        use_cb_palette: 是否使用色盲友好配色
        x_range: X轴范围 (min, max)
        normalize: 是否归一化强度
        base_height: 样品曲线的基础高度
    """
    setup_plot_style()
    
    # 处理输出路径，避免覆盖
    output_path = get_unique_output_path(output_path)
    # 确保输出目录存在
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # 选择配色方案
    palette = PALETTES['cb'] if use_cb_palette else PALETTES['default']
    
    # 创建图形
    fig = plt.figure(figsize=FIGURE_CONFIG['figsize'], dpi=FIGURE_CONFIG['dpi'])
    
    # 应用自定义边距
    plt.subplots_adjust(left=FIGURE_CONFIG['left_margin'], 
                        right=FIGURE_CONFIG['right_margin'],
                        bottom=FIGURE_CONFIG['bottom_margin'], 
                        top=FIGURE_CONFIG['top_margin'])
    
    # 打印绘图区域参数
    # print(f"绘图区域边距设置:")
    # print(f"  - 左边界: {FIGURE_CONFIG['left_margin']}")
    # print(f"  - 右边界: {FIGURE_CONFIG['right_margin']}")
    # print(f"  - 下边界: {FIGURE_CONFIG['bottom_margin']}")
    # print(f"  - 上边界: {FIGURE_CONFIG['top_margin']}")

        
    # 主坐标轴（用于样品曲线）
    # ax_main = fig.add_axes([0.12, 0.12, 0.83, 0.83])
    ax_main = fig.add_subplot(111)
    # 处理样品曲线
    sample_curves = []
    color_idx = 0
    
    for name, (angles, intensities) in samples.items():
        if normalize:
            intensities = normalize_intensity(intensities, target_max=base_height)
        sample_curves.append({
            'angles': angles, 'intensities': intensities, 
            'name': name, 'color': palette[color_idx % len(palette)]
        })
        color_idx += 1
    
    # 绘制样品曲线（堆叠显示）
    num_samples = len(sample_curves)
    num_standards = len(standards)
    sample_offsets = [base_height * i * 1 for i in range(num_samples)]  # 样品之间重叠少一些
    PDF_offsets = num_standards * base_height * FIGURE_CONFIG['card_height_ratio'] #这是为每张PDF卡片预留单样品指定倍数的高度。
    
    for i, curve in enumerate(sample_curves):
        y_values = curve['intensities'] + sample_offsets[i] + PDF_offsets
        ax_main.plot(curve['angles'], y_values, '-', 
                    color=curve['color'], linewidth=FIGURE_CONFIG['line_width'], 
                    label=curve['name'])
    
    # 处理标准卡片（在主坐标轴内通过offset堆叠）
    if num_standards > 0:
        for i, (name, (angles, intensities)) in enumerate(standards.items()):
            if normalize:
                intensities = normalize_intensity(intensities, target_max=base_height * FIGURE_CONFIG['card_height_ratio'])
            
            # 计算当前PDF卡片的Y轴偏移量
            card_offset = i * base_height * FIGURE_CONFIG['card_height_ratio']
            
            color = palette[color_idx % len(palette)]
            color_idx += 1
            
            # --- 核心改动：使用 vlines 和 hlines 替代 stem ---
            # 1. 绘制竖直线 (从基线 card_offset 画到峰顶 intensities + card_offset)
            # vlines(x, ymin, ymax) 不会自动连线，只画孤立的垂直线
            ax_main.vlines(
                angles, 
                card_offset, 
                intensities + card_offset, 
                colors=color, 
                linewidth=FIGURE_CONFIG['stem_width'], 
                alpha=0.8
            )
            
            # 2. 绘制水平基线 (只有 i > 0 时才画)
            if i > 0:
                # hlines(y, xmin, xmax) 画一条纯水平线
                ax_main.hlines(
                    card_offset, 
                    x_range[0], 
                    x_range[1], 
                    colors=color, 
                    linewidth=FIGURE_CONFIG['line_width']*2/3, # 可以用stem_width或line_width
                    alpha=0.8
                )

            # --- 标签部分保持不变 ---
            label_x = x_range[1] - 1  
            label_y = card_offset + base_height * FIGURE_CONFIG['card_height_ratio'] * 5 / 7
            #5/7意味着PDF卡对应的矩形区域约为70%的高度处
            ax_main.text(label_x, label_y, format_pdf_label(name), 
                        fontsize=FIGURE_CONFIG['tick_fontsize']-1,
                        fontweight='bold', color=color, 
                        ha='right', va='center')


    # 设置主坐标轴
    ax_main.set_xlim(x_range)
    ax_main.set_xlabel('2 Theta (degree)', fontweight='bold')
    ax_main.set_ylabel('Intensity (a.u.)', fontweight='bold')
    ax_main.set_yticks([])
    
    # 计算主坐标轴的Y轴范围
    if sample_curves:
        max_y = max([np.max(curve['intensities']) + sample_offsets[i] + PDF_offsets
                    for i, curve in enumerate(sample_curves)])
        ax_main.set_ylim(0, max_y + base_height * 0.2)
    
    # 添加图例
    if sample_curves:
        ax_main.legend(loc='upper right', fontsize=FIGURE_CONFIG['legend_fontsize'], 
                      frameon=False)
    
    # 设置主坐标轴spines
    for spine in ax_main.spines.values():
        spine.set_linewidth(FIGURE_CONFIG['line_width'])

    
    # 保存图片
    # fig.savefig(output_path, dpi=FIGURE_CONFIG['dpi'], 
    #             bbox_inches='tight')
    fig.savefig(output_path, dpi=FIGURE_CONFIG['dpi'])
    plt.close()
    
    print(f'Done! 图谱已保存至: {output_path}')

def main():
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_input_path = os.path.join(script_dir, 'XRDDatasource.xlsx')
    
    parser = argparse.ArgumentParser(description='XRD图谱自动绘制工具')
    parser.add_argument('-cb', action='store_true', 
                       help='使用色盲友好配色方案')
    parser.add_argument('-i', '--input', type=str, 
                       default=default_input_path,
                       help='输入Excel文件路径')
    parser.add_argument('-o', '--output', type=str, default=None,
                       help='输出图片路径（支持文件路径或目录路径）')
    parser.add_argument('--xmin', type=float, default=10, help='X轴最小值')
    parser.add_argument('--xmax', type=float, default=90, help='X轴最大值')
    parser.add_argument('--height', type=float, default=100, help='样品曲线基础高度')
    
    args = parser.parse_args()
    
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误: 文件不存在 - {input_path}")
        return
    
    # 处理输出路径
    if args.output:
        output_path = Path(args.output)
        
        # 判断是否为目录（以分隔符结尾或已存在的目录）
        if args.output.endswith(os.sep) or args.output.endswith('/') or output_path.is_dir():
            # 如果是目录，在目录下创建默认文件名
            output_path = output_path / 'XRD_plot.png'
        # 如果父目录不存在，尝试创建
        output_path.parent.mkdir(parents=True, exist_ok=True)
    else:
        # 默认输出到输入文件所在目录
        output_path = input_path.parent / 'XRD_plot.png'
    
    # 读取并绘制
    print("正在读取数据...")
    samples, standards = read_xrd_data(input_path)
    print(f"读取到 {len(samples)} 个样品: {list(samples.keys())}")
    print(f"读取到 {len(standards)} 个标准卡片: {list(standards.keys())}")
    
    print("正在绘制图谱...")
    plot_xrd(samples, standards, output_path, args.cb,
            (args.xmin, args.xmax), True, args.height)

if __name__ == '__main__':
    main()