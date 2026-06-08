#!/usr/bin/env python3
# -*- coding: utf-8 -*-

r"""
================================================================================
脚本名称: FEcolumn.py
功能描述: 读取 Excel 中的法拉第效率数据，生成带误差线的柱状图（支持堆积/分组），
         支持配置文件管理物质名称和LaTeX格式
         可选从SD工作表读取标准差数据绘制误差线
作者: 
创建日期: 
最后更新: 2026-06-07
================================================================================

【输入要求】
    数据文件: FEDataSourceForColumns.xlsx
    文件路径: 脚本所在目录（自动定位，无需手动指定路径）
    数据格式:
        - 工作表 'FE': 提供电位和各个物质的法拉第效率数据（必需）
        - 工作表 'SD': 提供电位和各个物质的标准差数据（使用 -e 参数时必需）
        - 第1列: 分类标签 (如电位值)，作为 X 轴刻度
        - 后续列: 各物质的法拉第效率数据 (原始值为小数，脚本会自动乘以100转换为百分比)
        - 程序会自动跳过全空列，并按从右到左的顺序绘制
        - SD工作表的列顺序应与FE工作表一致（使用 -e 参数时）

【物质名称配置】
    配置文件: FE_substance_config.py (可选，不存在时使用默认配置)
    配置内容:
        - display_name: 普通显示名称 (如 'H₂', 'CH₄')
        - latex: LaTeX 格式标签 (如 r'$\mathregular{H_2}$')
    
    支持的物质及 LaTeX 下标渲染:
        H2, CO, CH4, C2H4, CH3OH, HCOOH, C2H5OH, CH3COOH
    示例: CH4 -> $\mathregular{CH_4}$ (在图例中渲染为 CH₄)

【输出文件】
    默认文件名: 
        - 无误差线堆积柱状图: FEStackedBar.png
        - 带误差线堆积柱状图: FEStackedBar_with_Error.png
        - 无误差线分组柱状图: FEGroupedBar.png
        - 带误差线分组柱状图: FEGroupedBar_with_Error.png
    分辨率: 300 DPI
    尺寸: 3.2 × 2.56 英寸
    自动重名处理: 如文件已存在，自动添加 _1, _2 等后缀

【使用示例】
    # 基础用法 (默认配色，堆积柱状图，无误差线)
    python FEcolumn.py
    
    # 带误差线的堆积柱状图
    python FEcolumn.py -e
    
    # 使用 -o 参数指定输出文件路径
    python FEcolumn.py -o D:/MyFigures/my_chart.png -e
    
    # 使用色盲友好配色 + 误差线
    python FEcolumn.py -cb -e
    
    # 使用分组柱状图 + 误差线
    python FEcolumn.py -group -e
    
    # 分组柱状图 + 误差线 + 指定输出目录
    python FEcolumn.py -group -e -o D:/MyFigures/

【命令行参数】
    -o, --output    可选，输出图片路径（支持文件路径或目录路径）
                    - 如果是目录：在该目录下使用默认文件名
                    - 如果是文件：直接保存为该文件
                    - 不指定：默认保存到数据文件所在目录
    -cb             可选，启用色盲友好配色 (Wong 8色方案)，不加则使用 Matplotlib 默认配色
    -group          可选，生成分组并列柱状图，不加则默认生成堆积柱状图
    -e              可选，绘制误差线（需要SD工作表提供标准差数据），不加则只读取FE工作表

【图表样式说明】
    堆积柱状图:
        - 柱宽: 0.5
        - 图例位置: 图表上方内侧居中，横向单行排列
        - 图例字号: 8
    
    分组柱状图:
        - 柱宽: 0.25
        - 图例位置: 自动调整最佳位置，竖排排列
        - 图例字号: 7
    
    误差线样式（使用 -e 参数时）:
        - 颜色: 黑色
        - 线宽: 1 pt
        - 端点装饰capsize=1.5,capthick=0.75
    
    共同样式:
        - 字体: Arial
        - 柱边缘: 黑色边框
        - X轴标签: "Potential (V vs. RHE)" (粗体)
        - Y轴标签: "Faradaic Efficiency (%)" (粗体)
        - 轴标签字号: 8.5
        - 刻度值字号: 7
        - 轴线宽度: 1.5 pt
        - Y轴范围: 底部固定为0，顶部自动留空间

【配置文件说明】
    创建 FE_substance_config.py 文件可以自定义物质显示名称和LaTeX格式：
    
    SUBSTANCE_CONFIG = {
        'H2': {'display_name': 'H₂', 'latex': r'$\mathregular{H_2}$'},
        'HCOOH': {'display_name': 'Formate', 'latex': r'$\mathregular{HCOOH}$'},
        # 添加新物质
        'C3H6': {'display_name': 'C₃H₆', 'latex': r'$\mathregular{C_3H_6}$'}
    }
    
    def get_latex_label(col_name): ...
    def get_display_name(col_name): ...

【版本历史】
    v3.1 - 新增 -e 参数控制是否绘制误差线
         - 无 -e 参数时仅读取 FE 工作表
         - 有 -e 参数时同时读取 FE 和 SD 工作表并绘制误差线
    v3.0 - 新增标准差支持：从SD工作表读取误差数据
    v2.3 - 改进输入路径处理：数据文件自动定位到脚本同目录

================================================================================
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
import argparse
import os
from pathlib import Path
import sys
import numpy as np
sys.dont_write_bytecode = True


# 导入配置文件
try:
    from FE_substance_config import SUBSTANCE_CONFIG, get_latex_label, get_display_name
    print("成功加载配置文件: FE_substance_config.py")
except ImportError:
    # 如果配置文件不存在，使用默认配置（向后兼容）
    print("警告: 未找到 FE_substance_config.py，使用默认配置")
    
    # 默认配置字典
    SUBSTANCE_CONFIG = {
        'H2': {'display_name': 'H₂', 'latex': r'$\mathregular{H_2}$'},
        'CO': {'display_name': 'CO', 'latex': r'$\mathregular{CO}$'},
        'CH4': {'display_name': 'CH₄', 'latex': r'$\mathregular{CH_4}$'},
        'C2H4': {'display_name': 'C₂H₄', 'latex': r'$\mathregular{C_2H_4}$'},
        'CH3OH': {'display_name': 'CH₃OH', 'latex': r'$\mathregular{CH_3OH}$'},
        'HCOOH': {'display_name': 'HCOOH', 'latex': r'$\mathregular{HCOOH}$'},
        'C2H5OH': {'display_name': 'C₂H₅OH', 'latex': r'$\mathregular{C_2H_5OH}$'},
        'CH3COOH': {'display_name': 'CH₃COOH', 'latex': r'$\mathregular{CH_3COOH}$'}
    }
    
    def get_latex_label(col_name):
        """获取 LaTeX 格式标签（fallback）"""
        return SUBSTANCE_CONFIG.get(col_name, {}).get('latex', col_name)
    
    def get_display_name(col_name):
        """获取显示名称（fallback）"""
        return SUBSTANCE_CONFIG.get(col_name, {}).get('display_name', col_name)

# ── 全局字体 ──
mpl.rcParams['font.family'] = 'Arial'

# ── 配色方案字典 ──
PALETTES = {
    'default': [  # Matplotlib 默认配色 (普通配色)
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

# ── 默认文件名 ──
DEFAULT_STACKED_FILENAME = 'FEStackedBar.png'
DEFAULT_STACKED_FILENAME_WITH_ERROR = 'FEStackedBar_with_Error.png'
DEFAULT_GROUPED_FILENAME = 'FEGroupedBar.png'
DEFAULT_GROUPED_FILENAME_WITH_ERROR = 'FEGroupedBar_with_Error.png'

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

def load_fe_data(excel_path):
    """
    仅加载FE工作表数据
    
    参数:
        excel_path: Excel文件路径
    
    返回:
        df_fe: FE数据DataFrame
        categories: 类别标签列表
        non_empty_cols: 非空列列表（从右到左排序）
    """
    try:
        df_fe = pd.read_excel(excel_path, sheet_name='FE')
        print("成功读取工作表: FE")
    except Exception as e:
        print(f"错误: 无法读取FE工作表 - {e}")
        print("请确保Excel文件中包含名为 'FE' 的工作表")
        return None, None, None
    
    # 获取类别列（第1列）
    categories = df_fe.iloc[:, 0].astype(str).tolist()
    
    # 筛选有数据的列（剔除全空列），并保证顺序为表格从右到左
    non_empty_cols = [col for col in df_fe.columns[1:] if df_fe[col].notna().any()]
    non_empty_cols = non_empty_cols[::-1]  # 从右到左排序
    
    print(f"找到 {len(non_empty_cols)} 个有数据的物质列: {non_empty_cols}")
    
    return df_fe, categories, non_empty_cols

def load_fe_sd_data(excel_path):
    """
    加载FE和SD工作表数据（用于误差线）
    
    参数:
        excel_path: Excel文件路径
    
    返回:
        df_fe: FE数据DataFrame
        df_sd: SD数据DataFrame
        categories: 类别标签列表
        non_empty_cols: 非空列列表（从右到左排序）
    """
    # 读取两个工作表
    try:
        df_fe = pd.read_excel(excel_path, sheet_name='FE')
        df_sd = pd.read_excel(excel_path, sheet_name='SD')
        print("成功读取工作表: FE 和 SD")
    except Exception as e:
        print(f"错误: 无法读取工作表 - {e}")
        print("请确保Excel文件中包含名为 'FE' 和 'SD' 的工作表")
        return None, None, None, None
    
    # 验证数据形状
    if df_fe.shape != df_sd.shape:
        print(f"警告: FE工作表形状 {df_fe.shape} 与 SD工作表形状 {df_sd.shape} 不一致")
        print("请确保两个工作表具有相同的行列结构")
    
    # 获取类别列（第1列）
    categories = df_fe.iloc[:, 0].astype(str).tolist()
    
    # 筛选有数据的列（剔除全空列），并保证顺序为表格从右到左
    non_empty_cols = [col for col in df_fe.columns[1:] if df_fe[col].notna().any()]
    non_empty_cols = non_empty_cols[::-1]  # 从右到左排序
    
    print(f"找到 {len(non_empty_cols)} 个有数据的物质列: {non_empty_cols}")
    
    return df_fe, df_sd, categories, non_empty_cols

def prepare_error_data(df_sd, col, categories):
    """
    准备指定列的误差数据（标准差）
    
    参数:
        df_sd: SD数据DataFrame
        col: 列名
        categories: 类别标签列表（用于验证长度）
    
    返回:
        误差数组（已转换为百分比），如果数据无效则返回None
    """
    if col not in df_sd.columns:
        print(f"警告: 列 '{col}' 在SD工作表中不存在")
        return None
    
    # 获取标准差数据，转换为百分比
    errors = df_sd[col].values * 100
    
    # 检查是否有缺失值
    if pd.isna(errors).any():
        print(f"警告: 列 '{col}' 的标准差数据包含缺失值，将使用0替代")
        errors = np.nan_to_num(errors, nan=0.0)
    
    # 确保误差值非负
    errors = np.abs(errors)
    
    return errors

def generate_stacked_bar_without_error(ax, categories, non_empty_cols, df_fe, current_palette):
    """
    生成无误差线的堆积柱状图
    """
    x_pos = range(len(categories))
    bottom = None
    
    for i, col in enumerate(non_empty_cols):
        values = df_fe[col].values * 100
        ax.bar(
            x_pos,
            values,
            bottom=bottom,
            edgecolor='black',
            label=get_latex_label(col),
            width=0.5,
            color=current_palette[i % len(current_palette)]
        )
        bottom = values if bottom is None else bottom + values
    
    # Y轴范围设置
    y_max = bottom.max() if bottom is not None else 100
    ax.set_ylim(bottom=0, top=y_max * 1.15)
    
    # 图例设置（堆积图：上方内侧居中，横排）
    legend = ax.legend(
        loc='upper center',
        bbox_to_anchor=(0.5, 0.99),
        ncol=len(non_empty_cols),
        frameon=False,
        handlelength=1,
        handletextpad=0.4,
        columnspacing=1,
        prop={'weight': 'bold', 'size': 8},
        borderaxespad=0
    )

def generate_stacked_bar_with_error(ax, categories, non_empty_cols, df_fe, df_sd, current_palette):
    """
    生成带误差线的堆积柱状图
    
    注意: 堆积柱状图的误差线通常绘制在每个堆积段的顶部，
         但为了更好地可视化，这里为每个物质绘制单独的误差线
    """
    x_pos = range(len(categories))
    bottom = np.zeros(len(categories))
    error_data = []      # 存储每个物质的误差数据
    
    for i, col in enumerate(non_empty_cols):
        values = df_fe[col].values * 100
        
        # 获取误差数据
        errors = prepare_error_data(df_sd, col, categories)
        if errors is not None:
            error_data.append((i, errors, bottom.copy()))
        
        # 绘制堆积柱状图
        ax.bar(
            x_pos,
            values,
            bottom=bottom,
            edgecolor='black',
            label=get_latex_label(col),
            width=0.5,
            color=current_palette[i % len(current_palette)]
        )
        
        # 更新bottom值
        bottom = bottom + values
    
    # 添加误差线（在每个堆积段的顶部）
    for i, errors, bottom_values in error_data:
        # 计算误差线的位置（底部值 + 当前段的值）
        current_values = df_fe[non_empty_cols[i]].values * 100
        y_positions = bottom_values + current_values
        
        # 绘制误差线
        ax.errorbar(
            x_pos, 
            y_positions, 
            yerr=errors,
            fmt='none',           # 不显示数据点
            ecolor='black',       # 误差线颜色
            elinewidth=1,         # 误差线宽度
            capsize=1.5,            # 无端点装饰
            capthick=0.75
        )
    
    # Y轴范围设置
    y_max = bottom.max() if len(bottom) > 0 else 100
    ax.set_ylim(bottom=0, top=y_max * 1.15)
    
    # 图例设置（堆积图：上方内侧居中，横排）
    legend = ax.legend(
        loc='upper center',
        bbox_to_anchor=(0.5, 0.99),
        ncol=len(non_empty_cols),
        frameon=False,
        handlelength=1,
        handletextpad=0.4,
        columnspacing=1,
        prop={'weight': 'bold', 'size': 8},
        borderaxespad=0
    )


def generate_grouped_bar_without_error(ax, categories, non_empty_cols, df_fe, current_palette):
    """
    生成无误差线的分组并列柱状图
    """
    n_groups = len(categories)
    n_bars = len(non_empty_cols)
    
    bar_width = 0.25
    group_width = n_bars * bar_width
    
    # 绘制柱状图
    for i, col in enumerate(non_empty_cols):
        values = df_fe[col].values * 100
        x_pos = [x - group_width/2 + bar_width/2 + i * bar_width for x in range(n_groups)]
        
        ax.bar(
            x_pos,
            values,
            width=bar_width,
            edgecolor='black',
            color=current_palette[i % len(current_palette)],
            label=get_latex_label(col)
        )
    
    # Y轴范围设置
    y_max = df_fe[non_empty_cols].max().max() * 100
    ax.set_ylim(bottom=0, top=y_max * 1.1)
    
    # 图例设置（分组图：自动调整位置，竖排）
    legend = ax.legend(
        loc='best',
        ncol=1,
        frameon=False,
        handlelength=1,
        handletextpad=0.4,
        borderaxespad=0,
        prop={'weight': 'bold', 'size': 7}
    )

def generate_grouped_bar_with_error(ax, categories, non_empty_cols, df_fe, df_sd, current_palette):
    """
    生成带误差线的分组并列柱状图
    """
    n_groups = len(categories)
    n_bars = len(non_empty_cols)
    
    bar_width = 0.25
    group_width = n_bars * bar_width
    
    # 绘制柱状图并添加误差线
    for i, col in enumerate(non_empty_cols):
        values = df_fe[col].values * 100
        x_pos = [x - group_width/2 + bar_width/2 + i * bar_width for x in range(n_groups)]
        
        # 绘制柱子
        ax.bar(
            x_pos,
            values,
            width=bar_width,
            edgecolor='black',
            color=current_palette[i % len(current_palette)],
            label=get_latex_label(col)
        )
        
        # 获取误差数据并添加误差线
        errors = prepare_error_data(df_sd, col, categories)
        if errors is not None:
            ax.errorbar(
                x_pos, 
                values, 
                yerr=errors,
                fmt='none',           # 不显示数据点
                ecolor='black',       # 误差线颜色
                elinewidth=1,         # 误差线宽度
                capsize=1.5,            # 无端点装饰
                capthick=0.75
            )
    
    # Y轴范围设置
    y_max = df_fe[non_empty_cols].max().max() * 100
    ax.set_ylim(bottom=0, top=y_max * 1.1)
    
    # 图例设置（分组图：自动调整位置，竖排）
    legend = ax.legend(
        loc='best',
        ncol=1,
        frameon=False,
        handlelength=1,
        handletextpad=0.4,
        borderaxespad=0,
        prop={'weight': 'bold', 'size': 7}
    )


def generate_figure(output_path=None, use_cb=False, use_group=False, use_error=False):
    """
    读取数据并生成柱状图，保存到指定路径
    
    参数:
        output_path: 输出路径（支持文件路径或目录路径）
        use_cb: 是否使用色盲友好配色
        use_group: 是否使用分组柱状图（False为堆积柱状图）
        use_error: 是否绘制误差线（需要SD工作表）
    """
    # ── 确定使用的配色 ──
    current_palette = PALETTES['cb'] if use_cb else PALETTES['default']
    chart_type = "分组柱状图" if use_group else "堆积柱状图"
    error_status = "带误差线" if use_error else "无误差线"
    
    print(f"当前使用配色: {'色盲友好' if use_cb else '默认'}")
    print(f"当前图表类型: {chart_type}")
    print(f"误差线状态: {error_status}")

    # ── 读取数据 ──
    # 获取当前脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 拼接excel文件名
    excel_path = os.path.join(script_dir, 'FEDataSourceForColumns.xlsx')
    
    if not os.path.exists(excel_path):
        print(f"错误: 找不到数据源文件 {excel_path}")
        return
    
    # 根据是否使用误差线选择数据加载方式
    if use_error:
        # 加载FE和SD数据
        df_fe, df_sd, categories, non_empty_cols = load_fe_sd_data(excel_path)
        if df_fe is None:
            return
    else:
        # 仅加载FE数据
        df_fe, categories, non_empty_cols = load_fe_data(excel_path)
        if df_fe is None:
            return
        df_sd = None  # 不使用误差线时不需要SD数据
    
    # 检查数据是否为空
    if len(non_empty_cols) == 0:
        print("错误: 没有找到有效的数据列")
        return

    # ── 画布 ──
    fig, ax = plt.subplots(figsize=(3.2, 2.56))

    # ── 根据图表类型和误差线选择绘制方式 ──
    if use_group:
        # 分组柱状图
        if use_error:
            generate_grouped_bar_with_error(ax, categories, non_empty_cols, df_fe, df_sd, current_palette)
        else:
            generate_grouped_bar_without_error(ax, categories, non_empty_cols, df_fe, current_palette)
        # 设置X轴刻度位置
        ax.set_xticks(range(len(categories)))
    else:
        # 堆积柱状图
        if use_error:
            generate_stacked_bar_with_error(ax, categories, non_empty_cols, df_fe, df_sd, current_palette)
        else:
            generate_stacked_bar_without_error(ax, categories, non_empty_cols, df_fe, current_palette)
        # 设置X轴刻度位置
        ax.set_xticks(range(len(categories)))

    # ── 轴标签 ──
    ax.set_xlabel('Potential (V vs. RHE)', fontsize=8.5, fontweight='bold')
    ax.set_ylabel('Faradaic Efficiency (%)', fontsize=8.5, fontweight='bold')

    # ── 刻度标签 ──
    ax.set_xticklabels(categories)
    
    # ── 刻度线 ──
    ax.tick_params(axis='both', labelsize=7, width=1.5)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontweight('bold')

    # ── 轴线宽度 ──
    for spine in ax.spines.values():
        spine.set_linewidth(1.5)

    # ── 调整布局 ──
    # 自定义边距 (相对坐标，范围0-1)
    left_margin = 0.18    # 左边界
    right_margin = 0.95   # 右边界
    bottom_margin = 0.18  # 下边界
    top_margin = 0.95     # 上边界
    
    plt.subplots_adjust(left=left_margin, right=right_margin, 
                        bottom=bottom_margin, top=top_margin)

    # ── 处理输出路径 ──
    if output_path is None:
        # 默认输出到数据文件所在目录
        data_dir = Path(excel_path).parent
        # 根据图表类型和是否带误差线选择默认文件名
        if use_group:
            default_filename = DEFAULT_GROUPED_FILENAME_WITH_ERROR if use_error else DEFAULT_GROUPED_FILENAME
        else:
            default_filename = DEFAULT_STACKED_FILENAME_WITH_ERROR if use_error else DEFAULT_STACKED_FILENAME
        output_path = data_dir / default_filename
    else:
        output_path = Path(output_path)
        
        # 判断是否为目录（以分隔符结尾或已存在的目录）
        if str(output_path).endswith(('/', '\\')) or output_path.is_dir():
            # 如果是目录，在目录下创建默认文件名
            if use_group:
                default_filename = DEFAULT_GROUPED_FILENAME_WITH_ERROR if use_error else DEFAULT_GROUPED_FILENAME
            else:
                default_filename = DEFAULT_STACKED_FILENAME_WITH_ERROR if use_error else DEFAULT_STACKED_FILENAME
            output_path = output_path / default_filename
        # 确保输出目录存在
        output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # 处理文件重名，避免覆盖
    output_path = get_unique_output_path(output_path)
    
    # ── 保存图片 ──
    fig.savefig(output_path, dpi=300)
    plt.close(fig)
    print(f"图片已成功保存至: {output_path}")


if __name__ == '__main__':
    # ── 设置命令行参数解析 ──
    parser = argparse.ArgumentParser(description="生成法拉第效率柱状图（支持带/不带误差线的堆积/分组柱状图）")
    parser.add_argument('-o', '--output', type=str, default=None, 
                        help='输出图片路径（支持文件路径或目录路径）')
    parser.add_argument('-cb', action='store_true', 
                        help='使用色盲友好配色 (不加此参数则使用默认配色)')
    parser.add_argument('-group', action='store_true',
                        help='生成分组并列柱状图 (不加此参数则默认生成堆积柱状图)')
    parser.add_argument('-e', action='store_true',
                        help='绘制误差线（需要SD工作表提供标准差数据），不加则只读取FE工作表')
    
    args = parser.parse_args()

    # 执行绘图
    generate_figure(output_path=args.output, use_cb=args.cb, 
                   use_group=args.group, use_error=args.e)