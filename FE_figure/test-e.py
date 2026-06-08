#!/usr/bin/env python3
# -*- coding: utf-8 -*-

r"""
================================================================================
脚本名称: FEDataLinePlot.py
功能描述: 读取 Excel 中的法拉第效率数据，生成折线图，支持配置文件管理物质名称和LaTeX格式
         可选从SD工作表读取标准差数据绘制误差棒
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
        - 无误差棒折线图: FELineChart.png
        - 带误差棒折线图: FELineChart_with_Error.png
    分辨率: 300 DPI
    尺寸: 3.2 × 2.56 英寸
    自动重名处理: 如文件已存在，自动添加 _1, _2 等后缀

【使用示例】
    # 基础用法 (默认配色，无误差棒)
    python FEDataLinePlot.py
    
    # 带误差棒的折线图
    python FEDataLinePlot.py -e
    
    # 使用 -o 参数指定输出文件路径
    python FEDataLinePlot.py -o D:/MyFigures/my_chart.png -e
    
    # 使用色盲友好配色 + 误差棒
    python FEDataLinePlot.py -cb -e
    
    # 指定输出目录 + 误差棒
    python FEDataLinePlot.py -e -o D:/MyFigures/

【命令行参数】
    -o, --output    可选，输出图片路径（支持文件路径或目录路径）
                    - 如果是目录：在该目录下使用默认文件名
                    - 如果是文件：直接保存为该文件
                    - 不指定：默认保存到数据文件所在目录
    -cb             可选，启用色盲友好配色 (Wong 8色方案)，不加则使用 Matplotlib 默认配色
    -e              可选，绘制误差棒（需要SD工作表提供标准差数据），不加则只读取FE工作表

【图表样式说明】
    折线图样式:
        - 线条宽度: 1.2 pt
        - 标记点大小: 4 pt
        - 标记点样式: 循环使用 o, s, ^, D, v, p, h, *
        - 图例位置: 图表上方内侧居中，横向单行排列
        - 图例字号: 7
    
    误差棒样式（使用 -e 参数时）:
        - 颜色: 黑色
        - 线宽: 1 pt
        - 端点装饰线宽: 1 pt
        - 端点装饰长度: 3 pt
    
    共同样式:
        - 字体: Arial
        - X轴标签: "Potential (V vs. RHE)" (粗体)
        - Y轴标签: "Faradaic Efficiency (%)" (粗体)
        - 轴标签字号: 8.5
        - 刻度值字号: 7
        - 轴线宽度: 1.5 pt
        - Y轴范围: 底部固定为0，顶部自动留10%空白

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

【设计特点】
    1. 配色 + 形状双重编码：每种物质使用独特的颜色和标记点形状
    2. 色盲友好模式下使用 Wong 8色方案，提升可读性
    3. 图例横排在一行，适用于多物质对比场景
    4. 自动处理文件重名，避免覆盖
    5. 数据文件自动定位到脚本同目录，增强可移植性

【注意事项】
    1. 数据文件 FEDataSourceForColumns.xlsx 必须与脚本放在同一目录下
    2. 原始数据为小数 (如 0.85 代表 85%)，脚本自动转换为百分比显示
    3. 各物质数据按从右到左的顺序绘制
    4. 色盲友好模式下使用 Wong 8色方案，提升可读性
    5. 配置文件 FE_substance_config.py 与脚本放在同一目录下
    6. 图例使用 LaTeX 格式渲染，确保系统中已安装 LaTeX 支持（或使用 matplotlibrc 配置）

【版本历史】
    v3.0 - 新增 -e 参数控制是否绘制误差棒
         - 无 -e 参数时仅读取 FE 工作表
         - 有 -e 参数时同时读取 FE 和 SD 工作表并绘制误差棒
         - 误差棒样式与数据点颜色区分（统一使用黑色）
    v2.2 - 改进输入路径处理：数据文件自动定位到脚本同目录
    v2.1 - 新增 -o 参数支持文件路径和目录路径
    v2.0 - 添加配置文件支持
    v1.0 - 初始版本

================================================================================
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
import argparse
import os
import sys
import numpy as np
from pathlib import Path
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

# ── 标记点样式字典（双重保险：颜色 + 形状，提升可读性）──
MARKERS = ['o', 's', '^', 'D', 'v', 'p', 'h', '*']

# ── 默认输出文件名 ──
DEFAULT_FILENAME = 'FELineChart.png'
DEFAULT_FILENAME_WITH_ERROR = 'FELineChart_with_Error.png'

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
    加载FE和SD工作表数据（用于误差棒）
    
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

def generate_lineplot_without_error(ax, categories, non_empty_cols, df_fe, current_palette):
    """
    生成无误差棒的折线图
    """
    x_pos = range(len(categories))
    
    for i, col in enumerate(non_empty_cols):
        values = df_fe[col].values * 100
        ax.plot(
            x_pos,
            values,
            color=current_palette[i % len(current_palette)],  # 线条颜色
            marker=MARKERS[i % len(MARKERS)],                 # 标记点形状
            markersize=4,                                     # 标记点大小
            linewidth=1.2,                                   # 线条宽度
            label=get_latex_label(col)                        # 使用LaTeX格式标签
        )

def generate_lineplot_with_error(ax, categories, non_empty_cols, df_fe, df_sd, current_palette):
    """
    生成带误差棒的折线图
    
    误差棒统一使用黑色，与数据点的颜色区分开，提高可读性
    """
    x_pos = range(len(categories))
    
    for i, col in enumerate(non_empty_cols):
        values = df_fe[col].values * 100
        
        # 获取误差数据
        errors = prepare_error_data(df_sd, col, categories)
        
        # 绘制折线图（带误差棒）
        if errors is not None:
            # 先绘制折线和数据点
            line = ax.plot(
                x_pos,
                values,
                color=current_palette[i % len(current_palette)],  # 线条颜色
                marker=MARKERS[i % len(MARKERS)],                 # 标记点形状
                markersize=3.2,                                     # 标记点大小
                linewidth=1.2,                                   # 线条宽度
                label=get_latex_label(col)                        # 使用LaTeX格式标签
            )
            
            # 添加误差棒（统一使用黑色，与数据点颜色区分）
            ax.errorbar(
                x_pos,
                values,
                yerr=errors,
                fmt='none',           # 不显示数据点（已经通过plot绘制）
                ecolor='black',       # 误差棒颜色：黑色
                elinewidth=1,         # 误差棒线宽
                capsize=1.5,            # 端点装饰长度
                capthick=0.75           # 端点装饰线宽
            )
        else:
            # 如果没有误差数据，只绘制折线
            ax.plot(
                x_pos,
                values,
                color=current_palette[i % len(current_palette)],
                marker=MARKERS[i % len(MARKERS)],
                markersize=4,
                linewidth=1.2,
                label=get_latex_label(col)
            )

def generate_figure(output_path=None, use_cb=False, use_error=False):
    """
    读取数据并生成折线图，保存到指定路径
    
    参数:
        output_path: 输出路径（支持文件路径或目录路径）
        use_cb: 是否使用色盲友好配色
        use_error: 是否绘制误差棒（需要SD工作表）
    """
    # ── 确定使用的配色 ──
    current_palette = PALETTES['cb'] if use_cb else PALETTES['default']
    error_status = "带误差棒" if use_error else "无误差棒"
    
    print(f"当前使用配色: {'色盲友好' if use_cb else '默认'}")
    print(f"误差棒状态: {error_status}")

    # ── 读取数据 ──
    # 获取当前脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 拼接excel文件名
    excel_path = os.path.join(script_dir, 'FEDataSourceForColumns.xlsx')

    if not os.path.exists(excel_path):
        print(f"错误: 找不到数据源文件 {excel_path}")
        return
    
    # 根据是否使用误差棒选择数据加载方式
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
        df_sd = None  # 不使用误差棒时不需要SD数据
    
    # 检查数据是否为空
    if len(non_empty_cols) == 0:
        print("错误: 没有找到有效的数据列")
        return

    # ── 画布 ──
    fig, ax = plt.subplots(figsize=(3.2, 2.56))

    # ── 提取分类标签和位置 ──
    x_pos = range(len(categories))

    # ── 根据是否使用误差棒选择绘制方式 ──
    if use_error:
        generate_lineplot_with_error(ax, categories, non_empty_cols, df_fe, df_sd, current_palette)
    else:
        generate_lineplot_without_error(ax, categories, non_empty_cols, df_fe, current_palette)

    # ── 轴标签 ──
    ax.set_xlabel('Potential (V vs. RHE)', fontsize=8.5, fontweight='bold')
    ax.set_ylabel('Faradaic Efficiency (%)', fontsize=8.5, fontweight='bold')

    # ── 刻度线 ──
    ax.tick_params(axis='both', labelsize=7, width=1.5)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontweight('bold')

    # ── 刻度值：将数字坐标替换为分类标签 ──
    ax.set_xticks(x_pos)                     
    ax.set_xticklabels(categories)           

    # ── 轴线宽度 ──
    for spine in ax.spines.values():
        spine.set_linewidth(1.5)

    # ── Y轴范围：留出适当空白 ──
    # 计算所有数据的最大值
    y_max = df_fe[non_empty_cols].max().max() * 100
    
    # 如果使用误差棒，需要考虑误差棒的上限
    if use_error and df_sd is not None:
        for col in non_empty_cols:
            if col in df_sd.columns:
                values = df_fe[col].values * 100
                errors = df_sd[col].values * 100
                # 计算考虑误差棒后的最大值
                y_max_with_error = max(values + errors)
                y_max = max(y_max, y_max_with_error)
    
    ax.set_ylim(bottom=0, top=y_max * 1.1)  # 底部锁定0，顶部留10%空白

    # ── 图例 ──
    legend = ax.legend(
        loc='upper center',
        bbox_to_anchor=(0.5, 0.99),
        ncol=len(non_empty_cols),
        frameon=False,
        handlelength=1.5,
        handletextpad=0.4,
        columnspacing=1,
        prop={'weight': 'bold', 'size': 7},
        borderaxespad=0
    )

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
        default_filename = DEFAULT_FILENAME_WITH_ERROR if use_error else DEFAULT_FILENAME
        output_path = data_dir / default_filename
    else:
        output_path = Path(output_path)
        
        # 判断是否为目录（以分隔符结尾或已存在的目录）
        if str(output_path).endswith(('/', '\\')) or output_path.is_dir():
            # 如果是目录，在目录下创建默认文件名
            default_filename = DEFAULT_FILENAME_WITH_ERROR if use_error else DEFAULT_FILENAME
            output_path = output_path / default_filename
        # 确保输出目录存在
        output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # 处理文件重名，避免覆盖
    output_path = get_unique_output_path(output_path)
    
    # ── 保存图片 ──
    fig.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"图片已成功保存至: {output_path}")


if __name__ == '__main__':
    # ── 设置命令行参数解析 ──
    parser = argparse.ArgumentParser(description="生成法拉第效率折线图（支持带/不带误差棒）")
    parser.add_argument('-o', '--output', type=str, default=None, 
                        help='输出图片路径（支持文件路径或目录路径）')
    parser.add_argument('-cb', action='store_true', 
                        help='使用色盲友好配色 (不加此参数则使用默认配色)')
    parser.add_argument('-e', action='store_true',
                        help='绘制误差棒（需要SD工作表提供标准差数据），不加则只读取FE工作表')
    
    args = parser.parse_args()

    # 执行绘图
    generate_figure(output_path=args.output, use_cb=args.cb, use_error=args.e)