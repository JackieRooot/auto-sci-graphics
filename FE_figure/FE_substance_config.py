"""
================================================================================
物质配置文件: substance_config.py
功能: 管理物质的显示名称和LaTeX格式
================================================================================
"""

# 物质映射配置
# 格式: 'Excel列名': {'display_name': '显示名称', 'latex': 'LaTeX格式'}
SUBSTANCE_CONFIG = {
    'H2': {
        'display_name': 'H₂',
        'latex': r'$\mathregular{H_2}$'
    },
    'CO': {
        'display_name': 'CO',
        'latex': r'$\mathregular{CO}$'
    },
    'CH4': {
        'display_name': 'CH₄',
        'latex': r'$\mathregular{CH_4}$'
    },
    'C2H4': {
        'display_name': 'C₂H₄',
        'latex': r'$\mathregular{C_2H_4}$'
    },
    'CH3OH': {
        'display_name': 'CH₃OH',
        'latex': r'$\mathregular{CH_3OH}$'
    },
    'HCOOH': {
        'display_name': 'Formate',  # 修改这里即可
        'latex': r'$\mathregular{HCOOH}$'  # 或改为 r'$\mathregular{Formate}$'
    },
    'C2H5OH': {
        'display_name': 'C₂H₅OH',
        'latex': r'$\mathregular{C_2H_5OH}$'
    },
    'CH3COOH': {
        'display_name': 'CH₃COOH',
        'latex': r'$\mathregular{CH_3COOH}$'
    },
    # 新增物质示例
    'C3H6': {
        'display_name': 'C₃H₆',
        'latex': r'$\mathregular{C_3H_6}$'
    },
    'C3H7OH': {
        'display_name': 'C₃H₇OH',
        'latex': r'$\mathregular{C_3H_7OH}$'
    }
}

# 自动生成映射字典（向后兼容）
def get_latex_map():
    """返回Excel列名到LaTeX格式的映射"""
    return {name: config['latex'] for name, config in SUBSTANCE_CONFIG.items()}

def get_latex_label(col_name):
    """获取物质的 LaTeX 格式标签（用于图例）"""
    return SUBSTANCE_CONFIG.get(col_name, {}).get('latex', col_name)

def get_display_name(col_name):
    """获取物质的显示名称（用于图例）"""
    return SUBSTANCE_CONFIG.get(col_name, {}).get('display_name', col_name)
#由于latex是较好的渲染方式，这里的这个函数实际上并未调用，而是搁置不做处理。

def get_substance_list():
    """获取所有已配置的物质列表"""
    return list(SUBSTANCE_CONFIG.keys())

"""
对比三种写法
写法	物质存在且有latex	物质存在但无latex	物质不存在
.get('latex', col_name)	✅ 返回latex	✅ 返回col_name	✅ 返回col_name
.get('latex')	✅ 返回latex	❌ 返回None	❌ 返回None
['latex']	✅ 返回latex	❌ KeyError	❌ KeyError
"""