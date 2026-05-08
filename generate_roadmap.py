# -*- coding: utf-8 -*-
"""
绘制GNSS RTK/INS自适应滑动窗口因子图融合技术路线图
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib

# 配置中文字体
matplotlib.rcParams['font.sans-serif'] = ['Noto Sans CJK SC', 'Noto Sans CJK JP', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False


def draw_box(ax, x, y, w, h, text, color='#E3F2FD', ec='#1976D2', fontsize=10, fontweight='normal'):
    box = FancyBboxPatch((x, y), w, h,
                         boxstyle="round,pad=0.02,rounding_size=0.08",
                         linewidth=1.5, edgecolor=ec, facecolor=color)
    ax.add_patch(box)
    ax.text(x + w/2, y + h/2, text, ha='center', va='center',
            fontsize=fontsize, fontweight=fontweight, wrap=True)


def draw_arrow(ax, x1, y1, x2, y2, color='#333333', style='-|>'):
    arrow = FancyArrowPatch((x1, y1), (x2, y2),
                            arrowstyle=style, mutation_scale=18,
                            linewidth=1.4, color=color)
    ax.add_patch(arrow)


fig, ax = plt.subplots(figsize=(14, 11))
ax.set_xlim(0, 14)
ax.set_ylim(0, 14)
ax.axis('off')

# 标题
ax.text(7, 13.5, 'GNSS RTK/INS 自适应滑动窗口因子图融合技术路线图',
        ha='center', va='center', fontsize=16, fontweight='bold')

# ===== 第一层：原始传感器数据 =====
draw_box(ax, 0.5, 11.6, 3.0, 1.0, '多频多系统\nGNSS观测数据\n(伪距/载波/多普勒)', color='#FFF3E0', ec='#E65100', fontsize=10)
draw_box(ax, 4.0, 11.6, 2.7, 1.0, 'MEMS/战术级\nIMU数据\n(角速度/比力)', color='#FFF3E0', ec='#E65100', fontsize=10)
draw_box(ax, 7.2, 11.6, 2.7, 1.0, '基准站观测\n/星历/差分改正', color='#FFF3E0', ec='#E65100', fontsize=10)
draw_box(ax, 10.4, 11.6, 3.0, 1.0, '环境辅助信息\n(可视星数/多路径指标)', color='#FFF3E0', ec='#E65100', fontsize=10)

# ===== 第二层：预处理 =====
draw_box(ax, 0.5, 9.8, 4.5, 1.0, 'GNSS数据预处理\n周跳探测 / 整周模糊度初始化 / 双差构建',
         color='#E8F5E9', ec='#2E7D32', fontsize=10)
draw_box(ax, 5.3, 9.8, 4.0, 1.0, 'IMU预积分\n零偏建模 / 协方差传递',
         color='#E8F5E9', ec='#2E7D32', fontsize=10)
draw_box(ax, 9.6, 9.8, 3.8, 1.0, '观测质量评估\nSNR/高度角/残差统计',
         color='#E8F5E9', ec='#2E7D32', fontsize=10)

for x_start in [2.0, 5.35, 8.55, 11.9]:
    draw_arrow(ax, x_start, 11.6, x_start, 10.8)

# ===== 第三层：因子图构建 =====
draw_box(ax, 1.0, 8.0, 12.0, 1.2,
         '统一因子图建模：GNSS RTK 因子 + IMU 预积分因子 + 先验因子 + 零偏随机游走因子',
         color='#E1F5FE', ec='#0277BD', fontsize=12, fontweight='bold')

draw_arrow(ax, 2.75, 9.8, 3.5, 9.2)
draw_arrow(ax, 7.3, 9.8, 7.0, 9.2)
draw_arrow(ax, 11.5, 9.8, 10.5, 9.2)

# ===== 第四层：核心创新 —— 自适应滑动窗口 =====
draw_box(ax, 0.5, 5.6, 13.0, 2.0,
         '【核心创新】自适应滑动窗口机制',
         color='#FFF8E1', ec='#F57F17', fontsize=13, fontweight='bold')

draw_box(ax, 0.9, 6.2, 3.9, 0.9,
         '窗口尺寸自适应调整\n(基于观测质量/动态性)',
         color='#FFFDE7', ec='#F9A825', fontsize=9)
draw_box(ax, 5.05, 6.2, 3.9, 0.9,
         '因子权重在线估计\n(鲁棒核函数 + 方差自整定)',
         color='#FFFDE7', ec='#F9A825', fontsize=9)
draw_box(ax, 9.2, 6.2, 4.2, 0.9,
         '边缘化策略改进\n(信息保留 vs 稀疏性平衡)',
         color='#FFFDE7', ec='#F9A825', fontsize=9)

draw_arrow(ax, 7.0, 8.0, 7.0, 7.6)

# ===== 第五层：优化求解 =====
draw_box(ax, 1.0, 3.8, 5.8, 1.2,
         '非线性最小二乘优化\n(Ceres/GTSAM)\niSAM2增量求解',
         color='#F3E5F5', ec='#6A1B9A', fontsize=11)
draw_box(ax, 7.2, 3.8, 5.8, 1.2,
         '故障检测与剔除 (FDE)\n假设检验 / 新息χ²检验\n粗差隔离',
         color='#F3E5F5', ec='#6A1B9A', fontsize=11)

draw_arrow(ax, 4.5, 5.6, 4.0, 5.0)
draw_arrow(ax, 9.5, 5.6, 10.0, 5.0)

# ===== 第六层：状态输出 =====
draw_box(ax, 2.5, 2.0, 9.0, 1.2,
         '状态估计输出\n位置 / 速度 / 姿态 / 模糊度 / 零偏',
         color='#E0F2F1', ec='#00695C', fontsize=12, fontweight='bold')

draw_arrow(ax, 4.0, 3.8, 5.5, 3.2)
draw_arrow(ax, 10.0, 3.8, 8.5, 3.2)

# ===== 第七层：验证评估 =====
draw_box(ax, 0.5, 0.2, 4.0, 1.2,
         '仿真数据验证\n(不同动态/遮挡场景)',
         color='#FFEBEE', ec='#C62828', fontsize=10)
draw_box(ax, 5.0, 0.2, 4.0, 1.2,
         '实测数据验证\nGICI-LIB/urbanNav',
         color='#FFEBEE', ec='#C62828', fontsize=10)
draw_box(ax, 9.5, 0.2, 4.0, 1.2,
         '精度/鲁棒性/实时性\n对比传统EKF、固定窗口FGO',
         color='#FFEBEE', ec='#C62828', fontsize=10)

draw_arrow(ax, 5.5, 2.0, 2.5, 1.4)
draw_arrow(ax, 7.0, 2.0, 7.0, 1.4)
draw_arrow(ax, 8.5, 2.0, 11.5, 1.4)

plt.tight_layout()
plt.savefig('/projects/sandbox/writting2/technical_roadmap.png',
            dpi=200, bbox_inches='tight', facecolor='white')
plt.close()
print('技术路线图已生成: technical_roadmap.png')
