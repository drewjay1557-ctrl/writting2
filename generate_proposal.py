# -*- coding: utf-8 -*-
"""
生成GNSS RTK/INS自适应滑动窗口因子图融合开题报告Word文档
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


CN_FONT = 'SimSun'        # 中文字体（宋体，默认）
CN_FONT_HEI = 'SimHei'    # 黑体（标题）
EN_FONT = 'Times New Roman'


def set_run_font(run, size_pt=12, bold=False, cn_font=CN_FONT, en_font=EN_FONT, color=None):
    """设置字体：西文 Times New Roman + 中文宋体/黑体。"""
    run.font.name = en_font
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), en_font)
    rFonts.set(qn('w:hAnsi'), en_font)
    rFonts.set(qn('w:eastAsia'), cn_font)


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, size_pt=18, bold=True, cn_font=CN_FONT_HEI)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(18)
    return p


def add_heading(doc, text, level=1):
    """自定义标题样式"""
    sizes = {1: 16, 2: 14, 3: 12}
    size = sizes.get(level, 12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    set_run_font(run, size_pt=size, bold=True, cn_font=CN_FONT_HEI)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    return p


def add_body(doc, text, indent=True, first_line_indent_chars=2):
    """正文段落 (支持[1][2]上标参考文献标注)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    if indent:
        p.paragraph_format.first_line_indent = Pt(12 * first_line_indent_chars)

    import re
    # 用正则切分，保留 [数字] 或 [数字,数字] 作为独立段
    parts = re.split(r'(\[\d+(?:[,\-]\d+)*\])', text)
    for part in parts:
        if re.match(r'^\[\d+(?:[,\-]\d+)*\]$', part):
            # 参考文献上标
            run = p.add_run(part)
            set_run_font(run, size_pt=12)
            run.font.superscript = True
        else:
            run = p.add_run(part)
            set_run_font(run, size_pt=12)
    return p


def add_bullet(doc, text, level=0):
    """项目符号段落"""
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.left_indent = Pt(20 + level * 20)
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    import re
    parts = re.split(r'(\[\d+(?:[,\-]\d+)*\])', '• ' + text)
    for part in parts:
        if re.match(r'^\[\d+(?:[,\-]\d+)*\]$', part):
            run = p.add_run(part)
            set_run_font(run, size_pt=12)
            run.font.superscript = True
        else:
            run = p.add_run(part)
            set_run_font(run, size_pt=12)
    return p


def add_numbered(doc, text, num):
    """编号段落（第一级：1) 2) 3)）"""
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.left_indent = Pt(20)
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    import re
    parts = re.split(r'(\[\d+(?:[,\-]\d+)*\])', f"({num}) {text}")
    for part in parts:
        if re.match(r'^\[\d+(?:[,\-]\d+)*\]$', part):
            run = p.add_run(part)
            set_run_font(run, size_pt=12)
            run.font.superscript = True
        else:
            run = p.add_run(part)
            set_run_font(run, size_pt=12)
    return p


def add_reference(doc, num, text):
    """参考文献条目"""
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.left_indent = Pt(24)
    p.paragraph_format.first_line_indent = Pt(-24)  # 悬挂缩进
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(f"[{num}] {text}")
    set_run_font(run, size_pt=11)
    return p


def set_document_defaults(doc):
    """设置默认字体、页边距等"""
    # 默认 Normal 样式
    style = doc.styles['Normal']
    style.font.name = EN_FONT
    style.font.size = Pt(12)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), EN_FONT)
    rFonts.set(qn('w:hAnsi'), EN_FONT)
    rFonts.set(qn('w:eastAsia'), CN_FONT)

    # 页边距：上下 2.54cm，左右 3.18cm（A4标准）
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.18)
        section.right_margin = Cm(3.18)


def build_doc():
    doc = Document()
    set_document_defaults(doc)

    # ============ 标题 ============
    add_title(doc, "硕士学位论文开题报告")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("论文题目：基于自适应滑动窗口因子图融合的GNSS RTK/INS组合导航算法研究")
    set_run_font(run, size_pt=14, bold=True, cn_font=CN_FONT_HEI)
    p.paragraph_format.space_after = Pt(18)

    # ============ 1. 研究意义及国内外研究现状 ============
    add_heading(doc, "1  研究意义及国内外研究现状", level=1)

    add_heading(doc, "1.1  研究意义", level=2)
    add_body(doc,
        "高精度、高可靠的实时定位与导航是无人驾驶、无人机、智能机器人、精准农业等"
        "众多应用领域的核心基础。全球导航卫星系统（Global Navigation Satellite System, "
        "GNSS）因具备全天候、全球覆盖、绝对定位等特点而得到广泛应用，其中实时动态"
        "（Real-Time Kinematic, RTK）定位技术在开阔环境下可提供厘米级定位精度[1]。"
        "然而在城市峡谷、隧道、林荫道等遮挡环境下，GNSS信号易受非视距（NLOS）传播"
        "和多路径效应影响，卫星可见性下降、观测质量急剧劣化，单独使用GNSS难以满足"
        "应用对连续性和可靠性的要求[2,3]。"
    )
    add_body(doc,
        "惯性导航系统（Inertial Navigation System, INS）不依赖外部信号，能在短时间"
        "内提供高频率的位置、速度和姿态信息，与GNSS在时频域和信息域均具有良好的互补"
        "关系[4]。因此，GNSS与INS的组合导航已成为实现高精度连续定位的主流技术。"
        "传统组合方式主要基于扩展卡尔曼滤波（Extended Kalman Filter, EKF）及其变种，"
        "具有计算量小、实时性好等优点，但也存在以下不足：其一，滤波方法本质上是"
        "马尔科夫过程，仅利用当前时刻信息，对过去观测的利用不够充分；其二，线性化"
        "过程会引入模型误差，在强非线性场景下精度下降；其三，对野值和异常观测的"
        "鲁棒性较弱[5,6]。"
    )
    add_body(doc,
        "因子图优化（Factor Graph Optimization, FGO）作为一种概率图模型框架，"
        "通过在时间窗口内联合优化多时刻状态与多源观测，能够充分利用历史信息，"
        "并支持多传感器异步融合、延迟观测、非线性观测模型等复杂情况[7,8]。"
        "Kaess等提出的iSAM2增量平滑算法使因子图在实时导航中的应用成为可能[9]，"
        "近年来已在SLAM、视觉惯性导航等领域取得显著成果。将因子图方法引入"
        "GNSS RTK/INS紧组合，可有效提升在复杂环境下的定位精度和鲁棒性[10,11]。"
    )
    add_body(doc,
        "然而，现有基于因子图的GNSS/INS融合方法多采用固定长度的滑动窗口，"
        "难以兼顾精度与实时性：窗口过短会损失历史信息、降低优化效果；窗口过长则"
        "计算负担过大、难以满足实时性要求。此外，不同行驶场景（如高动态机动、"
        "静止停车、遮挡突变等）对窗口长度和因子权重的需求存在显著差异，"
        "固定参数难以适应动态变化的环境。因此，研究自适应滑动窗口因子图融合算法，"
        "根据观测质量和运动状态动态调整窗口尺寸与因子权重，对于提升复杂场景下"
        "GNSS RTK/INS组合导航系统的性能具有重要的理论意义和工程应用价值。"
    )

    add_heading(doc, "1.2  国内外研究现状", level=2)
    add_body(doc,
        "（1）GNSS/INS组合导航研究现状。GNSS与INS的组合方式按照耦合程度可分为"
        "松组合、紧组合和深组合三类[12]。松组合以GNSS定位结果作为输入，结构简单但"
        "在卫星数不足时难以工作；紧组合直接利用伪距、载波相位等原始观测，"
        "即使可见卫星少于四颗仍可提供修正信息，是当前研究的主流[13]。Wen等系统"
        "研究了城市环境下的GNSS/INS紧组合方法并提出多种误差抑制策略[3]。"
    )
    add_body(doc,
        "（2）因子图优化方法研究现状。Dellaert等开源的GTSAM库为因子图优化提供了"
        "通用平台[7]。Indelman等率先将因子图应用于GNSS/INS融合[8]。近年来"
        "Wen、Hsu等将因子图用于城市GNSS定位，结合鲁棒核函数显著提升了抗多路径能力"
        "[10,14]。2023年，Chi等发布了开源的GNSS/INS/相机集成导航库GICI-LIB，"
        "支持多种因子类型和优化后端，为相关研究提供了重要的开源平台[15]。"
    )
    add_body(doc,
        "（3）滑动窗口与自适应策略研究现状。滑动窗口技术起源于视觉惯性导航领域，"
        "Qin等在VINS-Mono中采用固定窗口滑动优化并辅以边缘化策略，实现了良好的"
        "实时性和精度[16]。Leutenegger等在OKVIS中也采用了类似方法[17]。在自适应"
        "滤波方面，Sage与Husa提出的自适应卡尔曼滤波能够在线估计噪声协方差[18]，"
        "为非平稳环境下的参数整定提供了思路。国内学者在自适应抗差滤波、方差分量"
        "估计等方面开展了大量研究[19,20]。然而，将自适应机制与滑动窗口因子图"
        "深度融合的研究尚处于起步阶段：Sünderhauf等提出Switchable Constraints"
        "机制以抑制异常因子[21]；Pfeifer等提出基于最大似然的动态协方差估计[22]。"
        "针对窗口长度的自适应研究相对较少，尚无成熟框架。"
    )
    add_body(doc,
        "综上所述，当前基于因子图的GNSS RTK/INS组合导航研究尽管取得了重要进展，"
        "但在复杂城市环境下的自适应能力、计算效率与鲁棒性方面仍有较大提升空间。"
        "针对这一问题开展深入研究具有重要意义。"
    )

    # ============ 2. 研究目标、研究内容及拟解决的关键性问题 ============
    add_heading(doc, "2  研究目标、研究内容及拟解决的关键性问题", level=1)

    add_heading(doc, "2.1  研究目标", level=2)
    add_body(doc,
        "针对复杂城市环境下GNSS RTK/INS组合导航精度波动大、鲁棒性不足的问题，"
        "本课题拟以开源GICI-LIB为基础实验平台，深入研究自适应滑动窗口因子图融合"
        "方法。总体研究目标如下：一是构建GNSS RTK与INS紧组合的因子图数学模型，"
        "完善各类因子（RTK双差因子、IMU预积分因子、先验因子等）的实现；"
        "二是提出基于观测质量和动态特性的自适应滑动窗口策略，实现窗口长度、因子"
        "权重与边缘化方案的在线动态调整；三是在仿真和实测数据集上验证所提方法"
        "相对于传统EKF和固定窗口FGO的性能提升，形成可在嵌入式平台运行的实用算法。"
    )

    add_heading(doc, "2.2  研究内容", level=2)
    add_body(doc,
        "本课题围绕研究目标，主要开展以下四方面研究工作："
    )
    add_numbered(doc,
        "GNSS RTK/INS紧组合因子图建模。建立统一的因子图数学描述，"
        "包括GNSS双差观测因子（伪距、载波相位、多普勒）、IMU预积分因子[23]、"
        "零偏随机游走因子、先验因子以及模糊度参数因子。研究各类因子的雅可比矩阵"
        "与协方差矩阵建模方法。", 1)
    add_numbered(doc,
        "自适应滑动窗口机制设计。针对窗口长度自适应问题，研究基于可见卫星数、"
        "信噪比、观测残差统计量和载体机动性的窗口长度动态调整规则；针对因子权重"
        "自适应问题，研究结合Huber、Cauchy等鲁棒核函数与方差分量估计的在线权重"
        "整定方法[21,22]；针对边缘化策略，研究信息保留与稀疏性平衡的改进方案。", 2)
    add_numbered(doc,
        "异常观测检测与剔除。基于新息卡方检验、RAIM及机器学习方法研究多路径和"
        "粗差探测[14]，结合鲁棒核函数形成多层次的异常处理机制，提升系统在GNSS信号"
        "质量劣化场景下的鲁棒性。", 3)
    add_numbered(doc,
        "算法实现与验证。基于GICI-LIB和GTSAM/Ceres开源框架实现所提算法[7,15]，"
        "在GICI公开数据集以及香港理工大学UrbanNav数据集[24]上进行对比实验，"
        "定量评估算法在不同场景下的精度、鲁棒性和实时性。", 4)

    add_heading(doc, "2.3  拟解决的关键问题", level=2)
    add_body(doc,
        "本课题拟解决以下三个关键科学与技术问题："
    )
    add_numbered(doc,
        "窗口长度的自适应调整策略。如何根据观测质量与运动状态合理选择窗口长度，"
        "在保证精度的同时控制计算量；研究动态窗口的扩展、收缩及边缘化触发条件，"
        "避免频繁调整带来的数值不稳定。", 1)
    add_numbered(doc,
        "异质观测的权重自适应与鲁棒融合。GNSS观测（NLOS/多路径/周跳）与IMU观测"
        "（温漂/振动噪声）的误差特性在不同场景下差异显著，如何在线估计各因子"
        "协方差并结合鲁棒核函数实现自适应融合，是提升系统鲁棒性的关键。", 2)
    add_numbered(doc,
        "算法的实时性与精度平衡。优化问题的维度随窗口长度快速增长，"
        "如何在嵌入式/车载平台上实现实时求解，需要兼顾边缘化策略、增量求解（iSAM2）"
        "以及工程实现层面的优化。", 3)

    # ============ 3. 拟采取的研究方法、技术路线、试验方案及其可行性分析 ============
    add_heading(doc, "3  拟采取的研究方法、技术路线、试验方案及其可行性分析", level=1)

    add_heading(doc, "3.1  研究方法", level=2)
    add_body(doc,
        "本课题综合运用理论分析、仿真研究与实测验证相结合的方法，具体包括："
    )
    add_bullet(doc,
        "文献调研法：系统梳理因子图优化、GNSS RTK、INS机械编排、自适应滤波等方向"
        "的国内外最新文献，把握研究脉络与前沿动态。")
    add_bullet(doc,
        "理论建模法：基于概率图模型和最大后验估计理论建立GNSS RTK/INS因子图数学"
        "模型，推导各类因子的观测方程、雅可比矩阵和协方差矩阵。")
    add_bullet(doc,
        "算法设计与仿真：基于MATLAB/Python进行原型仿真，验证自适应策略的有效性；"
        "进一步基于C++和GICI-LIB/GTSAM平台实现完整算法。")
    add_bullet(doc,
        "实测数据验证：采用开源数据集和自采数据进行多场景（城市、隧道、高架、"
        "静止）对比实验，与EKF、固定窗口FGO等基准算法进行性能比较。")
    add_bullet(doc,
        "对比分析法：从定位精度（RMSE）、姿态精度、鲁棒性（野值存在率、最大误差）、"
        "实时性（单帧耗时、内存占用）等多维度进行量化评估。")

    add_heading(doc, "3.2  技术路线", level=2)
    add_body(doc,
        "本课题的技术路线如图3-1所示。研究以多源观测数据预处理为起点，经过GNSS"
        "预处理（周跳探测、双差构造、电离层/对流层改正）和IMU预积分之后，构建包含"
        "RTK因子、预积分因子、先验因子与零偏因子的统一因子图。在核心优化环节引入"
        "自适应滑动窗口机制，实现窗口长度、因子权重和边缘化策略的动态调整；同时"
        "结合鲁棒核函数与新息检验进行异常探测。优化后输出高精度的位置、速度、姿态、"
        "模糊度与零偏估计结果，最后在仿真与实测数据上进行精度、鲁棒性和实时性的"
        "综合评估。"
    )

    # 插入技术路线图
    try:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture('/projects/sandbox/writting2/technical_roadmap.png', width=Cm(15))
        caption = doc.add_paragraph()
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = caption.add_run("图3-1  自适应滑动窗口因子图融合技术路线图")
        set_run_font(cap_run, size_pt=10.5, bold=True, cn_font=CN_FONT_HEI)
        caption.paragraph_format.space_after = Pt(12)
    except Exception as e:
        print(f"插入图片失败: {e}")

    add_heading(doc, "3.3  试验方案", level=2)
    add_body(doc,
        "试验方案分为三个层次进行：仿真试验、公开数据集验证与自采数据验证。"
    )
    add_numbered(doc,
        "仿真试验：基于MATLAB构建GNSS/INS仿真平台，模拟不同卫星数、不同NLOS比例、"
        "不同动态特性的运动轨迹，对比自适应窗口FGO与传统EKF、固定窗口FGO在理想"
        "条件和受污染条件下的性能表现。", 1)
    add_numbered(doc,
        "公开数据集验证：选用GICI-LIB提供的样例数据集以及香港UrbanNav数据集[24]，"
        "覆盖开阔、半遮挡、城市峡谷等多种场景，进行定位精度和鲁棒性评估。", 2)
    add_numbered(doc,
        "自采数据验证：使用课题组车载平台（含多频多系统GNSS接收机、MEMS IMU、"
        "高精度光纤惯导作为真值）在校园及城区采集多条典型路线数据，"
        "评估所提算法在实际工程中的有效性。", 3)

    add_heading(doc, "3.4  可行性分析", level=2)
    add_body(doc,
        "（1）理论可行性：因子图优化、IMU预积分、RTK双差定位、自适应滤波等基础理论"
        "均已成熟，相关文献基础扎实[7,9,18,23]。"
    )
    add_body(doc,
        "（2）技术可行性：GTSAM、Ceres、GICI-LIB等开源框架为算法实现提供了完善的"
        "支撑[7,15]；申请人已完成GICI-LIB样例数据的成功运行，对其代码结构与接口"
        "较为熟悉，具备进行二次开发的能力。"
    )
    add_body(doc,
        "（3）数据可行性：UrbanNav、GICI-LIB等公开数据集已面向学术界开放，同时"
        "课题组拥有完善的车载数据采集平台和高精度真值系统，可支撑充分的算法验证。"
    )
    add_body(doc,
        "（4）工程可行性：硕士研究阶段通常为2—3年，研究内容分阶段推进（见5.1节），"
        "时间充足；导师及课题组在组合导航领域积累深厚，可提供必要的理论指导。"
    )

    # ============ 4. 课题的创新 ============
    add_heading(doc, "4  课题的创新", level=1)
    add_body(doc, "本课题的创新点主要体现在以下三个方面：")
    add_numbered(doc,
        "提出一种基于观测质量与动态特性的自适应滑动窗口机制。不同于现有固定长度"
        "窗口方法，本课题综合利用可见卫星数、信噪比、残差统计量和载体机动性指标，"
        "构建多维自适应触发规则，实现窗口长度的在线动态调整，在保证精度的同时"
        "有效控制计算量。", 1)
    add_numbered(doc,
        "设计鲁棒核函数与方差分量估计相结合的因子权重自适应策略。将Huber、Cauchy等"
        "鲁棒核函数与在线方差估计相结合[21,22]，实现GNSS观测与IMU观测权重的"
        "动态平衡，显著提升城市环境下对NLOS和多路径的抑制能力。", 2)
    add_numbered(doc,
        "基于GICI-LIB开源平台实现完整的自适应滑动窗口RTK/INS因子图融合算法，"
        "并在多种场景下开展系统性实验验证，为后续相关研究提供可复现的开源代码"
        "与数据基础。", 3)

    # ============ 5. 计划进度和预期成果 ============
    add_heading(doc, "5  计划进度和预期成果", level=1)

    add_heading(doc, "5.1  计划进度", level=2)
    add_body(doc, "本课题计划分以下四个阶段完成：")
    add_bullet(doc, "第一阶段（2026年3月—2026年8月）：深入调研相关文献，完成GICI-LIB"
                    "源码阅读与二次开发环境搭建，复现基准算法。")
    add_bullet(doc, "第二阶段（2026年9月—2027年2月）：完成GNSS RTK/INS因子图建模与"
                    "自适应滑动窗口机制的初步设计与仿真验证。")
    add_bullet(doc, "第三阶段（2027年3月—2027年10月）：完成算法的C++工程化实现，"
                    "在公开数据集和自采数据上开展实验验证，完善算法细节。")
    add_bullet(doc, "第四阶段（2027年11月—2028年3月）：整理研究成果，撰写学位论文，"
                    "准备学术论文投稿及答辩材料。")

    add_heading(doc, "5.2  预期成果", level=2)
    add_body(doc, "预期取得的成果包括：")
    add_bullet(doc, "提出一套完整的自适应滑动窗口GNSS RTK/INS因子图融合算法，"
                    "在公开数据集和自采数据上相对EKF和固定窗口FGO在定位精度上提升5%以上。")
    add_bullet(doc, "形成一份可开源发布的算法实现代码（基于GICI-LIB平台进行扩展）。")
    add_bullet(doc, "发表学术论文2—3篇（含SCI/EI论文至少1篇）。")
    add_bullet(doc, "完成一篇高质量硕士学位论文。")

    # ============ 6. 与本课题有关的前期工作积累和已有的研究成果 ============
    add_heading(doc, "6  与本课题有关的前期工作积累和已有的研究成果", level=1)
    add_body(doc,
        "本人自进入硕士阶段以来，在组合导航方向开展了以下前期工作："
    )
    add_numbered(doc,
        "理论基础：系统学习了《卫星导航原理》、《惯性导航》、《概率图模型》、"
        "《最优估计理论》、《矩阵分析》等课程，掌握了GNSS定位原理、惯性导航机械编排、"
        "卡尔曼滤波、非线性最小二乘优化等基础知识。", 1)
    add_numbered(doc,
        "阅读了大量国内外关于因子图优化、GNSS/INS紧组合、自适应滤波及SLAM的"
        "经典文献与最新研究成果，掌握了研究领域的整体脉络和前沿动态。", 2)
    add_numbered(doc,
        "编程实践：熟练掌握C++、Python、MATLAB等编程语言，熟悉Linux开发环境，"
        "完成了EKF卡尔曼滤波、最小二乘定位等基础算法的编程实现。", 3)
    add_numbered(doc,
        "开源库学习：成功编译并运行了GICI-LIB开源库[15]，完成样例数据集的处理"
        "与可视化，对因子图构建、因子添加、优化求解等核心流程有了实际了解；"
        "同时对GTSAM、Ceres Solver等优化库的基本用法进行了学习。", 4)
    add_numbered(doc,
        "数据处理：使用RTKLIB、GAMIT等工具进行了GNSS数据后处理练习，"
        "积累了RINEX文件解析、基线解算、精度评估等方面的实践经验。", 5)
    add_numbered(doc,
        "在课题组参加组内讨论与汇报活动，与师兄师姐开展技术交流，对研究方向"
        "有了更清晰的认识。", 6)

    # ============ 7. 预期困难和解决措施 ============
    add_heading(doc, "7  预期困难和解决措施", level=1)

    add_heading(doc, "7.1  预期困难", level=2)
    add_body(doc, "在课题研究过程中，可能遇到的主要困难包括：")
    add_numbered(doc,
        "自适应窗口长度调整准则的合理设计。窗口过长、过短各有利弊，设计合理的"
        "触发阈值和调整步长并非易事，过于保守可能无效，过于激进可能导致数值"
        "不稳定。", 1)
    add_numbered(doc,
        "多类型因子的协方差在线估计。GNSS观测与IMU观测噪声模型差异很大，"
        "在有限窗口内准确估计方差分量存在数值稳定性问题。", 2)
    add_numbered(doc,
        "算法的实时性保障。因子图优化的计算量随窗口长度快速增长，"
        "在嵌入式或移动平台上实现实时求解面临挑战。", 3)
    add_numbered(doc,
        "真值数据获取及实验验证。在城市复杂环境下获取连续高精度真值（特别是"
        "隧道、地下停车场等GNSS完全失锁场景）存在客观困难。", 4)
    add_numbered(doc,
        "GICI-LIB源码深度定制。开源库代码量较大，对其核心模块的深度修改可能"
        "出现与原有接口不兼容的问题。", 5)

    add_heading(doc, "7.2  解决方案", level=2)
    add_numbered(doc,
        "针对自适应规则设计问题：采用分阶段策略，先基于仿真数据搜索最优参数区间，"
        "再引入有限离散状态的启发式规则；后续可进一步考虑引入轻量化机器学习方法"
        "对窗口长度进行智能预测。", 1)
    add_numbered(doc,
        "针对协方差在线估计问题：采用Sage-Husa自适应估计与方差分量估计相结合的"
        "方法[18]，并在异常场景下启用鲁棒核函数作为兜底方案，避免单一方法失效。", 2)
    add_numbered(doc,
        "针对实时性问题：采用iSAM2增量求解算法[9]减少重复计算；合理设置边缘化"
        "频率；对关键代码进行性能剖析和针对性优化；必要时引入并行计算。", 3)
    add_numbered(doc,
        "针对真值获取问题：采用高精度组合导航系统作为参考；对于隧道等GNSS失锁"
        "场景，通过起终点基准、前后段轨迹平滑得到近似真值，或通过仿真环境补充。", 4)
    add_numbered(doc,
        "针对源码定制问题：在深入理解GICI-LIB架构的基础上，以模块化扩展方式"
        "添加自适应策略，尽量保持原有接口的兼容性；充分利用源码注释与社区资源，"
        "必要时与原作者或开源社区交流请教。", 5)

    # ============ 参考文献 ============
    doc.add_page_break()
    add_heading(doc, "参考文献", level=1)

    refs = [
        "Teunissen P J G, Montenbruck O. Springer Handbook of Global Navigation Satellite Systems[M]. Cham: Springer, 2017.",
        "Groves P D. Shadow Matching: A New GNSS Positioning Technique for Urban Canyons[J]. Journal of Navigation, 2011, 64(3): 417-430.",
        "Wen W, Kan Y C, Hsu L T. Performance Comparison of GNSS/INS Integrations Based on EKF and Factor Graph Optimization[C]. Proceedings of the 32nd International Technical Meeting of the Satellite Division of The Institute of Navigation (ION GNSS+ 2019), 2019: 3019-3032.",
        "Titterton D, Weston J. Strapdown Inertial Navigation Technology[M]. 2nd ed. Stevenage: IET, 2004.",
        "Groves P D. Principles of GNSS, Inertial, and Multisensor Integrated Navigation Systems[M]. 2nd ed. Boston: Artech House, 2013.",
        "Li X X, Wang X B, Liao J C, et al. Semi-tightly Coupled Integration of Multi-GNSS PPP and S-VINS for Precise Positioning in GNSS-Challenged Environments[J]. Satellite Navigation, 2021, 2(1): 1-14.",
        "Dellaert F, Kaess M. Factor Graphs for Robot Perception[J]. Foundations and Trends in Robotics, 2017, 6(1-2): 1-139.",
        "Indelman V, Williams S, Kaess M, et al. Information Fusion in Navigation Systems via Factor Graph Based Incremental Smoothing[J]. Robotics and Autonomous Systems, 2013, 61(8): 721-738.",
        "Kaess M, Johannsson H, Roberts R, et al. iSAM2: Incremental Smoothing and Mapping Using the Bayes Tree[J]. International Journal of Robotics Research, 2012, 31(2): 216-235.",
        "Wen W, Pfeifer T, Bai X W, et al. Factor Graph Optimization for GNSS/INS Integration: A Comparison with the Extended Kalman Filter[J]. NAVIGATION, 2021, 68(2): 315-331.",
        "Watson R M, Gross J N. Robust Navigation in GNSS Degraded Environment Using Graph Optimization[C]. Proceedings of the 30th International Technical Meeting of the Satellite Division of The Institute of Navigation (ION GNSS+ 2017), 2017: 2906-2918.",
        "Farrell J A. Aided Navigation: GPS with High Rate Sensors[M]. New York: McGraw-Hill, 2008.",
        "Shin E H. Estimation Techniques for Low-Cost Inertial Navigation[D]. Calgary: University of Calgary, 2005.",
        "Hsu L T, Huang F J, Ng H F, et al. Hong Kong UrbanNav: An Open-source Multisensory Dataset for Benchmarking Urban Navigation Algorithms[J]. NAVIGATION, 2023, 70(4): navi.602.",
        "Chi C, Zhan X Q, Wang S N, et al. GICI-LIB: A GNSS/INS/Camera Integrated Navigation Library[J]. IEEE Robotics and Automation Letters, 2023, 8(12): 7970-7977.",
        "Qin T, Li P L, Shen S J. VINS-Mono: A Robust and Versatile Monocular Visual-Inertial State Estimator[J]. IEEE Transactions on Robotics, 2018, 34(4): 1004-1020.",
        "Leutenegger S, Lynen S, Bosse M, et al. Keyframe-based Visual-inertial Odometry Using Nonlinear Optimization[J]. International Journal of Robotics Research, 2015, 34(3): 314-334.",
        "Sage A P, Husa G W. Adaptive Filtering with Unknown Prior Statistics[C]. Proceedings of the Joint Automatic Control Conference, 1969: 760-769.",
        "杨元喜. 自适应动态导航定位[M]. 北京: 测绘出版社, 2017.",
        "Yang Y X, He H B, Xu G C. Adaptively Robust Filtering for Kinematic Geodetic Positioning[J]. Journal of Geodesy, 2001, 75(2): 109-116.",
        "Sünderhauf N, Protzel P. Switchable Constraints for Robust Pose Graph SLAM[C]. 2012 IEEE/RSJ International Conference on Intelligent Robots and Systems, 2012: 1879-1884.",
        "Pfeifer T, Protzel P. Expectation-Maximization for Adaptive Mixture Models in Graph Optimization[C]. 2019 IEEE International Conference on Robotics and Automation (ICRA), 2019: 3753-3759.",
        "Forster C, Carlone L, Dellaert F, et al. On-manifold Preintegration for Real-time Visual-Inertial Odometry[J]. IEEE Transactions on Robotics, 2017, 33(1): 1-21.",
        "Hsu L T, Kubo N, Wen W, et al. UrbanNav: An Open-sourced Multisensory Dataset for Benchmarking Positioning Algorithms Designed for Urban Areas[C]. Proceedings of the 34th International Technical Meeting of the Satellite Division of The Institute of Navigation (ION GNSS+ 2021), 2021: 226-256.",
    ]
    for i, ref in enumerate(refs, 1):
        add_reference(doc, i, ref)

    # ============ 保存 ============
    output_path = '/projects/sandbox/writting2/开题报告初稿.docx'
    doc.save(output_path)
    print(f"Word文档已生成：{output_path}")
    return output_path


if __name__ == '__main__':
    build_doc()
