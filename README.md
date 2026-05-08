# 硕士学位论文开题报告

**论文题目：基于自适应滑动窗口因子图融合的GNSS RTK/INS组合导航算法研究**

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `开题报告初稿.docx` | Word版开题报告初稿（含技术路线图） |
| `technical_roadmap.png` | 技术路线图（PNG，200 DPI） |
| `generate_roadmap.py` | 生成技术路线图的Python脚本 |
| `generate_proposal.py` | 生成Word开题报告的Python脚本 |

## 运行环境

- Python 3.9+
- matplotlib（绘图）
- python-docx（Word生成）
- 中文字体：Noto Sans CJK（系统字体）或 SimSun/SimHei

## 使用方法

```bash
# 1. 生成技术路线图
python generate_roadmap.py

# 2. 生成Word开题报告（依赖上一步生成的PNG）
python generate_proposal.py
```

## 报告结构

1. 研究意义及国内外研究现状
2. 研究目标、研究内容及拟解决的关键性问题
3. 拟采取的研究方法、技术路线、试验方案及其可行性分析
4. 课题的创新
5. 计划进度和预期成果
6. 与本课题有关的前期工作积累和已有的研究成果
7. 预期困难和解决措施
8. 参考文献（24条）
