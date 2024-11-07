import docx
import pandas as pd
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement

from docx.oxml.ns import qn

# 用于设置的表格
def set_cell_color(target_cell, color='5B9BD5'):
    # 创建 tcPr 元素
    tcPr = target_cell._element.tcPr if target_cell._element.tcPr is not None else OxmlElement('w:tcPr')

    # 创建 shd 元素并设置背景颜色
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color)  # 设置背景颜色为红色（FF0000）

    # 将 shd 元素添加到 tcPr 元素中
    tcPr.append(shd)

    # 将 tcPr 元素添加到单元格中
    # 将 tcPr 元素添加到单元格中
    if target_cell._element.tcPr is None:
        target_cell._element.append(tcPr)
