import hashlib

# ==========================================
# 🔧 修复: 解决 macOS/Python 环境下的 'usedforsecurity' 报错
# 这个补丁必须放在 import reportlab 之前或代码最顶部
# ==========================================
try:
    _original_md5 = hashlib.md5
    def _patched_md5(*args, **kwargs):
        # 如果调用时传入了 usedforsecurity 参数，将其移除，防止报错
        kwargs.pop('usedforsecurity', None)
        return _original_md5(*args, **kwargs)
    hashlib.md5 = _patched_md5
except Exception:
    pass
# ==========================================

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# 【修改点1：引入 Table 和 TableStyle】
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, KeepTogether, Flowable, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
# 【修改点2：引入加密模块】
from reportlab.lib.pdfencrypt import StandardEncryption
import os
import platform

# ================= 配置区域 =================
# 1. Excel 文件路径
EXCEL_PATH = "/xxxx/xxxxx/xxxxx/xxxx.xlsx"

# 2. 输出路径前缀 (会自动生成两个文件)
# 生成: /Users/.../xxxxx_解析版.pdf
# 生成: /Users/.../xxxxx_练习版.pdf
PDF_BASE_PATH = "/xxxx/xxxx/xxxx/你的文件名字"

# 3. 页眉内容
HEADER_TEXT = "xxxxxxxxxxxxxxxxxxx"

# ===========================================

class HorizontalLine(Flowable):
    """自定义分割线组件"""
    def __init__(self, width=440):
        Flowable.__init__(self)
        self.width = width

    def draw(self):
        self.canv.setStrokeColor(colors.lightgrey)
        self.canv.setLineWidth(0.5)
        self.canv.line(0, 0, self.width, 0)

def get_system_font_path():
    """
    自动寻找可用的中文字体路径，覆盖更多 macOS 路径
    """
    candidate_fonts = [
        "SimHei.ttf",                                
        "/Users/rongzhijin/Downloads/SimHei.ttf",    
        # macOS 常见字体
        "/System/Library/Fonts/PingFang.ttc",                # 苹方 (最稳)
        "/System/Library/Fonts/Supplemental/Songti.ttc",     # macOS 新版宋体
        "/System/Library/Fonts/Supplemental/STHeiti Light.ttc", # macOS 新版黑体
        "/System/Library/Fonts/STHeiti Medium.ttc",  
        "/System/Library/Fonts/STHeiti Light.ttc",   
        "/Library/Fonts/Songti.ttc",
        "/Library/Fonts/Arial Unicode.ttf",
        # Windows 常见字体
        "C:\\Windows\\Fonts\\simhei.ttf",            
        "C:\\Windows\\Fonts\\msyh.ttf"               
    ]

    for font_path in candidate_fonts:
        if os.path.exists(font_path):
            print(f"✅ 已自动找到可用字体: {font_path}")
            return font_path
    
    return None

# 【修改点3：定义绘制页眉的函数】
def draw_header(canvas, doc):
    """
    在每一页绘制页眉
    """
    canvas.saveState()
    
    # 尝试使用注册的中文字体，如果失败则回退
    try:
        canvas.setFont('ChineseFont', 9)
    except:
        canvas.setFont('Helvetica', 9)
    
    page_width, page_height = A4
    
    # 页眉文字位置：右上角
    # A4 宽度约为 595.27 points
    # 右对齐：x 坐标设为 (页面宽度 - 右边距)
    x_pos = page_width - 2*cm
    y_pos = page_height - 1.0*cm  # 距离顶部 1.0cm
    
    # 绘制文字 (右对齐)
    canvas.drawRightString(x_pos, y_pos, HEADER_TEXT)
    
    # 绘制页眉分割线 (灰色细线)
    canvas.setLineWidth(0.5)
    canvas.setStrokeColor(colors.grey)
    # 线条从左边距到右边距
    canvas.line(2*cm, y_pos - 0.2*cm, page_width - 2*cm, y_pos - 0.2*cm)
    
    canvas.restoreState()

def create_pdf_file(filename, single_choice, multi_choice, font_name, mode='inline'):
    """
    核心生成函数
    :param mode: 'inline' (答案在题目下) 或 'end' (答案在文档末尾)
    """
    # 计算有效内容宽度 (A4宽 - 左右边距)
    content_width = A4[0] - 4*cm
    
    # 【修改点4：设置加密权限】
    # userPassword="" 表示打开不需要密码
    # ownerPassword 设置一个复杂密码，用于限制权限
    # canModify=0 禁止修改, canPrint=1 允许打印, canCopy=1 允许复制
    encrypt_config = StandardEncryption(
        userPassword="", 
        ownerPassword="在这里输入密码", 
        canPrint=1, 
        canModify=0, 
        canCopy=1, 
        canAnnotate=0
    )

    doc = SimpleDocTemplate(
        filename, 
        pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
        encrypt=encrypt_config  # 应用加密
    )
    
    # --- 定义精美样式 ---
    styles = getSampleStyleSheet()
    
    # 标题
    title_style = ParagraphStyle(
        name='ExamTitle', parent=styles['Heading1'], fontName=font_name,
        fontSize=20, alignment=1, spaceAfter=20, textColor=colors.black
    )
    
    # 大类标题 (一、单选题)
    section_style = ParagraphStyle(
        name='SectionHeader', parent=styles['Heading2'], fontName=font_name,
        fontSize=15, spaceBefore=15, spaceAfter=10, 
        textColor=colors.HexColor("#2c3e50"), # 深蓝色
        borderPadding=5
    )
    
    # 题目文本
    question_style = ParagraphStyle(
        name='QuestionText', parent=styles['Normal'], fontName=font_name,
        fontSize=11, leading=18, spaceAfter=8, textColor=colors.black
    )
    
    # 选项文本
    option_style = ParagraphStyle(
        name='OptionText', parent=styles['Normal'], fontName=font_name,
        fontSize=10.5, leftIndent=15, leading=16, textColor=colors.HexColor("#34495e")
    )
    
    # 答案文本 (解析版用) - 【改为绿色系】
    answer_style = ParagraphStyle(
        name='AnswerText', parent=styles['Normal'], fontName=font_name,
        fontSize=10, textColor=colors.HexColor("#1e8449"), # 深绿色文字
        leftIndent=15, spaceBefore=5, spaceAfter=5,
        backColor=colors.HexColor("#e8f8f5"), # 淡绿色背景
        borderPadding=3
    )

    story = []
    story.append(Paragraph("文档里的大标题在此输入", title_style))
    story.append(Spacer(1, 0.5*cm))

    # 用于收集末尾答案
    single_choice_answers = []
    multi_choice_answers = []

    # === 处理单选题 ===
    if single_choice:
        story.append(Paragraph(f"一、单选题 (共 {len(single_choice)} 题)", section_style))
        story.append(HorizontalLine())
        story.append(Spacer(1, 0.3*cm))
        
        for i, q in enumerate(single_choice):
            idx = i + 1
            # 构建单题块 (使用 KeepTogether 防止跨页断裂)
            q_elements = []
            
            # 题目
            q_text = f"<b>{idx}.</b> {q['title']}"
            q_elements.append(Paragraph(q_text, question_style))
            
            # 选项
            for opt in q['options']:
                q_elements.append(Paragraph(opt, option_style))
            
            # 答案处理
            if mode == 'inline':
                q_elements.append(Paragraph(f"<b>【正确答案】 {q['answer']}</b>", answer_style))
            else:
                # 练习版：只收集答案
                single_choice_answers.append(f"{idx}.{q['answer']}")
            
            q_elements.append(Spacer(1, 0.4*cm))
            q_elements.append(HorizontalLine()) # 分割线
            q_elements.append(Spacer(1, 0.4*cm))
            
            story.append(KeepTogether(q_elements))

    # === 处理多选题 ===
    if multi_choice:
        story.append(PageBreak()) # 多选题另起一页
        story.append(Paragraph(f"二、多选题 (共 {len(multi_choice)} 题)", section_style))
        story.append(HorizontalLine())
        story.append(Spacer(1, 0.3*cm))
        
        for i, q in enumerate(multi_choice):
            idx = i + 1
            q_elements = []
            
            # 题目
            q_text = f"<b>{idx}.</b> {q['title']}"
            q_elements.append(Paragraph(q_text, question_style))
            
            # 选项
            for opt in q['options']:
                q_elements.append(Paragraph(opt, option_style))
            
            # 答案处理
            if mode == 'inline':
                q_elements.append(Paragraph(f"<b>【正确答案】 {q['answer']}</b>", answer_style))
            else:
                # 收集多选题答案
                multi_choice_answers.append(f"{idx}.{q['answer']}")
            
            q_elements.append(Spacer(1, 0.4*cm))
            q_elements.append(HorizontalLine())
            q_elements.append(Spacer(1, 0.4*cm))
            
            story.append(KeepTogether(q_elements))

    # === 如果是练习版，在末尾添加答案汇总 ===
    if mode == 'end':
        story.append(PageBreak())
        story.append(Paragraph("参考答案", title_style))
        story.append(HorizontalLine())
        story.append(Spacer(1, 0.5*cm))
        
        # 定义通用表格样式
        matrix_style = TableStyle([
            ('FONTNAME', (0,0), (-1,-1), font_name), # 字体
            ('FONTSIZE', (0,0), (-1,-1), 11),        # 字号
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),       # 左对齐
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),    # 垂直居中
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),    # 行底间距
            ('TOPPADDING', (0,0), (-1,-1), 6),       # 行顶间距
        ])

        # 1. 输出单选题答案 (方阵排版)
        if single_choice_answers:
            story.append(Paragraph(f"<b>一、单选题答案</b>", section_style))
            story.append(Spacer(1, 0.2*cm))
            
            # 【修改点2：使用 Table 实现方阵】
            cols_count = 8 # 每行8个
            table_data = []
            row = []
            for item in single_choice_answers:
                row.append(item)
                if len(row) == cols_count:
                    table_data.append(row)
                    row = []
            # 补齐最后一行
            if row:
                while len(row) < cols_count:
                    row.append("")
                table_data.append(row)
            
            # 自动计算列宽
            col_width = content_width / cols_count
            t = Table(table_data, colWidths=[col_width] * cols_count)
            t.setStyle(matrix_style)
            story.append(t)
            story.append(Spacer(1, 0.5*cm))

        # 2. 输出多选题答案 (方阵排版)
        if multi_choice_answers:
            story.append(Paragraph(f"<b>二、多选题答案</b>", section_style))
            story.append(Spacer(1, 0.2*cm))
            
            # 多选题比较长，每行5个
            cols_count = 5 
            table_data = []
            row = []
            for item in multi_choice_answers:
                row.append(item)
                if len(row) == cols_count:
                    table_data.append(row)
                    row = []
            if row:
                while len(row) < cols_count:
                    row.append("")
                table_data.append(row)
            
            col_width = content_width / cols_count
            t = Table(table_data, colWidths=[col_width] * cols_count)
            t.setStyle(matrix_style)
            story.append(t)

    # 生成文件
    try:
        print(f"📄 正在写入 PDF 文件: {filename} ...")
        # 【修改点5：绑定页眉绘制函数】
        doc.build(story, onFirstPage=draw_header, onLaterPages=draw_header)
        print(f"✅ 成功! 文件已生成: {filename}")
    except Exception as e:
        print(f"❌ 生成文件失败: {e}")

def generate_exam_pdf():
    print("🚀 开始 PDF 生成程序...")
    
    # 0. 检查 Excel 文件是否存在
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ 错误: 找不到 Excel 文件!")
        print(f"   路径: {EXCEL_PATH}")
        print("   请检查文件名是否正确，或者是否已经运行爬虫脚本生成了文件。")
        return

    # 1. 准备字体
    print("🔍 正在查找中文字体...")
    font_path = get_system_font_path()
    if not font_path:
        print("❌ 未找到中文字体，无法生成 PDF。")
        print("   请尝试手动下载 SimHei.ttf 并放到代码目录下。")
        return

    # 注册字体
    try:
        pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
    except Exception as e:
        try:
             # 尝试读取 TTC 集合中的第一个字体
             pdfmetrics.registerFont(TTFont('ChineseFont', font_path, subfontIndex=0))
        except:
             print(f"❌ 字体注册失败 ({font_path})，请检查文件权限或尝试手动下载 SimHei.ttf")
             return

    # 2. 读取 Excel
    print(f"📊 读取 Excel: {EXCEL_PATH} ...")
    try:
        df = pd.read_excel(EXCEL_PATH)
        df = df.fillna("")
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return

    # 3. 数据分类
    single_choice_list = []
    multi_choice_list = []

    for index, row in df.iterrows():
        try:
            ans = str(row['答案']).strip()
            clean_ans = ans.replace(" ", "").replace(",", "")
            
            question_data = {
                "title": str(row['题目']),
                "options": [],
                "answer": ans
            }
            
            for label in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                if label in row and str(row[label]).strip() != "":
                    question_data['options'].append(f"{label}. {row[label]}")
            
            if len(clean_ans) > 1:
                multi_choice_list.append(question_data)
            else:
                single_choice_list.append(question_data)
        except Exception as row_e:
            print(f"⚠️ 跳过一行数据 (格式错误): {row_e}")
            continue

    print(f"📚 题目统计：单选 {len(single_choice_list)} | 多选 {len(multi_choice_list)}")

    if not single_choice_list and not multi_choice_list:
        print("❌ 错误: Excel 中没有解析出任何题目，请检查 Excel 内容格式。")
        return

    # 4. 生成两个版本
    # 版本A: 解析版 (答案在题目下)
    file_path_inline = f"{PDF_BASE_PATH}_解析版.pdf"
    create_pdf_file(file_path_inline, single_choice_list, multi_choice_list, 'ChineseFont', mode='inline')

    # 版本B: 练习版 (答案在最后)
    file_path_end = f"{PDF_BASE_PATH}_练习版.pdf"
    create_pdf_file(file_path_end, single_choice_list, multi_choice_list, 'ChineseFont', mode='end')
    
    # 自动打开文件夹
    try:
        os.system(f"open {os.path.dirname(PDF_BASE_PATH)}")
    except:
        pass

    print("🎉 所有任务已完成!")

if __name__ == "__main__":
    try:
        generate_exam_pdf()
    except Exception as e:
        print(f"❌ 程序发生意外错误: {e}")