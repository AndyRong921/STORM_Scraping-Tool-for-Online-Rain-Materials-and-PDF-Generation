import pdfplumber
import re
import os

# 定义文件名
FILE_1 = "1111111.pdf"
FILE_2 = "2222222.pdf"
OUTPUT_FILE = "差异题目汇总.txt"

def clean_text(text):
    """
    清洗文本：去除标点、空格、换行、编号，仅保留中文和关键字符。
    用于生成“指纹”进行比对。
    """
    # 去除开头的 "1. ", "102. " 等编号
    text = re.sub(r'^\d+\.\s*', '', text)
    # 去除所有非中文、非英文、非数字的字符（忽略标点差异）
    cleaned = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '', text)
    return cleaned

def extract_questions(pdf_path):
    """
    从PDF中提取题目。
    返回一个字典：{题目指纹: 题目完整原文本}
    """
    questions_map = {}
    print(f"正在读取文件: {pdf_path} ...")
    
    full_text_lines = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        # 过滤页眉页脚和干扰行
                        if "适用学期" in line or "整理人" in line or "PAGE" in line:
                            continue
                        # 过滤答案区域（通常答案区域也是 1. A 这种格式，需要根据长度或关键词过滤）
                        if "参考答案" in line or "单选题答案" in line or "多选题答案" in line:
                            # 遇到答案标题，停止后续解析（假设答案在文件末尾）
                            break
                        full_text_lines.append(line)
    except Exception as e:
        print(f"读取文件 {pdf_path} 失败: {e}")
        return {}

    # 开始解析题目
    current_q_num = None
    current_q_text = []
    
    # 匹配题目开头的正则：数字 + 点 (例如 "1. ", "105.")
    q_start_pattern = re.compile(r'^(\d+)\.\s*(.*)')
    
    for line in full_text_lines:
        line = line.strip()
        if not line:
            continue

        match = q_start_pattern.match(line)
        
        # 如果匹配到新题目的开头（例如 "1. 统战工作的..."）
        # 且后面跟的内容不是纯粹的选项（防止误判 "1. A" 这种答案列表）
        if match and len(match.group(2)) > 5: 
            # 保存上一道题
            if current_q_text:
                full_content = "\n".join(current_q_text)
                # 提取题干部分（第一行通常是题干，或者在A.选项出现之前是题干）
                # 这里简单处理：将整个题目内容清洗后作为指纹
                # 为了防止选项顺序不同导致指纹不同，这里我们只取“题干”作为指纹会更准
                # 寻找第一个选项 "A." 的位置
                option_match = re.search(r'[A-Z]\.', full_content)
                if option_match:
                    stem = full_content[:option_match.start()]
                else:
                    stem = full_content # 没找到选项，用全文

                fingerprint = clean_text(stem)
                
                # 只有指纹足够长才存入（避免存入解析错误的碎片）
                if len(fingerprint) > 5:
                    questions_map[fingerprint] = full_content

            # 开始新题目
            current_q_num = match.group(1)
            current_q_text = [line]
        else:
            # 属于当前题目的内容（选项或多行题干）
            if current_q_text:
                current_q_text.append(line)

    # 保存最后一道题
    if current_q_text:
        full_content = "\n".join(current_q_text)
        option_match = re.search(r'[A-Z]\.', full_content)
        stem = full_content[:option_match.start()] if option_match else full_content
        fingerprint = clean_text(stem)
        if len(fingerprint) > 5:
            questions_map[fingerprint] = full_content
            
    print(f"文件 {pdf_path} 解析完成，共提取 {len(questions_map)} 道题。")
    return questions_map

def main():
    if not os.path.exists(FILE_1) or not os.path.exists(FILE_2):
        print("错误：请确保两个PDF文件都在当前目录下。")
        return

    # 1. 提取题目
    q1_map = extract_questions(FILE_1)
    q2_map = extract_questions(FILE_2)

    # 2. 比对差异
    # 在 1 中但不在 2 中
    unique_to_1 = [text for fp, text in q1_map.items() if fp not in q2_map]
    # 在 2 中但不在 1 中
    unique_to_2 = [text for fp, text in q2_map.items() if fp not in q1_map]

    # 3. 输出结果
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(f"=== 对比报告 ===\n")
        f.write(f"文件1: {FILE_1}\n")
        f.write(f"文件2: {FILE_2}\n\n")
        
        f.write(f"【仅在 文件1 中出现的题目】 (共 {len(unique_to_1)} 题):\n")
        f.write("="*50 + "\n")
        for i, q in enumerate(unique_to_1, 1):
            f.write(f"[{i}] {q}\n" + "-"*30 + "\n")
            
        f.write("\n\n")
        f.write(f"【仅在 文件2 中出现的题目】 (共 {len(unique_to_2)} 题):\n")
        f.write("="*50 + "\n")
        for i, q in enumerate(unique_to_2, 1):
            f.write(f"[{i}] {q}\n" + "-"*30 + "\n")

    print(f"\n对比完成！结果已保存至: {OUTPUT_FILE}")
    print(f"仅在文件1有的题目数: {len(unique_to_1)}")
    print(f"仅在文件2有的题目数: {len(unique_to_2)}")

if __name__ == "__main__":
    main()