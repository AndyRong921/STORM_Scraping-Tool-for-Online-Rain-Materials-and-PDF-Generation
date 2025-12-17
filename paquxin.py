import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import re
import os

# ================= 配置区域 =================
URL = "https://www.yuketang.cn/v2/web/index"
# 文件保存路径 (请确保这个路径是固定的，这样才能读取到旧文件)
SAVE_PATH = "/Users/xxxxxxxxxxx/xxxx.xlsx"  # 示例路径，请修改为你自己的
# ===========================================

def load_existing_data(filepath):
    """
    【新增功能】: 读取已有的Excel文件，恢复到内存中
    """
    if not os.path.exists(filepath):
        print("✨ 未检测到旧题库，将创建新文件。")
        return {}
    
    print(f"📂 正在加载旧题库: {filepath} ...")
    try:
        # 读取 Excel，并将空值填充为空字符串，防止 nan 报错
        df = pd.read_excel(filepath).fillna("")
        
        # 将 DataFrame 转回字典格式: { "题目内容": { "题目":..., "答案":... } }
        existing_db = {}
        for _, row in df.iterrows():
            q_text = row['题目'].strip() # 去除可能的首尾空格
            existing_db[q_text] = row.to_dict()
            
        print(f"✅ 成功加载历史题目: {len(existing_db)} 道")
        return existing_db
    except Exception as e:
        print(f"⚠️ 读取旧文件失败 (可能是格式不对)，将重新开始: {e}")
        return {}

def run_interactive_spider():
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    
    print("🚀 浏览器已启动...")
    
    # 【核心修改 1】: 启动时不再是空字典，而是先加载旧数据
    question_db = load_existing_data(SAVE_PATH)
    
    driver.get(URL)

    print("\n" + "="*60)
    print("📢 【交互模式 - 操作指南 (增量更新版)】")
    print("1. 请手动登录 -> 进课程 -> 开始答题。")
    print("2. 直接点【交卷】->【交卷】(不用做题)。")
    print("3. 点【查看试卷】，直到看见带有正确答案的详情页。")
    print("4. 回到这里按 【回车 (Enter)】，我开始智能抓取。")
    print("="*60 + "\n")
    
    batch_count = 1
    while True:
        user_input = input(f"waiting... 请操作到【答案页面】后按回车 (输入 q 退出): ")
        if user_input.lower() == 'q': break

        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])

        print(f"   ⚡️ 正在第 {batch_count} 次抓取...")

        try:
            blocks = driver.find_elements(By.CLASS_NAME, "result_item")
            
            if not blocks:
                print("   ⚠️ 没找到题目，请确认你在【查看试卷】页面！")
                continue

            new_count = 0
            for block in blocks:
                try:
                    # 1. 提取题目
                    q_text = block.find_element(By.CSS_SELECTOR, ".item-body h4").text.strip()
                    
                    # 【核心修改 2】: 查重逻辑
                    # 如果题目已经在库里（无论是刚才爬的，还是Excel里读出来的），直接跳过
                    if q_text in question_db:
                        continue

                    # --- 下面是新题处理逻辑 ---
                    
                    # 2. 智能提取选项
                    opt_eles = block.find_elements(By.CSS_SELECTOR, ".radioText, .checkboxText")
                    opts = [o.text.strip() for o in opt_eles if o.text.strip()]
                    
                    if not opts:
                        opt_eles = block.find_elements(By.CSS_SELECTOR, ".el-radio__label, .el-checkbox__label")
                        opts = [o.text.strip() for o in opt_eles if o.text.strip()]

                    # 3. 提取答案
                    full_text = block.text
                    ans_match = re.search(r"正确答案[：:]\s*([A-Za-z\s,]+)", full_text)
                    if ans_match:
                        ans = ans_match.group(1).replace(" ", "").replace(",", "").strip()
                    else:
                        ans = "未知"

                    # 4. 存入数据库
                    item_data = {
                        "题目": q_text,
                        "答案": ans
                    }
                    
                    labels = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
                    for i, label in enumerate(labels):
                        if i < len(opts):
                            item_data[label] = opts[i]
                        else:
                            item_data[label] = ""

                    question_db[q_text] = item_data
                    new_count += 1
                        
                except Exception as e:
                    continue
            
            # 显示统计信息
            print(f"   ✅ 抓取成功！本轮【新增】: {new_count} 题 | 题库总计: {len(question_db)} 题")
            
            # 只有当有新题时才写入文件，减少磁盘读写
            if new_count > 0:
                save_to_excel(question_db)
            else:
                print("   💤 本页题目都已存在，无需更新文件。")
            
            print("-" * 40)
            print("👉 下一步：手动点【返回】->【再次作答】->【交卷】->【查看试卷】")
            print("-" * 40)
            batch_count += 1

        except Exception as e:
            print(f"   ❌ 出错: {e}")

    print("程序结束。")
    driver.quit()

def save_to_excel(data):
    try:
        df = pd.DataFrame(data.values())
        # 强制按顺序排列列名
        cols = ["题目", "答案", "A", "B", "C", "D", "E", "F"]
        existing_cols = [c for c in cols if c in df.columns]
        df = df[existing_cols]
        
        df.to_excel(SAVE_PATH, index=False)
        print(f"📁 文件已保存更新: {SAVE_PATH}")
    except Exception as e:
        print(f"❌ 保存失败: {e}")

if __name__ == "__main__":
    run_interactive_spider()
