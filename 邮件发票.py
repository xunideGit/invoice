import os
import re
import pandas as pd
from email import policy
from email.parser import BytesParser
from collections import defaultdict

def is_russian_text(text):
    """通过字符频率分析检测文本是否为俄文"""
    if not text:
        return False
    
    # 俄文字母表（包括大小写）
    russian_chars = set('абвгдеёжзийклмнопрстуфхцчшщъыьэюя'
                        'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ')
    
    # 计算文本中俄文字符的比例
    russian_count = sum(1 for char in text if char in russian_chars)
    total_count = len(text.strip())
    
    if total_count == 0:
        return False
    
    # 如果俄文字符比例超过30%，则认为是俄文文本
    return (russian_count / total_count) > 0.3

def is_russian_pdf(pdf_path):
    """检查PDF文件是否包含俄文内容（使用简单的PDF文本提取）"""
    try:
        # 尝试简单的PDF文本提取
        text = ""
        with open(pdf_path, 'rb') as f:
            content = f.read()
            # 尝试提取可能的文本部分
            try:
                # 尝试解码为UTF-8
                text = content.decode('utf-8', errors='ignore')
            except UnicodeDecodeError:
                # 尝试解码为Latin-1
                text = content.decode('latin-1', errors='ignore')
        
        # 通过字符频率分析检测语言
        return is_russian_text(text)
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
        return False

def extract_amount(text):
    """从文本中提取金额信息，支持俄文数字格式"""
    # 匹配格式如: 123 456,78 或 123456,78 或 123456
    patterns = [
        r'(\d{1,3}(?:\s?\d{3})*(?:,\d{2})?)',  # 带千分位和小数点
        r'(\d+(?:,\d{2})?)',                   # 带小数点
        r'(\d+)'                               # 整数
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            amount_str = match.group(1).replace(' ', '').replace(',', '.')
            try:
                return float(amount_str)
            except ValueError:
                continue
    return None

def extract_vendor(text):
    """从文本中提取供应商信息（简化版）"""
    # 简单匹配公司名称常见词汇
    company_patterns = [
        r'ООО\s+([^\n]+)',  # 有限责任公司
        r'ЗАО\s+([^\n]+)',  # 封闭式股份公司
        r'ПАО\s+([^\n]+)',  # 开放式股份公司
        r'ИП\s+([^\n]+)',   # 个体经营者
    ]
    for pattern in company_patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1).strip()
    
    # 如果没有匹配到公司格式，尝试提取其他可能的供应商名称
    # 这里使用启发式方法，提取大写字母开头的连续单词
    lines = text.split('\n')
    for line in lines:
        if line.strip().isupper() or (line.strip() and line.strip()[0].isupper()):
            return line.strip()
    
    return "Неизвестный поставщик"  # 未知供应商

def process_eml_files(folder_path):
    """处理邮件文件夹中的所有eml文件"""
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' does not exist.")
        return
    
    # 创建临时文件夹存储提取的PDF
    temp_pdf_folder = os.path.join(os.getcwd(), "extracted_pdfs")
    os.makedirs(temp_pdf_folder, exist_ok=True)
    
    pdf_data = []  # 存储PDF文件信息
    
    # 处理所有eml文件
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.eml'):
            eml_path = os.path.join(folder_path, filename)
            try:
                # 解析eml文件
                with open(eml_path, 'rb') as f:
                    msg = BytesParser(policy=policy.default).parse(f)
                
                # 遍历邮件中的所有附件
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    
                    filename = part.get_filename()
                    if filename and filename.lower().endswith('.pdf'):
                        # 保存PDF附件
                        pdf_path = os.path.join(temp_pdf_folder, filename)
                        with open(pdf_path, 'wb') as fp:
                            fp.write(part.get_payload(decode=True))
                        
                        # 检查是否为俄文PDF
                        if is_russian_pdf(pdf_path):
                            # 提取PDF文本内容（使用简单方法）
                            text = ""
                            with open(pdf_path, 'rb') as f:
                                content = f.read()
                                try:
                                    text = content.decode('utf-8', errors='ignore')
                                except UnicodeDecodeError:
                                    text = content.decode('latin-1', errors='ignore')
                            
                            # 提取供应商和金额
                            vendor = extract_vendor(text)
                            amount = extract_amount(text)
                            
                            pdf_data.append({
                                '供应商': vendor,
                                '金额': amount,
                                'PDF文件名': filename,
                                '来源EML': eml_path
                            })
            except Exception as e:
                print(f"Error processing {eml_path}: {e}")
    
    # 创建DataFrame并生成Excel
    if pdf_data:
        df = pd.DataFrame(pdf_data)
        
        # 按供应商分组并计算总金额
        vendor_summary = df.groupby('供应商')['金额'].sum().reset_index()
        
        # 创建ExcelWriter对象
        excel_path = os.path.join(os.getcwd(), "俄文PDF分类统计.xlsx")
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # 写入详细数据
            df.to_excel(writer, sheet_name='详细数据', index=False)
            
            # 写入供应商汇总
            vendor_summary.to_excel(writer, sheet_name='供应商汇总', index=False)
            
            # 获取工作簿和工作表对象以进行格式设置
            workbook = writer.book
            worksheet_detail = writer.sheets['详细数据']
            worksheet_summary = writer.sheets['供应商汇总']
            
            # 设置金额列格式为货币格式
            money_format = workbook.add_format({'num_format': '#,##0.00'})
            for col_idx, col_name in enumerate(df.columns):
                if col_name == '金额':
                    worksheet_detail.set_column(col_idx, col_idx, 15, money_format)
            
            # 设置汇总表的金额列格式
            for col_idx, col_name in enumerate(vendor_summary.columns):
                if col_name == '金额':
                    worksheet_summary.set_column(col_idx, col_idx, 15, money_format)
        
        print(f"处理完成！Excel文件已保存至: {excel_path}")
        print(f"提取的PDF文件保存在: {temp_pdf_folder}")
    else:
        print("未找到符合条件的俄文PDF文件。")

if __name__ == "__main__":
    # 设置邮件文件夹路径
    EMAIL_FOLDER = "邮件"
    
    # 确保中文显示正常
    pd.set_option('display.unicode.ambiguous_as_wide', True)
    pd.set_option('display.unicode.east_asian_width', True)
    
    # 处理邮件文件
    process_eml_files(EMAIL_FOLDER)    
