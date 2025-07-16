import os
import re
import pdfplumber
import pandas as pd
from email import policy
from email.parser import BytesParser
from collections import defaultdict
import langdetect
from langdetect import detect


def is_russian_pdf(pdf_path):
    """检查PDF文件是否包含俄文内容"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            # 读取前5页判断语言
            for i, page in enumerate(pdf.pages):
                if i >= 5:
                    break
                page_text = page.extract_text()
                if page_text:
                    text += page_text
                if len(text) > 1000:  # 提取足够的文本用于语言检测
                    break
            # 使用langdetect检测语言
            if len(text.strip()) < 50:  # 文本过少无法准确判断
                return False
            lang = detect(text)
            return lang == 'ru'
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
        return False


def extract_amount(text):
    """从文本中提取金额信息，支持俄文数字格式"""
    # 匹配格式如: 123 456,78 或 123456,78 或 123456
    patterns = [
        r'(\d{1,3}(?:\s?\d{3})*(?:,\d{2})?)',  # 带千分位和小数点
        r'(\d+(?:,\d{2})?)',  # 带小数点
        r'(\d+)'  # 整数
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
        r'ИП\s+([^\n]+)',  # 个体经营者
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
                            # 提取PDF文本内容
                            with pdfplumber.open(pdf_path) as pdf:
                                text = ""
                                for page in pdf.pages:
                                    page_text = page.extract_text()
                                    if page_text:
                                        text += page_text

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