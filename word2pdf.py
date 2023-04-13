import os
import pdfplumber
from docx2pdf import convert
import shutil


def convert_word_to_pdf(dir1, dir2):
    files = [f for f in os.listdir(dir1) if f.endswith('.docx')]    
    for file in files:
        # 組合Word檔案和PDF檔案的完整路徑
        word_path = os.path.join(dir1, file)
        pdf_path = os.path.join(dir2, file[:-5] + '.pdf') #:-5把.docx去掉換成.pdf

        # 將Word檔案轉換為PDF格式
        convert(word_path, pdf_path)
        print(f'{file} 轉換完成')


def rename_pdf(folder_path):
    # for迴圈把路徑內所有.pdf檔存進file_name
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.pdf'):
            
            # 讀取 PDF 檔案
            pdf_path = os.path.join(folder_path, file_name)
            with pdfplumber.open(pdf_path) as pdf:                
                # 讀取第二頁
                page = pdf.pages[1]
                text = page.extract_text()

                # 擷取需要的資料
                data = text.split()
                user_name = data[9]  # 個案姓名
                user_id = data[13]  # 身分證
                number = data[7][3:]  # 請款單號

                # 新檔案名稱
                new_file_name = f"{user_name}_{user_id}_{number}.pdf"
                new_file_path = os.path.join(folder_path, new_file_name)
                pdf.close()

                # 重新命名檔案
                os.rename(pdf_path, new_file_path)
                print(new_file_name + '轉換完成！')


def copy_pdf_files(src_dir, dst_dir):
    # 檢查目的地資料夾是否存在，如果不存在就建立
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
    for file_name in os.listdir(src_dir):
        if file_name.endswith('.pdf'):
            src_path = os.path.join(src_dir, file_name)
            dst_path = os.path.join(dst_dir, file_name)
            shutil.copy(src_path, dst_path)
            print(f'{file_name} 已經複製到完成版資料夾！')


def clear_word_folder(dir):
    for file_name in os.listdir(dir):
        if file_name.endswith('.docx'):
            file_path = os.path.join(dir, file_name)
            os.remove(file_path)
            print(f'{file_name} 已經刪除！')






convert_word_to_pdf('word', 'pdf')
rename_pdf('pdf')
copy_pdf_files('pdf', '完成版')
clear_word_folder('word')
