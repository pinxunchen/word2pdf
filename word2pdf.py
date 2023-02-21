import os
import pdfplumber
from docx2pdf import convert


def convert_word_to_pdf(dir1, dir2):
    files = [f for f in os.listdir(dir1) if f.endswith('.docx')]
    
    for file in files:   # 使用for迴圈將所有Word檔案轉換為PDF格式，並將它們存儲在資料夾2中
        
        # 組合Word檔案和PDF檔案的完整路徑
        word_path = os.path.join(dir1, file)
        pdf_path = os.path.join(dir2, file[:-5] + '.pdf')

        # 將Word檔案轉換為PDF格式
        convert(word_path, pdf_path)

        print(f'{file} 轉換完成')


def rename_pdf(folder_path):
    # 遍歷指定路徑下的所有 PDF 檔案
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
                new_file_name = f"{user_name} {user_id} {number}.pdf"
                new_file_path = os.path.join(folder_path, new_file_name)
                pdf.close()

                # 重新命名檔案
                os.rename(pdf_path, new_file_path)
                print(f"{new_file_path} 轉換完成！")



#convert_word_to_pdf(轉檔路徑,轉檔目的地)
convert_word_to_pdf(r'C:\Users\Pin\Desktop\word2pdf\word',r'C:\Users\Pin\Desktop\word2pdf\pdf')

#rename_pdf(要重新命名的pdf路徑)
rename_pdf(r'C:\Users\Pin\Desktop\word2pdf\pdf')