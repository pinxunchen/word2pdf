import os
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

#convert_word_to_pdf(要轉檔的路徑,轉檔目的地)
convert_word_to_pdf(r'C:\Users\User\Desktop\word2pdf\word',r'C:\Users\User\Desktop\word2pdf\pdf')