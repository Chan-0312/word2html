from win32com import client as wc
import os
import sys

word = wc.DispatchEx('Word.Application')
# 后台运行，不显示，不警告
word.Visible = 0
word.DisplayAlerts = 0

def docx2html(path):
    print(path)
    doc = word.Documents.Open(path)
    doc.SaveAs2(path.split('.doc')[0] + '.html', FileFormat=8, AddToRecentFiles=False)
    doc.Close()
    


if __name__ == "__main__":
    assert len(sys.argv) == 2, "Usage：python word2html_pywin32.py word_dir"

    # 读取参数
    dir_path = sys.argv[1]

    dir_path = "A:\\Users\\Chan\\Desktop\\add\\"
    for fpath, dir_list, file_list in os.walk(dir_path):
        for file_path in file_list:
            path = fpath+file_path
            postfix = path.split('.')[-1]
            if postfix == "docx" or postfix == "doc":
                docx2html(path)
                
    word.Quit()
    

    