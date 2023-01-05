import mammoth
import os
import sys


def docx2html(path):
    print(path)
    with open(path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value  # The generated HTML

        full_html = (
            '<!DOCTYPE html><html><head><meta charset="utf-8"/></head><body>'
            + html
            + "</body></html>"
        )
        with open(path.split(".doc")[0]+".html", "w", encoding="utf-8") as f:
            f.write(full_html)



if __name__ == "__main__":
    assert len(sys.argv) == 2, "Usage：python word2html_mammoth.py word_dir"

    # 读取参数
    dir_path = sys.argv[1]
    
    for fpath, dir_list, file_list in os.walk(dir_path):
        for file_path in file_list:
            path = fpath+"/"+file_path
            if path.split('.')[-1] == "docx":
                docx2html(path)
