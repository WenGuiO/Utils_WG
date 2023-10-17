import os
import sys
import time

import win32com.client

out_path = os.path.dirname(os.path.abspath(__file__))


def to_pdf_file(filename, out_path, word):
    try:
        doc = word.Documents.Open(out_path + filename)
        # FileFormat=17 表示PDF格式
        doc.SaveAs(out_path + filename.split(".")[0] + ".pdf", FileFormat=17)
        doc.Close()
        return "> [OK] " + filename
    except Exception as e:
        return "> [X] " + filename


def get_current_dir_file(current_dir):
    files = []
    for file in os.listdir(current_dir):
        if file.endswith(".docx") or file.endswith(".doc"):
            files.append(file)

    return files


def start(filepath, out_path=''):
    if os.path.exists(filepath):
        # 创建Word应用程序对象
        word = win32com.client.Dispatch("Word.Application")
        start_time = time.time()

        out_path = out_path if out_path.endswith("\\") else out_path + "\\"

        # 单文件处理
        if os.path.isfile(filepath) and (filepath.endswith('.doc') or filepath.endswith('.docx')):
            to_pdf_file(filepath,out_path, word)

        # 多文件逐个处理
        elif os.path.isdir(filepath):
            files = get_current_dir_file(filepath)
            print(f"# [Info] 共{len(files)}个文档")
            for f in files:
                print(to_pdf_file(f, out_path, word))

        else:
            print("# [Error] 指定文件格式错误")

        word.Quit()
        print(f"\n# [Finish] 完成时间: {time.time() - start_time:.2f} s\n")

    else:
        print("# [Error] 路径/文件存在错误\n")


if __name__ == '__main__':
    if len(sys.argv) < 2 or sys.argv[1] == '-h':
        print("# [Help] python wordToPDF.py 文件/目录路径\n")
    else:
        start(sys.argv[1], os.path.dirname(sys.argv[1]))
