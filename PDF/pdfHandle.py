import os
import argparse
import sys
from copy import copy

from PyPDF2 import PdfReader, PdfWriter, PdfMerger


def pdf_encrypt(files: list, password: str = "123456"):
    """
    PDF 加密, 使用指定密码对源文件加密
    :param files: 加密文件
    :param password: 加密密码
    :return: None
    """
    total_files = len(files)    # 处理数量
    # 默认密码
    if password == "123456":
        print("# Not set password, use default password: 123456")

    for i, file in enumerate(files, 1):  # [file_name, file_path]
        pdf_reader = PdfReader(file[1])
        if pdf_reader.is_encrypted:
            continue
        pdf_writer = PdfWriter()

        # 逐页操作
        for page in range(len(pdf_reader.pages)):
            pdf_writer.add_page(pdf_reader.pages[page])

        pdf_writer.encrypt(user_password=password)  # 加密

        # 保存
        with open(file[1], 'wb') as f:
            pdf_writer.write(f)

        print_progress(i, total_files)  # 打印进度


def pdf_decrypt(files: list, password: str = '123456'):
    """
    PDF解密, 对原文件解密
    :param files: 加密的文件
    :param password: 解密密码
    :return: None
    """
    total_files = len(files)    # 处理数量

    for i, file in enumerate(files, 1):  # [file_name, file_path]
        pdf_reader = PdfReader(file[1])

        # 判断是否加密
        if pdf_reader.is_encrypted:
            pdf_reader.decrypt(password)
            # 解密错误
            if not pdf_reader.is_encrypted:
                print(f"# failed to decrypt, password:{password} is wrong!!")
                exit(1)
        else:
            continue

        pdf_writer = PdfWriter()

        # 逐页操作
        for page in range(len(pdf_reader.pages)):
            pdf_writer.add_page(pdf_reader.pages[page])

        # 保存
        with open(file[1], 'wb') as f:
            pdf_writer.write(f)

        print_progress(i, total_files)  # 打印进度


def pdf_water(files: list, water_file: str):
    """
    为文件添加水印
    :param files: 添加水印的文件
    :param water_file: 水印
    :return: None
    """
    total_files = len(files)  # 处理数量

    # 保存位置
    result_path = os.path.join(os.path.curdir, 'results')
    if not os.path.exists(result_path):
        os.makedirs(result_path)

    for i, file in enumerate(files, 1):  # [file_name, file_path]
        water_pdf = PdfReader(water_file)  # 读取水印
        mark_page = water_pdf.pages[0]  # 水印所在的页数
        pdf_reader = PdfReader(file[1])  # 读取原文件
        pdf_writer = PdfWriter()

        for page in range(len(pdf_reader.pages)):
            # 读取需要添加水印每一页pdf
            source_page = pdf_reader.pages[page]
            new_page = copy(mark_page)
            # new_page(水印)在下面，source_page原文在上面
            new_page.merge_page(source_page)
            pdf_writer.add_page(new_page)

        result_file = os.path.join(result_path, file[0])  # 文件保存位置
        # 保存
        with open(result_file, 'wb') as f:
            pdf_writer.write(f)

        print_progress(i, total_files)  # 打印进度

    print("# save-path: " + result_path)


def pdf_merge(files: list):
    """
    PDF 批量合并
    :param files: 合并的文件
    :return: None
    """
    total_files = len(files)  # 处理数量
    merger = PdfMerger()

    for i, file in enumerate(files, 1):
        merger.append(file[1])
        print_progress(i, total_files)  # 打印进度

    # 将合并后的PDF保存到当前目录
    result_file = os.path.join(os.path.curdir, 'merge_result.pdf')
    merger.write(result_file)

    merger.close()
    print("# save-path: " + result_file)


def find_all_pdf(current_dir):
    """
    寻找当前目录以及子目录 所有 PDF文件
    :param current_dir: 当前目录
    :return: PDF文件列表: [file_name, file_path]
    """
    target = []  # 记录文件路径

    # 遍历子目录所有文件
    for root, dirs, files in os.walk(current_dir):
        for file in files:
            if file.endswith(".pdf"):
                target.append([file, os.path.join(root, file)])

    return target


def find_current_dir_pdf(current_dir):
    """
    寻找当前目录的 所有 PDF文件
    :param current_dir: 当前目录
    :return: PDF文件列表: [file_name, file_path]
    """
    target = []  # 记录文件路径

    files = os.listdir(current_dir)
    # 遍历当前目录中的所有文件
    for file in files:
        if file.endswith(".pdf"):
            target.append([file, os.path.join(current_dir, file)])

    return target


def read_target_from_file(params_file):
    """
    从指定 txt 文件读取 PDF 文件
    :param params_file: 含有pdf文件路径的txt文件
    :return: PDF文件列表: [file_name, file_path]
    """
    target = []  # 记录文件路径

    with open(params_file, 'r') as f:
        lines = f.readlines()
        for line in lines:
            path = line.strip()  # 去除行尾的换行符
            if os.path.isfile(path) and path.endswith(".pdf"):  # 判断是否为pdf文件
                target.append([os.path.basename(path), path])  # 添加到列表中

    return target


def print_progress(iteration, total, prefix='Process', suffix='', length=50, fill='*'):
    """
    打印进度条信息
    :param iteration: 迭代次数
    :param total: 总量
    :param prefix: 进度条前缀
    :param suffix: 进度条后缀
    :param length: 进度条长度
    :param fill: 完成填充符
    :return: 进度条
    """
    percent = '{:.1%}'.format(iteration / total)    # 百分比进度
    filled_length = int(length * iteration // total)    # 填充长度
    bar = fill * filled_length + ' ' * (length - filled_length)     # 填充格式
    sys.stdout.write(f'\r{prefix}: [ {bar} ] {percent} {suffix} ')    # 打印输出
    sys.stdout.flush()  # 刷新缓冲


def main():
    global args
    # 文件检验
    if not os.path.exists(args.file):
        print("# File/Path does not exist")
        exit(1)

    # 获取文件路径
    if os.path.isfile(args.file):
        file_path = os.path.dirname(args.file)
    else:
        file_path = args.file
        if not (args.find_current_dir or args.find_all_dir):
            args.find_current_dir = True

    # 确认目标文件
    pdf_files = []
    if args.find_current_dir:  # 遍历当前目录
        pdf_files = find_current_dir_pdf(file_path)

    elif args.find_all_dir:  # 遍历所有子目录
        pdf_files = find_all_pdf(file_path)

    elif args.file.endswith('.txt'):  # 从文件中读取
        pdf_files = read_target_from_file(args.file)

    elif args.file.endswith('.pdf'):  # 单文件操作
        pdf_files.append([os.path.basename(args.file), args.file])

    # 输出信息
    if len(pdf_files) == 0:
        print("# PDF file is not found, please check your path")
        exit(1)

    print("-----" * 15 + "\n# The following file will be processed:")
    for f in pdf_files:
        print("\t-- " + f[1])
    print("# Processing: ")
    print("-----" * 15)

    # 执行功能
    try:
        if args.mode == 'e' or args.mode == 'encrypt':
            pdf_encrypt(pdf_files, args.password)

        elif args.mode == 'd' or args.mode == 'decrypt':
            pdf_decrypt(pdf_files, args.password)

        elif args.mode == 'm' or args.mode == 'merge':
            pdf_merge(pdf_files)

        elif args.mode == 'w' or args.mode == 'watermark':
            pdf_water(pdf_files, args.water_file)

        else:
            print("# You did not specify a function, please use \'-h\' for help!")
            exit(1)

    except Exception as e:
        print("# There something maybe useful:"
              "\n\t1. You maybe use the wrong arguments, please use \'-h' for help"
              "\n\t2. There maybe a BUG, please contact to 768476667@qq.com. Thank you!")
        print("# Details wrong are as follows:\n" + str(e))
        exit(1)

    print("\n\n### Finished ###")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog='pdfHandle',
        epilog='designed by wengui, mail: 768476667@qq.com'
    )

    # 操作原PDF文件 / 目录 / TXT读取文件
    parser.add_argument("file", help="file/directory")
    # 功能选择: 加密  解密  合并  水印
    parser.add_argument("mode", help="function to encrypt(e), decrypt(d), merge(m), watermark(w)")
    # 指定目录, 子目录所有 PDF 文件
    parser.add_argument("-r", "--all-dir", action='store_true', dest='find_all_dir',
                        help="all subdirectory files")
    # 指定当前目录所有 PDF 文件
    parser.add_argument("-d", "--current-dir", action='store_true', dest='find_current_dir',
                        help="current directory files")
    # 加密 / 解密 密码
    parser.add_argument('-p', '--password', dest='password', type=str, default="123456",
                        help='set password for encrypt or decrypt')
    # 指定 水印 文件
    parser.add_argument('-w', '--water-file', dest='water_file', type=str,
                        help='set water-file for watermark')

    args = parser.parse_args()

    print("\n# Welcome to use the PDF_HANDLE tool that designed by wengui.")
    # 执行
    main()
