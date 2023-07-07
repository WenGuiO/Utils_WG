import os
import argparse
import sys
from copy import copy

from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from argparse import Namespace


def pdf_encrypt(files: list, arg: Namespace):
    """
    PDF 加密, 使用指定密码对源文件加密
    :param files: 加密文件
    :param arg: 用户参数
    :return: None
    """
    files_nums = len(files)    # 文件数量
    failed_files = []   # 加密失败文件
    success_files = []    # 加密成功

    if arg.password == "123456":    # 未设置密码, 使用默认密码
        print("# Not set password, use default password: 123456")

    for i, file in enumerate(files, 1):  # [file_name, file_path]
        pdf_reader = PdfReader(file[1])

        if pdf_reader.is_encrypted:     # 已加密
            failed_files.append([file[1], '已加密'])   # 失败记录
            files_nums -= 1
            continue

        pdf_writer = PdfWriter()

        # 逐页操作
        page_length = len(pdf_reader.pages)
        for page in range(page_length):
            pdf_writer.add_page(pdf_reader.pages[page])
            # 打印进度条:  当前文件序号/文件总数 + (单文件当前执行页数/总页数 * 进度条长度/文件总数)
            print_progress(i, files_nums, extra_add=(page * 50) / (page_length * files_nums))

        try:
            pdf_writer.encrypt(user_password=arg.password)  # 加密

        except Exception as e:      # 加密异常
            failed_files.append([file[1], '加密失败: ' + str(e)])   # 失败记录
            files_nums -= 1
            continue

        # 保存
        with open(file[1], 'wb') as f:
            pdf_writer.write(f)
        success_files.append(file[1])   # 成功记录

        return success_files, failed_files


def pdf_decrypt(files: list, arg: Namespace):
    """
    PDF解密, 对原文件解密
    :param files: 加密的文件
    :param arg: 用户参数
    :return: None
    """
    files_num = len(files)    # 文件数量
    failed_files = []  # 加密失败文件
    success_files = []  # 加密成功

    for i, file in enumerate(files, 1):  # [file_name, file_path]
        pdf_reader = PdfReader(file[1])

        # 判断是否加密
        if pdf_reader.is_encrypted:
            pdf_reader.decrypt(arg.password)
            # 解密错误
            if not pdf_reader.is_encrypted:     # TODO: 添加密码本
                print(f"# failed to decrypt, password:{arg.password} is wrong!!")
                exit(1)
        else:
            continue

        pdf_writer = PdfWriter()

        page_length = len(pdf_reader.pages)
        # 逐页操作
        for page in range(page_length):
            pdf_writer.add_page(pdf_reader.pages[page])
            # 打印进度条:  当前文件序号/文件总数 + (单文件当前执行页数/总页数 * 进度条长度/文件总数)
            print_progress(i, files_num, extra_add=(page * 50) / (page_length * files_num))

        # 保存
        with open(file[1], 'wb') as f:
            pdf_writer.write(f)

        print_progress(i, files_num)  # 打印进度


def pdf_watermark(files: list, arg: Namespace):
    """
    为文件添加水印
    :param files: 添加水印的文件
    :param arg: 用户参数
    :return: None
    """
    total_files = len(files)  # 处理数量

    # 保存位置
    result_path = os.path.join(os.path.curdir, 'results')
    if not os.path.exists(result_path):
        os.makedirs(result_path)

    for i, file in enumerate(files, 1):  # [file_name, file_path]
        water_pdf = PdfReader(arg.water_file)  # 读取水印
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


def pdf_merge(files: list, arg: Namespace):
    """
    PDF 批量合并
    :param files: 合并的文件
    :param arg: 用户参数
    :return: None
    """
    files_num = len(files)  # 处理数量
    merger = PdfMerger()

    for i, file in enumerate(files, 1):
        merger.append(file[1])
        print_progress(i, files_num)  # 打印进度

    # 将合并后的PDF保存到当前目录
    result_file = os.path.join(os.path.curdir, 'merge_result.pdf')
    merger.write(result_file)

    merger.close()
    print("# save-path: " + result_file)


def find_all_pdf(current_dir, ex_name: str = '.pdf'):
    """
    寻找当前目录以及子目录 所有 PDF文件
    :param current_dir: 当前目录
    :param ex_name:
    :return: PDF文件列表: [file_name, file_path]
    """
    target = []  # 记录文件路径

    # 遍历子目录所有文件
    for root, dirs, files in os.walk(current_dir):
        for file in files:
            if file.endswith(ex_name):
                target.append([file, os.path.join(root, file)])

    return target


def find_current_dir_pdf(current_dir, ex_name: str = '.pdf'):
    """
    寻找当前目录的 所有 PDF文件
    :param current_dir: 当前目录
    :param ex_name:
    :return: PDF文件列表: [file_name, file_path]
    """
    target = []  # 记录文件路径

    files = os.listdir(current_dir)     # 获取所有文件

    # 遍历当前目录中的所有文件
    for file in files:
        if file.endswith(ex_name):
            target.append([file, os.path.join(current_dir, file)])

    return target


def read_target_from_file(book: str, file_ex_name: str = '.pdf'):
    """
    从指定 txt 文件读取 PDF 文件
    :param book: 含有pdf文件路径的txt文件
    :param file_ex_name: ss
    :return: PDF文件列表: [file_name, file_path]
    """
    target = []  # 记录文件路径

    with open(book, 'r') as f:
        lines = f.readlines()
        for line in lines:
            path = line.strip()  # 去除行尾的换行符
            if os.path.isfile(path) and path.endswith(file_ex_name):  # 判断是否为pdf文件
                target.append([os.path.basename(path), path])  # 添加到列表中

    return target


def get_files(path: str, ex_name: str = '.pdf', mode: int = 0):
    """
    获取<指定目录>的<指定类型>文件。使用 path、ex_name、mode 是为方便后续可能出现的拓展
    :param path: 指定扫描目录
    :param ex_name: 指定类型文件
    :param mode:  模式
    :return: 文件列表 [ [filename, filepath], ...]
    """
    global args
    # 文件检验
    if not os.path.exists(path):
        print("# File/Path does not exist")
        exit(1)

    # 获取文件所在目录
    if os.path.isfile(path):
        file_path = os.path.dirname(path)
    else:
        file_path = args.file

    # 确认目标文件
    files = []
    if path.endswith('.txt'):  # 从文件中读取, 不受'-d'和'-r'限制, 故放最上面
        files = read_target_from_file(path, ex_name)

    elif mode == 1:     # 指定模式: 寻找所有文件, 默认不执行
        files = find_all_pdf(file_path)

    elif args.find_current_dir:  # 遍历当前目录, 优先级比'-r'高
        files = find_current_dir_pdf(file_path)

    elif args.find_all_dir:  # 遍历所有子目录
        files = find_all_pdf(file_path)

    elif path.endswith('.pdf'):  # 单文件操作
        files.append([os.path.basename(args.file), args.file])

    else:  # 对其他后缀文件, 默认遍历当前目录
        files = find_current_dir_pdf(file_path)

    return files


def print_progress(iteration: float, total: int, prefix: str = 'Process', suffix: str = '',
                   length: int = 50, fill: str = '*', extra_add: float = 0):
    """
    打印进度条信息
    :param iteration: 迭代次数
    :param total: 总量
    :param prefix: 进度条前缀
    :param suffix: 进度条后缀
    :param length: 进度条长度
    :param fill: 完成填充符
    :param extra_add:
    :return: 进度条
    """
    percent = '{:.2%}'.format(iteration / total + extra_add)    # 百分比进度
    filled_length = int(length * (iteration / total + extra_add))    # 填充长度
    bar = fill * filled_length + ' ' * (length - filled_length)     # 填充格式
    sys.stdout.write(f'\r{prefix}: [ {bar} ] {percent} {suffix} ')    # 打印输出
    sys.stdout.flush()  # 刷新缓冲


def main():
    """
    参数逻辑解析
    :return: None
    """
    global args
    # 所有函数
    FUNCTIONS = {
        'encrypt': pdf_encrypt,
        'e': pdf_encrypt,
        'decrypt': pdf_decrypt,
        'd': pdf_decrypt,
        'merge': pdf_merge,
        'm': pdf_merge,
        'watermark': pdf_watermark,
        'w': pdf_watermark
    }

    # 获取操作文件
    pdf_files = get_files(args.file)
    # 输出提示信息
    if len(pdf_files) == 0:
        print("# file is not found, please check your path !")
        exit(1)

    print("-----" * 15 + "\n# The following file will be processed:")
    for f in pdf_files:
        print("\t-- " + f[1])
    print("-----" * 15)

    # 功能执行
    if args.function in FUNCTIONS:
        try:
            FUNCTIONS[args.function](pdf_files, args)   # 执行

        except Exception as e:  # 捕获可能非预期的异常
            print("# There something maybe useful:"
                  "\n\t1. You maybe use the wrong arguments, please use \'-h' for help"
                  "\n\t2. There maybe a BUG, please contact to 768476667@qq.com. Thank you!")
            print("# Details wrong are as follows:\n" + str(e))
            exit(1)
    # 功能之外
    else:
        print("# You did not specify a function, please use \'-h\' for help!")
        exit(1)

    print("\n\n### Finished ###")


if __name__ == "__main__":
    # 解析器
    parser = argparse.ArgumentParser(
        prog='pdfHandle',
        epilog='designed by wengui, mail: 768476667@qq.com'
    )

    # 操作原PDF文件 / 目录 / TXT读取文件
    parser.add_argument("file", help="file/directory")
    # 功能选择: 加密  解密  合并  水印
    parser.add_argument("function", help="function to encrypt(e), decrypt(d), merge(m), watermark(w)")
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
