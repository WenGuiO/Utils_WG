## :star:1. PDF处理

### :sparkles:功能支持

- 实现对PDF文件进行加密，解密，合并，水印
- 实现从指定目录进行PDF文件读取并进行操作
- 实现对指定批量PDF文件进行操作

### :sparkles:参数说明

位置参数：

- `file`：指定文件或目录。当为PDF文件时，对指定文件进行操作；当为TXT文件时，读取指定批量文件进行操作；当为文件目录时，不指定模式则默认遍历当前目录操作。
- `mode`：功能模式

  - `e`、`encrypt`：加密
  - `d`、`decrypt`：解密
  - `m`、`merge`：合并
  - `w`、`watermark`：水印

可选参数：

- `-r`、`--all-dir` ：对指定目录以及子目录所有PDF文件操作
- `-d`、`--current-dir`：对指定目录的所有PDF文件操作
- `-p`、 `--password`：为加解密指定密码
- `-w` , `--water-file`：指定水印文件

### :sparkles:使用说明

**基本使用：**

```bash
python pdfHandle.py file mode [-r] [-d] [-p] [-w] 
```

**使用示例：**

```yaml
#对指定文件加密
python pdfHandle.py a.pdf encrypt -p 1234

#对指定目录所有文件解密
python pdfHandle.py ./path/a decrypt -p 1234 -d

#合并指定目录以及子目录所有文件
python pdfHandle.py ./path/a merge -r

#为指定文件添加水印, 为 a.pdf 添加 b.pdf 水印
python pdfHandle.py ./path/a.pdf watermakr -w b.pdf

加解密、水印支持单文件以及多文件操作
```

**运行示例：**

```bash
(virenv) PS D:\pycharm\pythonProject> python .\utils\pdfHandle.py .\resource d -p 1234

# Welcome to use the PDF_HANDLE tool that designed by wengui.
---------------------------------------------------------------------------
# The following file will be processed:
        -- .\resource\c.pdf
        -- .\resource\wg.pdf
        -- .\resource\xc.pdf
        -- .\resource\xc2.pdf
# Processing:
---------------------------------------------------------------------------
Process: [ ************************************************** ] 100.0%  

### Finished ###

```

### :sparkles:其他

此为闲暇之余所写，目前只实现了基本功能，暂未实现其他拓展功能。

脚本仍有些不足，有待完善，如：对于pdf文件横页添加水印时存在bug，输出不理想...

同时更多操作可根据需要进行拓展。如：添加指定输出位置以及指定输出名...
