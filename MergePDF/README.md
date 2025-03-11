# MergePDF
把多个PDF文件合并成一个PDF文件

把所有需要转换的word文档放在同一个文件夹，然后把BatchDocxToPDF.vbs文件和文档放在同一个文件夹下，例如：

```
新建文件夹/
│
├── merge.py
├── 单面打印/
├── 双面打印/
```

双击执行代码后，代码会创建 merge_output 文件夹，然后把所有的合并的pdf文件存到里面

```
新建文件夹/
│
├── merge.py
├── 单面打印/
├── 双面打印/
├── merge_output/
│   ├── 单面打印.pdf
│   ├── 双面打印.pdf
```