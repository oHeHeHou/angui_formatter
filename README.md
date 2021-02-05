# angui_formatter
将电力安规题库Excel文件转换为可读性更好的Word文件

原有格式：

![](.imgs/src.png)

转换后格式：

![](.imgs/out.png)


## 说明
1. excel_parser.py 将原始题库excel转换为word文档
2. template_generator.py 将原始题库excel转换为考试吧可以导入的excel模板

## 依赖

* [python-docx](https://github.com/python-openxml/python-docx)
* [xlrd](https://github.com/python-excel/xlrd)
* [xlwt](https://github.com/python-excel/xlwt)
