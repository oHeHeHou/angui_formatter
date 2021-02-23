import xlrd
import xlwt

'''
把Excel内容转成考试吧需要的导入模板格式
# https://www.kaoshibao.com/
'''
# 要转的文件名称
# file = '信息专业一般工作人员安规题库.xls'
file = '04变电专业一般工作人员安规题库.xls'
print('文件名:{0}'.format(file))
suffix_index = file.rindex('.xls')
file_prefix = file[0:suffix_index]

book = xlrd.open_workbook(file)
all_list = []
total_line_num = 0

sheets = book.sheets()
# new_sheets = []
# new_sheets.append(sheets[1])
# new_sheets.append(sheets[2])
# new_sheets.append(sheets[0])

# for sh in new_sheets:
for sh in sheets:
    total_line_num += sh.nrows
    sh_name = sh.name
    # 分类：判断题，多选题，单选题
    sh_type = sh_name

    subject_list = []
    choice_list = []
    answer_list = []

    subject_header_index = 0
    content_header_index = 0
    choice_header_index = 0
    answer_header_index = 0

    for rx in range(sh.nrows):
        #  第一行，表头
        if rx == 0:
            row_content = sh.row(rx)

            # 记录表头中对应列的序号，用于后面获取内容
            for index, row in enumerate(row_content):
                header = row.value
                if '题型' in header:
                    type_header_index = index
                elif '题目' in header:
                # elif '题干' in header:
                    content_header_index = index
                elif '选项' in header:
                    choice_header_index = index
                elif '正确答案' in header:
                # elif '答案' in header:
                    answer_header_index = index
            continue

        row_content = sh.row(rx)
        # 题目
        row_subject = str(row_content[content_header_index].value).strip()
        row_choice = str(row_content[choice_header_index].value).strip()
        # 选项
        row_choice_list = row_choice.split('|')
        # 答案
        row_answer = str(row_content[answer_header_index].value).strip()

        subject_list.append(row_subject)
        choice_list.append(row_choice_list)
        answer_list.append(row_answer)

    sh_map = {'type': sh_type, 'subject': subject_list, 'choice': choice_list,
              'answer': answer_list}
    all_list.append(sh_map)

# 写入excel
wb = xlwt.Workbook()
ws = wb.add_sheet('1')
# 标题
ws.write(0, 0, '题干')
ws.write(0, 1, '题型')
ws.write(0, 2, '选项 A')
ws.write(0, 3, '选项 B')
ws.write(0, 4, '选项 C')
ws.write(0, 5, '选项 D')
ws.write(0, 6, '选项 E')
ws.write(0, 7, '选项 F')
ws.write(0, 8, '选项 G')
ws.write(0, 9, '正确答案')
ws.write(0, 10, '解析')
ws.write(0, 11, '章节')
ws.write(0, 12, '难度')

subject_index = 1
answer_index = 1
choice_index = 1
for type_dict in all_list:
    subject_list = type_dict['subject']
    choice_list = type_dict['choice']
    answer_list = type_dict['answer']
    subject_type = type_dict['type']
    for subject in subject_list:
        ws.write(subject_index, 0, subject)
        ws.write(subject_index, 1, subject_type)
        subject_index += 1
    for index, choice_ll in enumerate(choice_list):
        for i, choice in enumerate(choice_ll):
            prefix = choice.rindex('-')
            choice = choice[prefix + 1:]
            ws.write(choice_index, i + 2, choice)
        choice_index += 1
    for answer in answer_list:
        ws.write(answer_index, 9, answer)
        answer_index += 1

wb.save('example-out.xls')
