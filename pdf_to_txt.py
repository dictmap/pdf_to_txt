import glob
import pdfplumber
import re


def check_lines(page, top, buttom):
    lines = page.extract_words()[::]
    text = ''
    last_top = 0
    last_check = 0
    for l in range(len(lines)):
        each_line = lines[l]
        check_re = '(?:。|；|单位：元|单位：万元|币种：人民币|\d|报告(?:全文)?(?:（修订版）|（修订稿）|（更正后）)?)$'
        if top == '' and buttom == '':
            if abs(last_top - each_line['top']) <= 2:
                text = text + each_line['text']
            elif last_check > 0 and (page.height * 0.9 - each_line['top']) > 0 and not re.search(check_re, text):

                text = text + each_line['text']
            else:
                text = text + '\n' + each_line['text']
        elif top == '':
            if each_line['top'] > buttom:
                if abs(last_top - each_line['top']) <= 2:
                    text = text + each_line['text']
                elif last_check > 0 and (page.height * 0.85 - each_line['top']) > 0 and not re.search(check_re,
                                                                                                      text):
                    text = text + each_line['text']
                else:
                    text = text + '\n' + each_line['text']
        else:
            if each_line['top'] < top and each_line['top'] > buttom:
                if abs(last_top - each_line['top']) <= 2:
                    text = text + each_line['text']
                elif last_check > 0 and (page.height * 0.85 - each_line['top']) > 0 and not re.search(check_re,
                                                                                                      text):
                    text = text + each_line['text']
                else:
                    text = text + '\n' + each_line['text']
        last_top = each_line['top']
        last_check = each_line['x1'] - page.width * 0.85

    return text


def drop_empty_cols(data):
    # 转置数据，使得每个子列表代表一列而不是一行
    transposed_data = list(map(list, zip(*data)))
    # 过滤掉全部为空的列
    filtered_data = [col for col in transposed_data if not all(cell is '' for cell in col)]
    # 再次转置数据，使得每个子列表代表一行
    result = list(map(list, zip(*filtered_data)))
    return result

def change_pdf_to_txt(name):
    pdf = pdfplumber.open(name)
    last_num = 0

    all_text = {}
    allrow = 0
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        buttom = 0
        tables = page.find_tables()
        if len(tables) >= 1:
            count = len(tables)
            for table in tables:
                if table.bbox[3] < buttom:
                    pass
                else:
                    count = count - 1
                    top = table.bbox[1]
                    text = check_lines(page, top, buttom)
                    text_list = text.split('\n')
                    for _t in range(len(text_list)):
                        all_text[allrow] = {}
                        all_text[allrow] = {'page': page.page_number, 'allrow': allrow, 'type': 'text',
                                            'inside': text_list[_t]}
                        allrow = allrow + 1

                    buttom = table.bbox[3]
                    new_table = table.extract()
                    r_count = 0

                    for r in range(len(new_table)):
                        row = new_table[r]
                        if row[0] == None:
                            r_count = r_count + 1
                            for c in range(len(row)):

                                if row[c] != None and row[c] != '' and row[c] != ' ':
                                    if new_table[r - r_count][c] == None:
                                        new_table[r - r_count][c] = row[c]
                                    else:
                                        new_table[r - r_count][c] = new_table[r - r_count][c] + row[c]
                                    new_table[r][c] = None
                        else:
                            r_count = 0
                    end_table = []
                    for row in new_table:
                        if row[0] != None:
                            cell_list = []
                            for cell in row:
                                if cell != None:
                                    cell = cell.replace('\n', '')
                                else:
                                    cell = ''
                                cell_list.append(cell)
                            end_table.append(cell_list)
                    end_table = drop_empty_cols(end_table)


                    for row in end_table:
                        all_text[allrow] = {'page': page.page_number, 'allrow': allrow, 'type': 'excel',
                                            'inside': str(row)}
                        # all_text[allrow] = {'page': page.page_number, 'allrow': allrow, 'type': 'excel',
                        #                     'inside': ' '.join()}
                        allrow = allrow + 1

                    if count == 0:
                        text = check_lines(page, '', buttom)
                        text_list = text.split('\n')
                        for _t in range(len(text_list)):
                            all_text[allrow] = {'page': page.page_number, 'allrow': allrow, 'type': 'text',
                                                'inside': text_list[_t]}
                            allrow = allrow + 1

        else:
            text = check_lines(page, '', '')
            text_list = text.split('\n')
            for _t in range(len(text_list)):
                all_text[allrow] = {'page': page.page_number, 'allrow': allrow, 'type': 'text',
                                         'inside': text_list[_t]}
                allrow = allrow + 1
        first_re = '[^计](?:报告(?:全文)?(?:（修订版）|（修订稿）|（更正后）)?)$'
        end_re = '^(?:\d|\\|\/|第|共|页|-|_| ){1,}'
        if last_num==0:
            first_text = str(all_text[1]['inside'])
            end_text = str(all_text[len(all_text) - 1]['inside'])
            if re.search(first_re, first_text) and not re.search('\[', first_text):
                all_text[1]['type'] = '页眉'
                if re.search(end_re, end_text) and not re.search('\[', end_text):
                    all_text[len(all_text) - 1]['type'] = '页脚'
        else:
            first_text = str(all_text[last_num + 2]['inside'])

            end_text = str(all_text[len(all_text) - 1]['inside'])

            if re.search(first_re, first_text) and not re.search('\[', first_text):
                all_text[last_num+2]['type'] = '页眉'
            if re.search(end_re, end_text) and not re.search('\[', end_text) :
                all_text[len(all_text) - 1]['type'] = '页脚'

        last_num = len(all_text)-1

    save_path_1 = 'D:\\test_txt2\\'+name.split('\\')[-1].replace('.pdf', '.txt')
    for key in all_text.keys():
        with open(save_path_1, 'a+', encoding='utf-8') as file:
            # file.write(str(all_text[key]['inside']) + '\n')
            file.write(str(all_text[key]) + '\n')


folder_path = 'D:\\test_data2'
file_names = glob.glob(folder_path + '/*')
file_names = sorted(file_names, reverse=True)
name_list = []
for file_name in file_names:
    name_list.append(file_name)
    allname = file_name.split('\\')[-1]
    date = allname.split('__')[0]
    name = allname.split('__')[1]
    year = allname.split('__')[4]
    change_pdf_to_txt(file_name)






