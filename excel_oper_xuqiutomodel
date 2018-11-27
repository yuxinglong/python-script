import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
from openpyxl.styles import Color, Fill, Font, Alignment, PatternFill, Border, Side

#打开excel
wb = openpyxl.load_workbook(r'C:\Users\hasee\Desktop\aa***接口需求V3.1.xlsx')
model = openpyxl.load_workbook(r'C:\Users\hasee\Desktop\aa\****物理模型模板.xlsm')


start_row=0
start_col=0
sheet = wb['目录']
for row in range(4, sheet.max_row + 1): #外围遍历列表
    # 获取目录的表中文名
    tb_name_cn = sheet['E' + str(row)].value
    tb_name_en = sheet['D' + str(row)].value
    #表的sheet页
    tb_sheet=wb[tb_name_cn]
    #进入单个sheet页进行遍历
    #得到开始行
    for row in tb_sheet.iter_rows():  # sheet页的每行
        for cell in row:
            if (cell.value == '表中文名'):
                #得到开始行
                start_row=cell.row+3
                #得到开始列
                start_col=cell.column
                break

    #开始遍历每一行
    list = []
    # 将daatdt加到第一行
    list.append('数据日期')
    list.append('')
    list.append('')
    list.append('')
    list.append('数据日期')
    list.append('DATA_DT')
    list.append('DATE')
    list.append('PA')
    for row in range(start_row, wb[tb_name_cn].max_row + 1):  #外围遍历列表  # 单个sheet页的每行
        #不是行末
        if not tb_sheet[start_col+str(row)].value is None:
            #获取字段
            col_en=str(tb_sheet[start_col+str(row)].value)
            #列+1获取字段中文名
            col_cn=str(tb_sheet[get_column_letter(column_index_from_string(start_col)+1)+str(row)].value)
            #列+2获取字段类型
            col_type=str(tb_sheet[get_column_letter(column_index_from_string(start_col)+2)+str(row)].value)
            list.append(col_cn)
            list.append('')
            list.append('')
            list.append('')
            list.append(col_cn)
            list.append(col_en)
            list.append(col_type)
            list.append('')


#生成
    panel_sheet=model['模板']
    stg_sheet = model['STG']
    panel_list=[]
    for row in range(4,8):  # sheet页的每行
        for col in range(2,21):
            panel_list.append(panel_sheet[get_column_letter(col)+ str(row)].value)
    #开始STG的写入,设置stg开始位置
    stg_start_row=stg_sheet.max_row+3
    #存储初始值
    store_stg_start_row=stg_start_row
    content_start_row=stg_start_row
    #写入STGsheet
    i=0
    for row in range(stg_start_row,stg_start_row+4):  # sheet页的每行
        for col in range(2,21):
            if not panel_list[i] is None:
                stg_sheet[get_column_letter(col)+str(row)]=str(panel_list[i])
            else:
                stg_sheet[get_column_letter(col) + str(row)] = ''
            i=i+1
    #补充中文表名和英文表名
    stg_sheet['C' + str(store_stg_start_row+1)] = tb_name_cn
    stg_sheet['G' + str(store_stg_start_row+1)] = 'S_'+tb_name_en
    #合并单元格
    start='B'+str(store_stg_start_row)
    end='D'+str(store_stg_start_row)
    stg_sheet.merge_cells('B'+str(store_stg_start_row)+':'+'D'+str(store_stg_start_row))
    stg_sheet.merge_cells('F'+str(store_stg_start_row)+':'+'T'+str(store_stg_start_row))
    two_row=store_stg_start_row+1
    stg_sheet.merge_cells('C'+str(two_row)+':'+'D'+str(two_row))
    stg_sheet.merge_cells('G' + str(two_row) + ':' + 'T' + str(two_row))

    # 设置边框
    border = Border(left=Side(style='thin', color='FF000000'), right=Side(style='thin', color='FF000000'),
                    top=Side(style='thin', color='FF000000'), bottom=Side(style='thin', color='FF000000'),
                    diagonal=Side(style='thin', color='FF000000'), diagonal_direction=0,
                    outline=Side(style='thin', color='FF000000'), vertical=Side(style='thin', color='FF000000'),
                    horizontal=Side(style='thin', color='FF000000'))
    #设置左边边框的样式
    for row in range(store_stg_start_row, store_stg_start_row + 4):
        for col in range(2, 5):
            stg_sheet[get_column_letter(col) + str(row)].border = border
            #左边边框每个格子的边框设置
            if (row == store_stg_start_row and col == 2):
                stg_sheet[get_column_letter(col) + str(row)].border =Border(left=Side(style='medium'),
                                                                            right=Side(style='thin'),
                                                                            top=Side(style='medium'),
                                                                            bottom=Side(style='thin'))
            if (row == store_stg_start_row and col == 3 ):
                stg_sheet[get_column_letter(col) + str(row)].border =Border(left=Side(style='thin'),
                                                                            right=Side(style='thin'),
                                                                            top=Side(style='medium'),
                                                                            bottom=Side(style='thin'))
            if (row == store_stg_start_row and col == 4):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                            right=Side(style='medium'),
                                                                            top=Side(style='medium'),
                                                                            bottom=Side(style='thin'))
            if (row == store_stg_start_row + 1 and col == 2):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='medium'),
                                                                            right=Side(style='thin'),
                                                                            top=Side(style='thin'),
                                                                            bottom=Side(style='thin'))
            if (row == store_stg_start_row + 1 and col == 3):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='thin'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row + 1 and col == 4):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                            right=Side(style='medium'),
                                                                            top=Side(style='thin'),
                                                                            bottom=Side(style='thin'))
            if (row == store_stg_start_row + 2 and col == 2):
                stg_sheet[get_column_letter(col) + str(row)].border =Border(left=Side(style='medium'),
                                                                            right=Side(style='thin'),
                                                                            top=Side(style='thin'),
                                                                            bottom=Side(style='thin'))
            if (row == store_stg_start_row + 2 and col == 3):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='thin'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row + 2 and col == 4):
                stg_sheet[get_column_letter(col) + str(row)].border =Border(left=Side(style='thin'),
                                                                            right=Side(style='medium'),
                                                                            top=Side(style='thin'),
                                                                            bottom=Side(style='thin'))
            if (row == store_stg_start_row+3 and col == 2):
                stg_sheet[get_column_letter(col) + str(row)].border =Border(left=Side(style='medium'),
                                                                            right=Side(style='thin'),
                                                                            top=Side(style='thin'),
                                                                            bottom=Side(style='medium'))
            if (row == store_stg_start_row+3 and col == 3):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                            right=Side(style='thin'),
                                                                            top=Side(style='thin'),
                                                                            bottom=Side(style='medium'))
            if (row == store_stg_start_row+3 and col == 4):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                            right=Side(style='medium'),
                                                                            top=Side(style='thin'),
                                                                            bottom=Side(style='medium'))
            #左边方框LDM居中显示
            if (row == store_stg_start_row and col == 2):
                stg_sheet[get_column_letter(col) + str(row)].alignment  = Alignment(horizontal='center',vertical='center',wrap_text=True)

    #设置右边方框的样式
    for row in range(store_stg_start_row, store_stg_start_row + 4):  # sheet页的每行
        for col in range(6, 21):
            stg_sheet[get_column_letter(col) + str(row)].border = border
            #右边方框的样式
            if (row == store_stg_start_row and col == 6):
                stg_sheet[get_column_letter(col) + str(row)].border =Border(left=Side(style='medium'),
                                                                            right=Side(style='thin'),
                                                                            top=Side(style='medium'),
                                                                            bottom=Side(style='thin'))
            if (row == store_stg_start_row and (col == 7 or col == 8 or col == 9 or
                                                col == 10 or col == 11 or col == 12 or
                                                col == 13 or col == 14 or col == 15 or
                                                col == 16 or col == 17 or col == 18 or
                                                col == 19 )):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='thin'),
                                                                             top=Side(style='medium'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row and col == 20):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='medium'),
                                                                             top=Side(style='medium'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row+1 and col == 6):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='medium'),
                                                                             right=Side(style='thin'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row+1 and (col == 7 or col == 8 or col == 9 or
                                                        col == 10 or col == 11 or col == 12 or
                                                        col == 13 or col == 14 or col == 15 or
                                                        col == 16 or col == 17 or col == 18 or
                                                        col == 19)):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='thin'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row+1 and col == 20):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='medium'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row + 2 and col == 6):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='medium'),
                                                                             right=Side(style='thin'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row + 2 and (col == 7 or col == 8 or col == 9 or
                                                            col == 10 or col == 11 or col == 12 or
                                                            col == 13 or col == 14 or col == 15 or
                                                            col == 16 or col == 17 or col == 18 or
                                                            col == 19)):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='thin'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row + 2 and col == 20):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='medium'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='thin'))
            if (row == store_stg_start_row + 3 and col == 6):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='medium'),
                                                                             right=Side(style='thin'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='medium'))
            if (row == store_stg_start_row + 3 and (col == 7 or col == 8 or col == 9 or
                                                            col == 10 or col == 11 or col == 12 or
                                                            col == 13 or col == 14 or col == 15 or
                                                            col == 16 or col == 17 or col == 18 or
                                                            col == 19)):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='thin'),
                                                                             top=Side(style='thin'),
                                                                             bottom=Side(style='medium'))
            if (row == store_stg_start_row + 3 and col == 20):
                stg_sheet[get_column_letter(col) + str(row)].border = Border(left=Side(style='thin'),
                                                                             right=Side(style='medium'),
                                                                             top=Side(style='thin'),
                                                                        bottom=Side(style='medium'))
             # 右边PDM居中显示
            if (row == store_stg_start_row  and col == 6 ):
                stg_sheet[get_column_letter(col) + str(row)].alignment  = Alignment(horizontal='center',vertical='center',wrap_text=True)

    val_start_row=store_stg_start_row+4
    val_start_col=2
    for i in list:
        stg_sheet[get_column_letter(val_start_col) + str(val_start_row)] = i
        if(val_start_col%9==0):
            val_start_col=2
            val_start_row = val_start_row + 1
        else:
            val_start_col = val_start_col + 1

model.save(r'C:\Users\hasee\Desktop\aa\生成结果.xlsx')






























