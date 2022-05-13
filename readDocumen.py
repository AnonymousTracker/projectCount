import pandas as pd

dict_path = 'F:\\Python\\100_项目台账整理'  # 文件夹地址
# 项目台账文件名、sheet页名
proj_cur = '\\工作情况汇总.xlsx'  # 读取本期文件
proj_old = '\\工作情况汇总.xlsx'  # 读取上期文件
proj_sheet = '项目台账2'  # 读取sheet页名
main_key = '项目名称'  # 项目台账文件主键
sub_key = '对应系统/模块'

# 系统清单文件名、sheet页名
sys_file = '\\系统领域清单.et'  # 读取文件名
sys_sheet = '系统分类清单'  # 读取sheet页名
sys_main_key = '对应系统/模块'  # 业务领域文件主键
sys_sub_key = '业务领域'  # 业务领域文件主键


# 文件读取
def readFile(dict_name, file, sheet):
    path = dict_name + file
    file_type_label = file.index('.')
    file_len = len(file)
    file_type = file[file_type_label:file_len]
    # 不同文件类型读取
    if file_type == '.xlsx':
        read_file = pd.read_excel(path, sheet, engine='openpyxl')
    elif file_type == '.et':
        read_file = pd.read_excel(path, sheet)
    else:
        print('\n%s 该文件类型无法读取，请转换为.xlsx或者.et文件' % file)
    projData = pd.DataFrame(read_file)
    return projData


# 数据清洗：重建数据列索引，删除空数据、重复数据，重建行索引
def cleanData(sheet_proj, key_fir, key_sec):
    # 重建项目台账，起始列、终止列
    cols_s = 0
    cols_e = sheet_proj.shape[1]  # 除行名外的总列数
    # 重建项目台账，起始行、终止行
    rows_e = sheet_proj.shape[0]  # 除列名外的行总数
    # 读取列索引
    names = sheet_proj.columns.values.tolist()
    # 重建项目台账，起始行
    flag = True
    if key_fir in names:  # 若，'项目名称'所在行是列索引
        pass
    else:  # 若，'项目名称'所在行不是列索引
        for a in range(0, rows_e):  # 遍历DataFrame，获得'项目名称'所在行索引
            for x in range(cols_s, cols_e):
                label_t = sheet_proj.iat[a, x]
                if label_t == key_fir:
                    # [a, x]分别为'项目名称'所在行、列索引
                    # rows_s = a
                    # proj_loc = x
                    sheet_proj.columns = sheet_proj.loc[a].tolist()  # 重新创建列索引，仅在'项目名称'非列索引情况下需要进行
                    sheet_proj = sheet_proj.drop(x, axis=0)  # 删除'项目名称'行列
                    flag = False
                    break
            if not flag:
                break
        # if not flag:
        #     break
    sheet_proj = sheet_proj.dropna(axis=0, subset=[key_fir])  # 列表空行删除，判定依据：第一主键是否为空
    sheet_proj = sheet_proj.drop_duplicates(subset=[key_fir], keep='last')  # 列表重复项目剔除，判定依据：第一、第二主键是否相同
    unna_form = sheet_proj.reset_index(drop=True)  # 重新创建行索引
    return unna_form


# 读取本期项目文件
proj_cur = readFile(dict_path, proj_cur, proj_sheet)
proj_Cur = cleanData(proj_cur, main_key, sub_key)
proj_name1 = proj_Cur.at[2, '项目名称']
print('proj_name1', proj_name1)