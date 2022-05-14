import pandas as pd

dict_path = 'F:\\Python\\100_项目台账整理\\TestData'  # 文件夹地址
proj_cur = '\\工作情况汇总.xlsx'  # 读取本期文件
proj_old = '\\工作情况汇总-上期.xlsx'  # 读取上期文件
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
                    sheet_proj.columns = sheet_proj.loc[a].tolist()  # 重新创建列索引，仅在'项目名称'非列索引情况下需要进行
                    sheet_proj = sheet_proj.drop(a, axis=0)  # 删除'项目名称'行列
                    flag = False
                    break
            if not flag:
                break
    sheet_proj = sheet_proj.dropna(axis=0, subset=[key_fir])  # 列表空行删除，判定依据：第一主键是否为空
    sheet_proj = sheet_proj.drop_duplicates(subset=[key_fir, key_sec], keep='last')  # 列表重复项目剔除，判定依据：第一、第二主键是否相同
    unna_form = sheet_proj.reset_index(drop=True)  # 重新创建行索引
    return unna_form


# 读取上期项目文件
proj_Old = readFile(dict_path, proj_old, proj_sheet)
proj_Old = cleanData(proj_Old, main_key, sub_key)
# print('proj_Cur：\n', proj_Old)

# 读取本期项目文件
proj_Cur = readFile(dict_path, proj_cur, proj_sheet)
proj_Cur = cleanData(proj_Cur, main_key, sub_key)
# print('proj_Cur：\n', proj_Cur)

# 读取系统领域清单
sys_table_ori = readFile(dict_path, sys_file, sys_sheet)
sys_table = cleanData(sys_table_ori, sys_main_key, sys_sub_key)
# print('sys_table: \n', sys_table)
sys_list = sys_table['对应系统/模块'].tolist()  # 项目所属系统/模块
# print('sys_list: \n', sys_list)
business_area = sys_table['业务领域'].tolist()  # 系统/模块所属业务领域


# 判断是否可转换字符串
def isNumber(target_str):
    try:
        float(target_str)
        return True
    except:
        pass
    if target_str.isnumeric():
        return True
    return False


# 工作量类型转换
def amouWoTypeChange(proj_cur):
    """
    1. 工作量为空
    2. 工作量为int or float
    3. 工作量为str类型
    3.1 可转换为浮点型
    3.2 不可转换为浮点型
    4. 其他情况
    """
    amount_work = proj_cur['西研测试工作量\n（人月）'].tolist()
    for x in range(0, len(amount_work)):
        if pd.isna(amount_work[x]):
            proj_cur.at[x, '西研测试工作量\n（人月）'] = 0
        elif type(amount_work[x]) == str:
            if isNumber(amount_work[x]) is True:
                proj_cur.at[x, '西研测试工作量\n（人月）'] = float(amount_work[x])
            else:
                proj_cur.at[x, '西研测试工作量\n（人月）'] = 0
        elif type(amount_work[x]) is int or float:
            continue
        else:
            proj_cur.at[x, '西研测试工作量\n（人月）'][x] = 0
    return proj_cur


# 功能点类型转换
def funPoinTypeChange(proj_cur):
    """
    1.功能点为空
    2. 功能点int or float
    3. 功能点为str
    3.1 可转换为数字
    3.2 不可转换为数字
    4.其他
    """
    function_point = proj_cur['西研测试功能点数'].tolist()
    amount_work = proj_cur['西研测试工作量\n（人月）'].tolist()
    constant = 184.6
    for i in range(0, len(function_point)):
        # print('\nfunction_point[i]:', function_point[i])
        if pd.isna(function_point[i]):
            proj_cur.at[i, '西研测试功能点数'] = amount_work[i] * constant
        elif type(function_point[i]) == str:
            if isNumber(function_point[i]) is True:
                proj_cur.at[i, '西研测试功能点数'] = float(function_point[i])
            else:
                proj_cur.at[i, '西研测试功能点数'] = amount_work[i] * constant
        elif type(function_point[i]) == int or float:
            continue
        else:
            proj_cur.at[i, '西研测试功能点数'][i] = 0
    return proj_cur


# 项目状态出现非空后，后续为空项目状态均调整为“未开始”
def changeStatus(proj_cur):
    statusList = proj_cur['任务状态'].tolist()
    for i in range(0, len(statusList)):
        if pd.isna(statusList[i]):
            continue
        else:
            for j in range(i, len(statusList)):
                if pd.isna(statusList[j]):
                    proj_cur.at[j, '任务状态'] = '未开始'
                else:
                    continue
    return proj_cur


proj_Cur = amouWoTypeChange(proj_Cur)
proj_Cur = funPoinTypeChange(proj_Cur)
proj_Cur = changeStatus(proj_Cur)


# 将项目台账数据按列转换为列表
def newDataFrame(proj_from):
    proj_name = proj_from['项目名称'].tolist()  # 项目名称
    charge_group = proj_from['职能组'].tolist()  # 项目负责职能组
    system_name = proj_from['对应系统/模块'].tolist()  # 项目对应系统/模块
    charge_man = proj_from['测试负责人'].tolist()  # 项目负责人
    status = proj_from['任务状态'].tolist()  # 项目状态
    amount_work = proj_from['西研测试工作量\n（人月）'].tolist()  # 西研测试工作量
    function_point = proj_from['西研测试功能点数'].tolist()  # 西研测试功能点数
    project_role = proj_from['主协办'].tolist()  # 项目主协办
    data_frame = [proj_name, charge_group, system_name, charge_man, status, amount_work, function_point, project_role]
    return data_frame


# 判断单个项目/任务状态
def getStatus(a, statustemp):
    endMark = ['已结束', '末次准出', '无实际工作']
    exeMark = ['执行中']
    # 如果任务仅包含endMark中状态，则项目已结束，否则在研
    status = []
    # a = [17, 18, 19]
    if type(a) is int:
        status = [statustemp[a]]
    else:
        for i in range(0, len(a)):
            x = a[i]
            status.append(statustemp[x])

    if set(status) <= set(endMark):     # 运算符<=表示status是endMark的子集
        status_proj = '已结束'
    elif set(exeMark) <= set(status):
        status_proj = '执行中'
    else:
        status_proj = '未开始'
    return status_proj


# 统计任务数、功能点数
def getTaskData(x, y, function_point, num_list, fun_list):  # x：业务领域编号；  y：任务/项目遍历号
    # print('统计前各领域任务数: %s' % num_list)
    # print('统计前各领域功能点数: %s' % fun_list)
    """
    如果 功能点为0:
            仅任务数累加
    否则 功能点不为0：
            任务数累加，工作量转换为功能点累加
    """
    if function_point[y] == 0:
        num_list[x] = num_list[x] + 1
    else:
        num_list[x] = num_list[x] + 1
        fun_list[x] = fun_list[x] + function_point[y]
    # print('统计后各领域任务数: %s' % num_list)
    # print('统计后各领域功能点数: %s' % fun_list)
    return [num_list, fun_list]


# 系统/模块按不同业务领域分类，获得不同领域系统/模块清单
def sys_area(sys_list, business_area):
    # 不用业务领域承接任务数量统计
    supervi_list = []  # 监管
    antimonlaun_list = []  # 反洗钱
    finman_list = []  # 理财子公司
    assliab_list = []  # 资产与负债
    boeing_list = []  # BoEing下移
    for a in range(0, len(sys_list)):
        if business_area[a] == '监管':
            supervi_list.append(sys_list[a])
        elif business_area[a] == '反洗钱':
            antimonlaun_list.append(sys_list[a])
        elif business_area[a] == '理财子公司':
            finman_list.append(sys_list[a])
        elif business_area[a] == '资产负债':
            assliab_list.append(sys_list[a])
        elif business_area[a] == 'BoEing下移':
            boeing_list.append(sys_list[a])
        else:
            pass
    return supervi_list, antimonlaun_list, finman_list, assliab_list, boeing_list


# 构建不同业务领域所属系统/模块清单(监管, 反洗钱, 理财子公司, 资产与负债, BoEing下移)
areaDetail = sys_area(sys_list, business_area)


# 判定项目/任务业务领域
def judBusinArea(y, project_name, function_point, num_list, fun_list, system_name, area_detail):
    """
    1. 项目对应系统/模块在【系统领域清单】中存在；
    1.1 根据项目对应系统/模块，确定系统业务领域
    2.项目对应系统/模块在【系统领域清单】中不存在；
    2.1 前台展示项目名称、项目对应系统/模块、负责人，获取系统/模块所属业务领域信息，对项目进行分类；
    2.2 将所属业务领域信息保存至sys_list
    """
    supervi_list = area_detail[0]
    antimonlaun_list = area_detail[1]
    finman_list = area_detail[2]
    assliab_list = area_detail[3]
    boeing_list = area_detail[4]
    # 监管领域
    if system_name[y] in supervi_list:
        [num_list, fun_list] = getTaskData(0, y, function_point, num_list, fun_list)
    # 反洗钱领域
    elif system_name[y] in antimonlaun_list:
        [num_list, fun_list] = getTaskData(1, y, function_point, num_list, fun_list)
    # 理财子公司领域
    elif system_name[y] in finman_list:
        [num_list, fun_list] = getTaskData(2, y, function_point, num_list, fun_list)
    # 资产负债领域
    elif system_name[y] in assliab_list:
        [num_list, fun_list] = getTaskData(3, y, function_point, num_list, fun_list)
    # BoEing下移领域
    elif system_name[y] in boeing_list:
        [num_list, fun_list] = getTaskData(4, y, function_point, num_list, fun_list)
    elif pd.isna(system_name[y]):  # 项目/任务对应系统/模块名称为空
        print('以下任务未添加对应系统/模块，请在《项目台账》中添加:\n'
              '%s\n' % project_name[y])
    else:
        print('以下系统模块未确定业务领域，请在《系统领域清单》中添加:\n'
              '%s\n' % system_name[y])
    # print('各领域任务数: %s' % num_list)
    # print('各领域功能点数: %s' % fun_list)
    return num_list, fun_list


# 获得同一项目不同任务索引
def getDualProjIndexIndex(L, f):
    # L表示列表,x表示索引值，v表示values，f表示要查找的元素
    proj_index = [x for (x, v) in enumerate(L) if v == f]
    return proj_index


# 判断项目属性，包括职能组、主协办，统计出不同职能组项目总数、主办总数、协办总数
def projNumCount(proj_cur, a, proj_num_list):
    [project_name, charge_group, system_name, charge_man, status, amount_work, function_point,
     project_role] = newDataFrame(proj_cur)
    [proj_num, first_group, first_group_main, first_group_assist, first_group_exe, sec_group, sec_group_main,
     sec_group_assist, sec_group_exe] = proj_num_list
    # 确定项目领域
    proj_num = proj_num + 1
    if charge_group[a] == '测试一组':
        first_group = first_group + 1
        if status[a] == '执行中':
            first_group_exe = first_group_exe + 1
        if project_role[a] == '主办':
            first_group_main = first_group_main + 1
        else:
            first_group_assist = first_group_assist + 1
    else:
        sec_group = sec_group + 1
        if status[a] == '执行中':
            sec_group_exe = sec_group_exe + 1
        if project_role[a] == '主办':
            sec_group_main = sec_group_main + 1
        else:
            sec_group_assist = sec_group_assist + 1
    proj_num_list = [proj_num, first_group, first_group_main, first_group_assist, first_group_exe, sec_group,
                     sec_group_main, sec_group_assist, sec_group_exe]
    return proj_num_list


# 对比本期、上期项目台账，获得项目新增、准出统计数据
def updateProj(proj_old, proj_cur, update_num):
    out_num = update_num[0]
    insert_num = update_num[1]
    old_name = proj_old['项目名称'].tolist()
    old_status = proj_old['任务状态'].tolist()
    cur_name = proj_cur['项目名称'].tolist()
    cur_status = proj_cur['任务状态'].tolist()
    '''
    1. 项目编号已存在于本期同名项目索引中，跳过
    2.1 项目在研，且不存在于上期项目，新增递增
    2.2 项目已结束，且不存在于上期项目，新增、准出递增
    2.3 项目已结束，且项目上期状态为在研，准出递增
    2.4 其他情况，跳过
    '''
    for i in range(0, len(cur_name)):
        proj_index_cur = getDualProjIndexIndex(cur_name, cur_name[i])  # 本期同名项目索引
        status_cur = getStatus(proj_index_cur, cur_status)  # 本期项目状态
        proj_index_old = getDualProjIndexIndex(old_name, cur_name[i])  # 待判定项目在上期项目台账中的索引
        status_old = getStatus(proj_index_old, old_status)  # 上期项目状态
        if i != proj_index_cur[0] and i in proj_index_cur:
            continue
        elif cur_name[i] not in old_name and status_cur != '已结束':
            insert_num = insert_num + 1
        elif cur_name[i] not in old_name and status_cur == '已结束':
            out_num = out_num + 1
            insert_num = insert_num + 1
        elif status_cur == '已结束' and status_cur != status_old:
            out_num = out_num + 1
        else:
            continue
    return out_num, insert_num


updateNum = [0, 0]
[outNum, insertNum] = updateProj(proj_Old, proj_Cur, updateNum)
print('\n本期准出项目%d个，准入项目%d个\n' % (outNum, insertNum))


# 统计项目数据
def projCount(proj_cur, proj_num, area_num, proj_fun):
    [project_name, charge_group, system_name, charge_man, status, amount_work, function_point,
     project_role] = newDataFrame(proj_cur)
    for i in range(0, len(project_name)):
        projIndex = getDualProjIndexIndex(project_name, project_name[i])  # 获得同一项目所有索引
        projNameCunt = len(projIndex)
        statusProj = getStatus(projIndex, status)  # 获得项目的项目状态
        '''
        1.项目已被统计，跳过遍历
        2.项目仅一个任务，且非'已结束'
        3.项目多个任务，且非'已结束'
        3.1 存在主办，则项目计入主办所在职能组、业务领域
        3.2 不存在主办
        3.2.1 不存在主办，任务功能点数均不为0，获得功能点数最大协办索引，项目归属于最大功能点对应职能组、业务领域
        3.2.2 不存在主办，且功能点数为0，项目归属于第一个任务对应职能组、业务领域
        4.其他情形，跳过遍历
        '''
        if i != projIndex[0] and i in projIndex:
            continue    # 1.项目已被统计，跳过遍历
        elif projNameCunt == 1 and statusProj != '已结束':     # 2.项目仅一个任务，且非'已结束'
            proj_num = projNumCount(proj_cur, i, proj_num)
            [area_num, proj_fun] = judBusinArea(i, project_name, function_point, area_num, proj_fun,
                                                system_name, areaDetail)
        elif projNameCunt > 1 and statusProj != '已结束':     # 3.项目多个任务，且非'已结束'
            projectRoleTemp = [project_role[a] for a in projIndex]
            functionPointTemp = [function_point[a] for a in projIndex]
            if '主办' in projectRoleTemp:     # 3.1 存在主办，则项目计入主办所在职能组、业务领域
                jTemp = projectRoleTemp.index('主办')
                if type(jTemp) is not int:
                    j = projIndex[jTemp[0]]
                else:
                    j = projIndex[jTemp]
                proj_num = projNumCount(proj_cur, j, proj_num)
                [area_num, proj_fun] = judBusinArea(j, project_name, function_point, area_num, proj_fun,
                                                    system_name, areaDetail)
            elif sum(functionPointTemp) != 0:     # 3.1 不存在主办，任务功能点数均不为0，获得功能点数最大协办索引，项目归属于最大功能点对应职能组、业务领域
                jTemp = functionPointTemp.index(max(functionPointTemp))
                j = projIndex[jTemp]
                proj_num = projNumCount(proj_cur, j, proj_num)
                [area_num, proj_fun] = judBusinArea(j, project_name, function_point, area_num, proj_fun,
                                                    system_name, areaDetail)
            else:     # 3.2 不存在主办，且功能点数为0，项目归属于第一个任务对应职能组、业务领域
                j = projIndex[0]
                proj_num = projNumCount(proj_cur, j, proj_num)
                [area_num, proj_fun] = judBusinArea(j, project_name, function_point, area_num, proj_fun,
                                                    system_name, areaDetail)
        else:
            continue    # 4.其他情形，跳过遍历
    return proj_num, area_num, proj_fun


# 不同业务领域项目数、不同业务领域项目数/功能点数
projNumList = [0, 0, 0, 0, 0, 0, 0, 0, 0]
projAreaNum = [0, 0, 0, 0, 0]
projFunList = [0, 0, 0, 0, 0]
[projNumList, projAreaNum, projFunList] = projCount(proj_Cur, projNumList, projAreaNum, projFunList)
[projNum, firstGroup, firstGroupMain, firstGroupAssist, firstGroupExe, secondGroup, secondGroupMain,
 secondGroupAssist, secondGroupExe] = projNumList
[superviProjNum, antimonlaunProjNum, finmanProjNum, assliabProjNum, boeingProjNum] = projAreaNum
projMainNum = firstGroupMain + secondGroupMain
projAssistNum = firstGroupAssist + secondGroupAssist
'''
判定条件
项目总数 = 主办 + 协办 = 一组项目数 + 二组项目数
项目总数 = 一组项目数 + 二组项目数
'''

# 项目统计报告
print('目前正在推进项目 %d 个。其中：\n'
      '1）主办 %d 个，协办 %d 个；\n'
      '2）一组 %d 个。其中主办 %d 个，协办 %d 个, 执行中 %d 个.\n'
      '3）二组 %d 个。其中主办 %d 个，协办 %d 个, 执行中 %d 个.\n'
      % (projNum, projMainNum, projAssistNum, firstGroup, firstGroupMain, firstGroupAssist, firstGroupExe,
         secondGroup, secondGroupMain, secondGroupAssist, secondGroupExe))


print('项目数量按业务领域分布：\n'
      '监管  反洗钱  理财子公司 资产负债  BoEing下移\n'
      ' %d      %d       %d       %d       %d\n' % (superviProjNum, antimonlaunProjNum, finmanProjNum, assliabProjNum,
                                                    boeingProjNum))

# 不同领域任务统计数
def taskCount(proj_cur, area_detail, task_num, task_fun):
    [project_name, charge_group, system_name, charge_man, status, amount_work, function_point,
     project_role] = newDataFrame(proj_cur)
    for i in range(0, len(project_name)):
        '''
        1. 项目名称为空，跳过；
        2. 项目名称不为空：
        2.1 项目'在研'
        2.2 其他情况，跳过
        '''
        statusTask = getStatus(i, status)
        if pd.isna(project_name[i]):
            continue
        elif statusTask != '已结束':
            [task_num, task_fun] = judBusinArea(i, project_name, function_point, task_num, task_fun, system_name,
                                                area_detail)
        else:
            continue
    return task_num, task_fun


# 不同领域承接任务数量、功能点
taskNumList = [0, 0, 0, 0, 0]  # 不同业务领域，任务个数统计列表
taskFunList = [0, 0, 0, 0, 0]  # 不同业务领域，任务功能点数统计列表
[taskNumList, taskFunList] = taskCount(proj_Cur, areaDetail, taskNumList, taskFunList)
[superviTaskNum, antimonlaunTaskNum, finmanTaskNum, assliabTaskNum, boeingTaskNum] = taskNumList
[superviTaskFun, antimonlaunTaskFun, finmanTaskFun, assliabTaskFun, boeingTaskFun] = taskFunList
# 不同业务领域项目分布
print(
    '任务数量按业务领域分布：\n'
    '监管  反洗钱  理财子公司 资产负债  BoEing下移\n'
    ' %d      %d       %d       %d       %d\n'
    '\n任务功能点按业务领域分布：\n'
    ' 监管     反洗钱    理财子公司 资产负债  BoEing下移\n'
    '%d      %d         %d       %d       %d\n' % (superviTaskNum, antimonlaunTaskNum, finmanTaskNum, assliabTaskNum,
                                                   boeingTaskNum, superviTaskFun, antimonlaunTaskFun, finmanTaskFun,
                                                   assliabTaskFun, boeingTaskFun))
