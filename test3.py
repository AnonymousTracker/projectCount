
import pandas as pd
import numpy as np

#
# a = list(range(1, 11))
# a_reshape = np.array(a).reshape(2, 5).T
# b = pd.DataFrame(a_reshape)
# print(b)
# c = b.drop(2, axis=0)        # axis=0，删除行
# print('axis=0:\n', c)
# d = b.drop(0, axis=1)        # axis=1，删除列
# print('axis=1:\n', d)

# # 返回列表索引值
# name = ['序号', '职能组', '项目名称', '对应系统/模块', '测试负责人', '任务状态', '西研测试功能点数', '西研测试工作量\n（人月）', '主协办', '项目经理', '测试经理', '参与人']
# # b = name.index('项目名称')
# # b = name[0:-1]
# a = [0, 0, 0]
#
#
# b = set(a)
# print('\nb:', b)

# import pandas as pd
# df = pd.DataFrame({
#     'brand': ['Yum Yum', 'Yum Yum', 'Indomie', 'Indomie', 'Indomie'],
#     'style': ['cup', 'cup', 'cup', 'pack', 'pack'],
#     'rating': [4, 4, 3.5, 15, 5]
# })
# print('df:\n', df)
# result = []
# result = df.drop_duplicates(subset=['brand', 'style'], keep='last')
# print('\nresult\n', result)

endMark = ['已结束', '末次准出', '无实际工作']
# endMark = ['已结束末次准出无实际工作']
# 如果任务仅包含endMark中状态，则项目已结束，否则在研
status = ['已结束']
# status = set(status)

# if set(status).issubset(set(endMark)):
if set(status) <= set(endMark):
    status_proj = '已结束'
else:
    status_proj = '在研'
print('\nstatus_proj', status_proj)

