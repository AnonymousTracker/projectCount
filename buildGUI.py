import tkinter
import tkinter.messagebox  # 导入tkinter模块
import xlwt
import xlrd

ytm = tkinter.Tk()  # 创建Tk对象
ytm.title("login")  # 设置窗口标题
ytm.geometry("300x300")  # 设置窗口尺寸
l1 = tkinter.Label(ytm, text="请扫码:")  # 标签
l1.pack()  # 指定包管理器放置组件
user_text = tkinter.Entry()  # 创建文本框
user_text.pack()

Result_text = tkinter.Entry()  # 创建第二个文本框，用于存储结果。
Result_text.pack()


# def wrong():
#    ytm.messagebox.showwarning('出错了~！','请先扫描，再存入数据~！')

def wrong():
    tkinter.messagebox.showinfo('出错了~！', '请先扫描，再存入数据~！')


def listout():
    tkinter.messagebox.showinfo('出错了~！', '这张表存满了！~')


def getuser():
    data = xlrd.open_workbook('saoma.xls')
    table = data.sheets()[0]
    textcon = user_text.get()  # 获取文本框内容
    if user_text.get() == '':
        wrong()
    else:
        nrows = table.nrows
        # 获取行数
        print(nrows)
        if nrows == 0:
            worksheet.write(1, 0, textcon)
            workbook.save('saoma.xls')
        else:
            worksheet.write(nrows, 0, textcon)
            workbook.save('saoma.xls')
    print(nrows)
    print('OK')
    user_text.delete('0', 'end')


def tryagain():
    data = xlrd.open_workbook('saoma.xls')
    table = data.sheets()[0]
    textRe = Result_text.get()  # 获取文本框内容
    row = 0
    if Result_text.get() == '':
        wrong()
    else:
        list = table.col_values(3)
        # 将目标列的值存入列表
        while '' in list:
            list.remove('')
        # 遍历列表，删除列表中的空元素
        ditnum = len(list)
        worksheet.write(ditnum, 3, textRe)
        # 按列写入
        workbook.save('saoma.xls')
    print('OK')
    Result_text.delete('0', 'end')


workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet("sheets1")
# TableHeads = ['条形码','设备类型号','工作模式','测量结果']
TableHeads = ['BarCode', 'Device type number', 'Work mode', 'Measurement result']

col = 0
for i in TableHeads:
    worksheet.write(0, col, i)
    col += 1

workbook.save('saoma.xls')

tkinter.Button(ytm, text="确定", command=getuser).pack()  # command绑定获取文本框内容方法
tkinter.Button(ytm, text="再存", command=tryagain).pack()  # command绑定获取文本框内容方法

ytm.mainloop()  # 进入主循环