import datetime
import xlrd
import xlwt
from mysqlhelper import *
import StoreInfoObj


# 操作工作簿
def operateWorkbook(data):
    print("------ 操作工作簿 ------")
    print(data.sheets())  # 获取所有的sheet
    print(data.sheets()[0])  # 获取索引为0的sheet
    print(data.sheet_by_index(0))  # 获取索引为0的sheet
    print(data.sheet_names())  # 获取所有工作表的名字
    print(data.sheet_by_name("Sheet1"))  # 获取名字为Sheet1的工作表
    print(data.nsheets)  # 返回工作表的数量


# 操作excel行
def operateRows(data):
    print("------ 操作excel行 ------")
    sheet = data.sheet_by_index(0)  # 获取第一个工作表
    print(sheet.nrows)  # 获取sheet下的有效行数
    print(sheet.row_types(0))  # 获取指定行的单元格的数据类型 0 empty，1 string，2 number，3 date，4 boolean，5 error
    print(sheet.row(0))  # 获取指定行
    print(sheet.row(0)[0])  # 获取指定单元格
    print(sheet.row(0)[0].value)  # 获取指定单元格的值
    print(sheet.row_values(0))  # 获取指定行单元格的value
    print(sheet.row_len(1))  # 得到单元格的长度


# 操作excel列
def operateColumns(data):
    print("------ 操作excel列 ------")
    sheet = data.sheet_by_index(0)  # 获取第一个工作表
    print(sheet.ncols)  # 获取sheet下的有效列数
    print(sheet.col_types(0))  # 获取指定列的单元格的数据类型 0 empty，1 string，2 number，3 date，4 boolean，5 error
    print(sheet.col(0))  # 获取指定列
    print(sheet.col(0)[0])  # 获取指定单元格
    print(sheet.col(0)[0].value)  # 获取指定单元格的值
    print(sheet.col_values(0))  # 获取指定列单元格的value


# 操作excel单元格
def operateCells(data):
    print("------ 操作excel单元格 ------")
    sheet = data.sheet_by_index(0)  # 获取第一个工作表
    print(sheet.cell(0, 0))  # 获取指定单元格
    print(sheet.cell(0, 0).value)  # 获取指定单元格的值
    print(sheet.cell_value(0, 0))  # 获取指定单元格的值
    print(sheet.cell(0, 0).ctype)  # 获取指定单元格的数据类型
    print(sheet.cell_type(0, 0))  # 获取指定单元格的数据类型
    print("true" if sheet.cell_type(0, 0) == int(1) else "false")  # 三目运算符 【为真时候的值 if 条件 else 为假时候的值】


# 数据清洗逻辑
def clean_data(data):
    source_list = []  # 定义一个空列表来接收原数据
    target_list = []  # 定义一个空列表来接收处理过的数据
    sheet = data.sheet_by_index(0)  # 获取第一个工作表
    # 将读出的数据处理一下，去掉标题行后转为列表的形式
    for i in range(sheet.nrows):
        if i > 0:
            source_list.append(sheet.row_values(i))
    # 如果列表有数据
    if len(source_list) > 0:
        for i in range(len(source_list)):
            data_店铺名称 = str(source_list[i][0])
            data_店铺链接 = str(source_list[i][1])
            data_店铺地址 = str(source_list[i][2])
            data_人均消费 = float(str(source_list[i][3]).replace('¥', '')) if str(source_list[i][3]).replace('¥',
                                                                                                         '').isdigit() else str(
                source_list[i][3])
            data_店铺评分 = float(str(source_list[i][4]).replace('分', '')) if str(source_list[i][4]).replace('分',
                                                                                                         '').replace(
                '.', '').isdigit() else str(source_list[i][4])
            data_评论数 = int(str(source_list[i][5]).replace('人评论', '')) if str(source_list[i][5]).replace('人评论',
                                                                                                        '').isdigit() else str(
                source_list[i][5])
            target_list.append([data_店铺名称, data_店铺链接, data_店铺地址, data_人均消费, data_店铺评分, data_评论数])
    return target_list


# 填充数据
def write_excel(target_list):
    # 合并行序列为0到1，列序列为0到5的单元格
    ws.write_merge(0, 1, 0, 5, "wx奶茶店铺数据202204", title_style)
    title = ("店铺名称", "店铺链接", "店铺地址", "人均消费", "店铺评分", "评论数")
    # 写入标题行
    for i, item in enumerate(title):
        ws.write(2, i, item, subtitle_style)
    # 写入数据行
    for i, item in enumerate(target_list):
        for j, val in enumerate(item):
            ws.write(i + 3, j, val, text_style)
    # 补充图片插入
    ws.insert_bitmap("a3.bmp", len(target_list) + 3, 0)


# 设置字体样式
def setTitleStyle():
    # 标题
    # 实例化字体样式
    title_font = xlwt.Font()
    # 字体
    title_font.name = "宋体"
    # 加粗
    title_font.bold = True
    # 字号
    title_font.height = 12 * 20
    # 颜色 可以从xlwt.XFStyle()中找，也可以直接对应索引
    title_font.colour_index = 0x08
    # 应用样式
    title_style.font = title_font
    # 实例化对齐方式
    cell_align = xlwt.Alignment()
    # 水平居中
    cell_align.horz = 0x02
    # 垂直居中
    cell_align.vert = 0x01
    # 应用对齐方式
    title_style.alignment = cell_align
    # 实例化边框
    borders = xlwt.Borders()
    borders.right = xlwt.Borders.THIN
    borders.left = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    # 应用边框
    title_style.borders = borders
    # 背景样式
    # 实例化背景样式
    bgcolor = xlwt.Pattern()
    bgcolor.pattern = xlwt.Pattern.SOLID_PATTERN
    bgcolor.pattern_fore_colour = 22
    # 应用背景样式
    title_style.pattern = bgcolor
    # 副标题
    # 实例化字体样式
    subtitle_font = xlwt.Font()
    # 字体
    subtitle_font.name = "宋体"
    # 加粗
    subtitle_font.bold = True
    # 字号
    subtitle_font.height = 8 * 20
    # 颜色 可以从xlwt.XFStyle()中找，也可以直接对应索引
    subtitle_font.colour_index = 0x08
    # 应用样式
    subtitle_style.font = subtitle_font
    # 应用对齐方式
    subtitle_style.alignment = cell_align
    # 应用边框
    subtitle_style.borders = borders
    # 正文
    # 实例化字体样式
    text_font = xlwt.Font()
    # 字体
    text_font.name = "宋体"
    # 加粗
    text_font.bold = False
    # 字号
    text_font.height = 8 * 20
    # 颜色 可以从xlwt.XFStyle()中找，也可以直接对应索引
    text_font.colour_index = 0x08
    # 应用样式
    text_style.font = text_font
    # 应用对齐方式
    text_style.alignment = cell_align
    # 应用边框
    text_style.borders = borders


# 定义实体类
class StoreInfo:
    pass


if __name__ == "__main__":
    start_time = datetime.datetime.now()
    # 通过xlrd读取表格数据
    data = xlrd.open_workbook("美团外卖店铺信息采集（去重）.xlsx")
    # 操作工作簿
    # operateWorkbook(data)
    # 操作excel行
    # operateRows(data)
    # 操作excel列
    # operateColumns(data)
    # 操作excel单元格
    # operateCells(data)
    # 清洗数据逻辑
    target_list = clean_data(data)
    # 文件写入
    # 初始化标题样式
    title_style = xlwt.XFStyle()
    # 初始化副标题样式
    subtitle_style = xlwt.XFStyle()
    # 初始化正文样式
    text_style = xlwt.XFStyle()
    # 设置字体样式
    setTitleStyle()
    # 第一步：创建工作簿
    wb = xlwt.Workbook()
    # 第二步：创建工作表
    ws = wb.add_sheet("wx奶茶店铺数据集")
    # 第三步：填充数据
    write_excel(target_list)
    # 第四步：保存(分为入库+生成excel)
    # 定义一个列表用来接收数据
    data_list = []
    for i, item in enumerate(target_list):
        # 先实例化后用于接受
        obj = StoreInfoObj.StoreInfoObj(target_list[i][0],
                                        target_list[i][1],
                                        target_list[i][2],
                                        target_list[i][3],
                                        target_list[i][4],
                                        target_list[i][5])
        # obj = StoreInfo()
        # obj.store_name = target_list[i][0]
        # obj.store_url = target_list[i][1]
        # obj.store_address = target_list[i][2]
        # obj.store_consumption = target_list[i][3]
        # obj.store_score = target_list[i][4]
        # obj.comment_count = target_list[i][5]
        data_list.append(obj)
    print(data_list)
    # 链接到数据库
    db = dbhelper('127.0.0.1', 3306, "root", "123456", "test")
    # 插入语句
    sql = "insert into store_info(store_name,store_url,store_address,store_consumption,store_score,comment_count) VALUES (%s,%s,%s,%s,%s,%s)"
    val = []  # 空列表来存储元组
    for item in data_list:
        val.append((item.store_name, item.store_url, item.store_address, item.store_consumption, item.store_score,
                    item.comment_count))
    print(val)
    # db.executemanydata(sql, val)
    # wb.save("wx奶茶店铺数据202204_%s.xls" % (datetime.datetime.now().strftime("%Y%m%d%H%M%S")))
    end_time = datetime.datetime.now()
    print(end_time - start_time)
