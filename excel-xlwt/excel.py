# 写入Excel文件的扩展工具
import xlwt

# 导入random包
import random

# 创建workboos对象
book = xlwt.Workbook(encoding="utf-8", style_compression=0)

# 创建工作表
sheet = book.add_sheet('test', cell_overwrite_ok=True)


# 写入数据
row = 0 # 行
column = 0 # 列
for i in range(72):
    # 创建一个样式对象，初始化样式 style
    style = xlwt.XFStyle()

    # 常用颜色：0-黑色 1-白色 2-红色 3-绿色 4-蓝色 5-黄色 6-粉色 7-青色 ...
    # 背景色
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    num = random.randint(10, 60) # 生成一个随机颜色
    pattern.pattern_fore_colour = num  # 设置背景颜色
    style.pattern = pattern

    # 边框
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN # DASHED-虚线 NO_LINE-没有 THIN-实线
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0
    borders.right_colour = 0
    borders.top_colour = 0
    borders.bottom_colour = 0
    style.borders = borders

    # 字体
    font = xlwt.Font()
    font.name = 'Calibri' # 字体
    font.colour_index = 1 # 字体颜色
    font.height = 400 # 字体大小
    font.bold = True # 字体是否为粗体
    style.font = font

    # 对齐方式
    alignment = xlwt.Alignment()
    # alignment.horz = xlwt.Alignment.HORZ_CENTER # 水平对齐方式
    alignment.horz = 0x02 # 水平对齐：0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.vert = 0x01 # 垂直对齐：0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.wrap = 1 # 设置自动换行
    style.alignment = alignment

    # 写入数据，第row行，第column列，具体内容是i
    sheet.write(row, column, i, style)

    column += 1
    if column > 8:
        column = 0
        row += 1
    else:
        # 设置列宽
        sheet.col(column).width = 256 * (column + 10)


# 设置冻结窗口
sheet.set_panes_frozen('1') # 设置冻结为真
sheet.set_horz_split_pos(1) # 水平冻结
# work_sheet.set_vert_split_pos(1) # 垂直冻结

# 设置行高
line_height = xlwt.easyxf('font:height 720')
sheet.row(0).set_style(line_height)  # 给第1行设置行高
sheet.row(1).set_style(line_height)  # 给第2行设置行高
sheet.row(2).set_style(line_height)  # 给第3行设置行高


# 保存
book.save('test.xls')