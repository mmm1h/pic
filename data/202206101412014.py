import xlwings as xw
import os
from PIL import Image
import  re
import sys
import shutil
def execute_xw(file,paths):
    app = xw.App(visible=False,add_book=False)
    wb = app.books.open(file)
    i = 2
    for path in paths:
        excel_deal(path,wb,'B'+str(i))
        print(os.path.basename(str(path)))
        wb.sheets['sheet1'].range('A'+str(i)).value = os.path.splitext(os.path.basename(str(path)))[0]
        i += 1
    wb.save(file)
    wb.close()
    app.quit()

def add_center(sht, target, filePath, match=False, width=None, height=None, column_width=None, row_height=None):
    '''Excel智能居中插入图片

    优先级：match > width & height > column_width & row_height
    建议使用column_width或row_height，定义单元格最大宽或高

    :param sht: 工作表
    :param target: 目标单元格，字符串，如'A1'
    :param filePath: 图片绝对路径
    :param width: 图片宽度
    :param height: 图片高度
    :param column_width: 单元格最大宽度，默认100像素，0 <= column_width <= 1557.285
    :param row_height: 单元格最大高度，默认75像素，0 <= row_height <= 409.5
    :param match: 绝对匹配原图宽高，最大宽度1557.285，最大高度409.5
    '''
    unit_width = 6.107  # Excel默认列宽与像素的比
    rng = sht.range(target)  # 目标单元格
    name = os.path.basename(filePath)  # 文件名
    _width, _height = Image.open(filePath).size  # 原图片宽高
    NOT_SET = True  # 未设置单元格宽高
    # match
    if match:  # 绝对匹配图像
        width, height = _width, _height
    else:  # 不绝对匹配图像
        # width & height
        if width or height:
            if not height:  # 指定了宽，等比计算高
                height = width / _width * _height
            if not width:  # 指定了高，等比计算宽
                width = height / _height * _width
        else:
            # column_width & row_height
            if column_width and row_height:  # 同时指定单元格最大宽高
                width = row_height / _height * _width  # 根据单元格最大高度假设宽
                height = column_width / _width * _height  # 根据单元格最大宽度假设高
                area_width = column_width * height  # 假设宽优先的面积
                area_height = row_height * width  # 假设高优先的面积
                if area_width > area_height:
                    width = column_width
                else:
                    height = row_height
            elif not column_width and not row_height:  # 均无指定单元格最大宽高
                column_width = 333
                row_height = 250
                rng.column_width = column_width / unit_width  # 更新当前宽度
                rng.row_height = row_height  # 更新当前高度
                NOT_SET = False
                width = row_height / _height * _width  # 根据单元格最大高度假设宽
                height = column_width / _width * _height  # 根据单元格最大宽度假设高
                area_width = column_width * height  # 假设宽优先的面积
                area_height = row_height * width  # 假设高优先的面积
                if area_width > area_height:
                    height = row_height
                else:
                    width = column_width
            else:
                width = row_height  / _height * _width if row_height else column_width  # 仅设了单元格最大宽度
                height = column_width / _width * _height if column_width else row_height  # 仅设了单元格最大高度
    assert 0 <= width / unit_width <= 255
    assert 0 <= height <= 409.5

    if NOT_SET:
        rng.column_width = (width + 40) / unit_width  # 更新当前宽度
        rng.row_height = height + 20 # 更新当前高度
    left = rng.left + (rng.width - width) / 2  # 居中
    top = rng.top + (rng.height - height) / 2
    try:
        sht.pictures.add(filePath, left=left, top=top, width=width, height=height, scale=None, name=name)
    except Exception:  # 已有同名图片，采用默认命名
        pass

def excel_deal(file,wb,target):
    add_center(sht=wb.sheets['sheet1'],target=target,filePath = file,
               column_width=200,match=False)


def list_file(baseDir, paths=[],exts=[]):
    for entry in os.listdir(baseDir):
        if entry == '.svn' or entry == '.git':
            continue
        fullPath = baseDir + os.sep + entry
        if os.path.isdir(fullPath):
            list_file(fullPath, paths, exts)
        elif exts == None or os.path.splitext(entry)[1].lower() in exts:
            paths.append(fullPath)



if __name__ == '__main__':
    path = []
    cur_path = os.path.dirname(os.path.realpath(sys.argv[0]))
    p_path = os.path.abspath(os.path.join(cur_path, ".."))
    list_file(p_path,path,exts=[".png",".jpg",".jpeg",".bmp",".gif"])
    list.sort(path,key = lambda p : int("0"+re.search("([0-9]+)\.?",os.path.splitext(os.path.basename(p))[0]).group(1)),
                                        reverse=False)
    i=1
    count = 100
    while len(path)>0:
        f_name = cur_path + os.sep + "文件"+str(i)+".xlsx"
        try:
            shutil.copyfile(cur_path + os.sep +'模板.xlsx',f_name)
        except IOError as e:
            print("Unable to copy file. %s" % e)
        except:
            print("Unexpected error:", sys.exc_info())
        p = path[0:count]
        del path[0:count]
        execute_xw(f_name,p)
        print("文件"+str(i)+".xlsx" + " 生成完成 !")
        i += 1
    print("excel complete !!!")