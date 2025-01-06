import xlwings as xw
import os


def cm_to_points(cm):
    """将厘米转换为 Excel 点（point）单位"""
    return cm * 28.35


def insert_image_to_excel(image_path, picture_cell, left_offset, top_offset, width_cm, height_cm, sheet):
    """将图片插入到 Excel 指定位置，并调整图片尺寸"""
    width_points = cm_to_points(width_cm)
    height_points = cm_to_points(height_cm)

    # 获取目标单元格的左上角位置
    cell = sheet.range(picture_cell)

    # 在目标位置插入图片并调整大小
    sheet.pictures.add(image_path,
                       left=cell.left + left_offset,
                       top=cell.top + top_offset,
                       width=width_points,
                       height=height_points)
    print(f"图片 {image_path} 已经插入到工作表 {sheet.name} 的 {picture_cell} 单元格")


def insert_images_from_folder(image_folder, excel_path, sheet_name, image_cell_mapping, left_offset=10, top_offset=10,
                              width_cm=11.21, height_cm=8.77):
    """从文件夹批量插入图片到 Excel"""
    # 尝试连接到已打开的 Excel 实例
    app = None
    for excel_app in xw.apps:
        # 如果文件已经打开，找到该文件并连接
        for book in excel_app.books:
            if book.fullname == excel_path:
                app = excel_app
                workbook = book
                break
        if app:
            break

    # 如果工作簿没有打开，打开它
    if not app:
        app = xw.App(visible=True, add_book=False)  # 修改这里，避免启动新空白工作簿
        workbook = app.books.open(excel_path)

    # 获取目标工作表
    sheet = workbook.sheets[sheet_name]

    # 获取文件夹中的所有图片文件
    image_files = [f for f in os.listdir(image_folder) if f.lower().endswith(('png', 'jpg', 'jpeg', 'gif', 'bmp'))]

    # 遍历图片文件并根据映射插入图片
    for image_file in image_files:
        image_name = os.path.splitext(image_file)[0]

        if image_name in image_cell_mapping:
            picture_cell = image_cell_mapping[image_name]
            image_path = os.path.join(image_folder, image_file)

            # 调用插入图片的函数
            insert_image_to_excel(image_path, picture_cell, left_offset, top_offset, width_cm, height_cm, sheet)

    # 保存 Excel 文件
    workbook.save()
    print(f"所有图片已插入并保存到 {excel_path} 中的 {sheet_name} 工作表")


# 示例调用函数
image_folder = r'C:\Users\Seir\Desktop\DC-SCM\POL Test Pictures\kunlunshan\P3V3_NVME2'  # 图片文件夹路径
excel_path = r'C:\Users\Seir\Desktop\DC-SCM\8NVME_board_YZBB-02874-101_DCDC_20241230.xlsx'  # Excel 文件路径
sheet_name = 'P3V3_NVME2'  # 插入图片的工作表名称

# 图片文件名到目标单元格的映射
image_cell_mapping = {
    'T1-1': 'F25',
    'T2-1': 'F47',
    'T3-1': 'F69',
    'T4-1': 'F91',
    'T5-1': 'F113',
    'T5-2': 'R113',
    'T6-1': 'F135',
    'T6-2': 'R135',
    # 更多映射
}

# 插入图片的偏移量和尺寸
left_offset = 40  # 左偏移量
top_offset = 5  # 上偏移量
width_cm = 11.21  # 图片宽度（单位：厘米）
height_cm = 8.77  # 图片高度（单位：厘米）

# 批量插入图片
insert_images_from_folder(image_folder, excel_path, sheet_name, image_cell_mapping, left_offset, top_offset, width_cm,
                          height_cm)
