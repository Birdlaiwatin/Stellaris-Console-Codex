from A配置路径 import *
import os
import imageio.v3 as imageio
import xlsxwriter


def convert_dds_to_png():
    # 确保输出目录存在
    os.makedirs(ROOM_IMAGES_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 查找所有以"_room"为后缀的.dds文件
    room_files = []
    for filename in os.listdir(ROOM_DIR):
        if filename.endswith('_room.dds'):
            room_files.append(filename)

    # 准备表格数据
    table_data = []

    for dds_file in room_files:
        # 构建完整文件路径
        dds_path = os.path.join(ROOM_DIR, dds_file)

        # 生成输出png文件名（去掉.dds后缀）
        png_filename = os.path.splitext(dds_file)[0] + '.png'
        png_path = os.path.join(ROOM_IMAGES_DIR, png_filename)

        try:
            # 读取DDS文件并转换为PNG
            image = imageio.imread(dds_path)
            imageio.imwrite(png_path, image)

            # 准备表格行数据（图片名称不带后缀）
            image_name = os.path.splitext(dds_file)[0]
            table_data.append({
                'image_path': png_path,
                'image_name': image_name,
                'png_filename': png_filename
            })

            print(f"成功转换: {dds_file} -> {png_filename}")

        except Exception as e:
            print(f"转换失败 {dds_file}: {str(e)}")

    # 生成Excel表格
    if table_data:
        excel_path = os.path.join(OUTPUT_DIR, f'{VERSION}房间提取.xlsx')

        workbook = xlsxwriter.Workbook(excel_path)
        worksheet = workbook.add_worksheet('房间图片')

        # 设置表头
        headers = ['图片', '图片名称']
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})

        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 根据图片尺寸设置行高和列宽
        # 图片大约1000×350像素，在Excel中设置合适的行高列宽
        worksheet.set_column('A:A', 30)  # 图片列设置较宽
        worksheet.set_column('B:B', 20)  # 图片名称列

        # 写入数据并插入图片
        for row, data in enumerate(table_data, 1):
            # 在A列插入图片
            worksheet.insert_image(
                f'A{row + 1}',  # Excel位置（从A2开始）
                data['image_path'],
                {
                    'x_scale': 0.2,  # 缩放比例，根据需要调整
                    'y_scale': 0.2,  # 缩放比例，根据需要调整
                    'object_position': 1  # 移动和调整单元格大小
                }
            )

            # 在B列写入图片名称
            worksheet.write(row, 1, data['image_name'])

            # 设置行高以适应图片
            worksheet.set_row(row, 80)  # 设置行高

        workbook.close()
        print(f"表格已生成: {excel_path}")
        print(f"共处理 {len(table_data)} 个文件")
    else:
        print("未找到可转换的_room.dds文件")


if __name__ == "__main__":
    convert_dds_to_png()