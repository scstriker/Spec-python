from PIL import Image, ImageDraw

# 创建一个200x200的图像，使用RGBA模式（支持透明度）
size = 200
image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)

# 定义颜色
bg_color = (192, 192, 192, 255)  # Windows 98 灰色
border_light = (255, 255, 255, 255)  # 白色边框
border_dark = (64, 64, 64, 255)  # 深灰色边框

# 绘制主体（一个3D效果的圆形）
draw.ellipse([20, 20, size-20, size-20], fill=bg_color)

# 绘制3D效果的边框
for i in range(8):
    # 左上边框（亮色）
    draw.arc([20-i, 20-i, size-20+i, size-20+i], 135, 315, border_light)
    # 右下边框（暗色）
    draw.arc([20-i, 20-i, size-20+i, size-20+i], -45, 135, border_dark)

# 绘制抽奖标志（一个简单的星形）
star_points = [
    (size//2, 40),  # 顶点
    (size//2 + 30, size//2 - 20),  # 右上
    (size - 40, size//2),  # 右
    (size//2 + 30, size//2 + 20),  # 右下
    (size//2, size - 40),  # 底
    (size//2 - 30, size//2 + 20),  # 左下
    (40, size//2),  # 左
    (size//2 - 30, size//2 - 20),  # 左上
]
draw.polygon(star_points, fill=(0, 0, 128, 255))  # 使用深蓝色

# 保存为ICO文件
image.save('app_icon.ico', format='ICO', sizes=[(32, 32), (48, 48), (64, 64), (128, 128), (200, 200)]) 