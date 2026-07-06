import os
from PIL import Image

def compress_png(input_dir, output_dir, colors=256):
    """
    批量压缩 PNG 图片
    :param input_dir: 原始图片所在目录
    :param output_dir: 压缩后图片保存目录
    :param colors: 颜色保留数量（越小压缩率越高，一般 256 能保持较好画质同时大幅减小体积）
    """
    # 如果输出文件夹不存在，则创建
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 遍历输入目录中的所有文件
    for filename in os.listdir(input_dir):
        if filename.lower().endswith('.png'):
            input_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, filename)

            try:
                print(f"正在处理: {filename}")
                # 打开图片
                with Image.open(input_path) as img:
                    # 如果有透明通道 (RGBA)，在转换时需要保留
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')
                    
                    # 修改图片尺寸为 256x256 (32x32 实在太小了，会导致严重马赛克/模块化)
                    # 256x256 既能保证清晰度，又能让大小下降到 几十 KB 的目标
                    img = img.resize((256, 256), Image.LANCZOS)
                    
                    # 移除 'P' 模式强制转换，保留原本的色彩平滑度，避免透明边缘出现锯齿
                    compressed_img = img

                    # 保存图片，开启 optimize 优化
                    compressed_img.save(output_path, format='PNG', optimize=True)
                    
                    # 计算并打印压缩效果
                    original_size = os.path.getsize(input_path) / 1024
                    new_size = os.path.getsize(output_path) / 1024
                    print(f"成功压缩: {filename} ({original_size:.2f} KB -> {new_size:.2f} KB)")
            except Exception as e:
                print(f"压缩 {filename} 时出错: {e}")

if __name__ == "__main__":
    # 获取当前脚本所在的绝对路径目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 设置输入文件夹为当前脚本所在目录
    INPUT_FOLDER = script_dir
    # 设置输出文件夹为当前脚本同级别下的 'output' 文件夹
    OUTPUT_FOLDER = os.path.join(script_dir, 'output')

    print(f"脚本所在目录: {INPUT_FOLDER}")
    print(f"压缩后输出目录: {OUTPUT_FOLDER}")
    print("开始压缩...")
    compress_png(INPUT_FOLDER, OUTPUT_FOLDER, colors=256)
    print("压缩完成！")
