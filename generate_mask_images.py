# generate_mask_images.py
"""
根据Excel文件中的每行内容生成图片（白底黑字）
命名规则：使用Excel中的内容作为文件名（去除特殊字符）
"""

import os
import re
import sys
from pathlib import Path
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("⚠️ 警告：未安装 pandas，请安装：pip install pandas openpyxl")

try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("⚠️ 警告：未安装 Pillow，请安装：pip install Pillow")


def get_base_dir():
    """获取程序基础目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def sanitize_filename(text: str) -> str:
    """
    清理文本，使其可以作为文件名
    保留中文字符、字母、数字、连字符和下划线
    """
    # 移除或替换不允许的文件名字符
    # Windows不允许: < > : " / \ | ? *
    text = re.sub(r'[<>:"/\\|?*]', '_', text)
    # 移除首尾空格和点
    text = text.strip(' .')
    # 限制文件名长度（Windows最大255字符，但为了安全限制为200）
    if len(text) > 200:
        text = text[:200]
    return text if text else "unnamed"


def create_text_image(text: str, output_path: Path, font_size: int = 48, padding: int = 20) -> bool:
    """
    创建白底黑字的文本图片（图片大小完全自适应文本内容）
    
    Args:
        text: 要显示的文本
        output_path: 输出图片路径
        font_size: 字体大小
        padding: 内边距
        
    Returns:
        是否成功
    """
    if not PIL_AVAILABLE:
        return False
    
    try:
        # 尝试加载字体（优先使用系统字体）
        font = None
        font_paths = [
            # Windows系统字体
            r"C:\Windows\Fonts\simhei.ttf",  # 黑体
            r"C:\Windows\Fonts\simsun.ttc",  # 宋体
            r"C:\Windows\Fonts\msyh.ttc",    # 微软雅黑
            r"C:\Windows\Fonts\msyhbd.ttc",  # 微软雅黑粗体
        ]
        
        for font_path in font_paths:
            if os.path.exists(font_path):
                try:
                    font = ImageFont.truetype(font_path, font_size)
                    break
                except Exception:
                    continue
        
        # 如果找不到字体，使用默认字体
        if font is None:
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
            except Exception:
                font = ImageFont.load_default()
        
        # 分割文本为多行
        text_lines = text.split('\n')
        
        # 使用一个足够大的临时图片来测量文本尺寸
        # 使用较大的尺寸确保能测量到所有文本
        temp_width = 5000
        temp_height = 5000
        temp_img = Image.new('RGB', (temp_width, temp_height), color='white')
        temp_draw = ImageDraw.Draw(temp_img)
        
        # 计算每行的宽度和高度
        max_line_width = 0
        total_height = 0
        line_heights = []
        line_widths = []
        
        for line in text_lines:
            if line.strip():
                try:
                    # 使用textbbox获取精确的文本边界框
                    bbox = temp_draw.textbbox((0, 0), line, font=font)
                    line_width = bbox[2] - bbox[0]
                    line_height = bbox[3] - bbox[1]
                except Exception:
                    # 如果textbbox不可用，使用textsize（旧版本Pillow）
                    try:
                        line_width, line_height = temp_draw.textsize(line, font=font)
                    except Exception:
                        # 如果都不可用，使用估算值
                        line_width = len(line) * font_size * 0.6
                        line_height = font_size
                
                max_line_width = max(max_line_width, line_width)
                line_heights.append(line_height)
                line_widths.append(line_width)
                total_height += line_height
            else:
                # 空行使用较小的行高
                empty_line_height = font_size * 0.3
                line_heights.append(empty_line_height)
                line_widths.append(0)
                total_height += empty_line_height
        
        # 如果没有任何文本，使用最小尺寸
        if max_line_width == 0:
            max_line_width = 100
        if total_height == 0:
            total_height = font_size
        
        # 根据文本内容计算最终图片尺寸（完全自适应）
        final_width = int(max_line_width + padding * 2)
        final_height = int(total_height + padding * 2)
        
        # 确保最小尺寸（至少能显示一个字符）
        min_width = font_size * 2
        min_height = font_size * 2
        final_width = max(final_width, min_width)
        final_height = max(final_height, min_height)
        
        # 创建最终图片
        img = Image.new('RGB', (final_width, final_height), color='white')
        draw = ImageDraw.Draw(img)
        
        # 绘制文本（居中显示）
        y = padding
        for i, line in enumerate(text_lines):
            if line.strip():
                # 获取当前行的宽度
                current_line_width = line_widths[i] if i < len(line_widths) else max_line_width
                
                # 水平居中
                x = (final_width - current_line_width) // 2
                
                # 绘制文本
                draw.text((x, y), line, fill='black', font=font)
                
                # 移动到下一行
                y += line_heights[i] if i < len(line_heights) else font_size
            else:
                # 空行，只增加高度
                y += line_heights[i] if i < len(line_heights) else font_size * 0.3
        
        # 保存图片
        img.save(output_path, 'PNG', quality=95)
        return True
        
    except Exception as e:
        print(f"  ❌ 生成图片失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def process_excel_file(excel_path: Path, output_dir: Path) -> bool:
    """
    处理Excel文件，为每行内容生成图片
    
    Args:
        excel_path: Excel文件路径
        output_dir: 输出目录路径
        
    Returns:
        是否成功
    """
    if not PANDAS_AVAILABLE:
        print("❌ 错误：未安装 pandas，无法读取Excel文件")
        print("   请安装：pip install pandas openpyxl")
        return False
    
    if not PIL_AVAILABLE:
        print("❌ 错误：未安装 Pillow，无法生成图片")
        print("   请安装：pip install Pillow")
        return False
    
    try:
        # 读取Excel文件
        print(f"正在读取Excel文件: {excel_path.name}")
        df = pd.read_excel(excel_path)
        
        if df.empty:
            print("⚠️ Excel文件为空")
            return False
        
        # 获取第一列（假设第一列包含要生成图片的文本）
        first_column = df.columns[0]
        print(f"使用列: {first_column}")
        print(f"共 {len(df)} 行数据")
        print()
        
        success_count = 0
        fail_count = 0
        
        # 遍历每一行
        for idx, row in df.iterrows():
            text = str(row.iloc[0]).strip()
            
            if not text or text == 'nan':
                print(f"  跳过第 {idx + 1} 行（内容为空）")
                continue
            
            # 生成文件名（使用文本内容，清理特殊字符，保持原始内容作为文件名）
            safe_filename = sanitize_filename(text)
            
            # 如果文件名太长，截断（但尽量保留重要部分）
            if len(safe_filename) > 200:
                # 保留前180个字符，加上序号
                safe_filename = safe_filename[:180] + f"_{idx + 1}"
            
            output_path = output_dir / f"{safe_filename}.png"
            
            # 如果文件已存在，添加序号避免覆盖
            counter = 1
            original_output_path = output_path
            while output_path.exists():
                # 在文件名末尾添加序号
                name_without_ext = original_output_path.stem
                output_path = output_dir / f"{name_without_ext}_{counter}.png"
                counter += 1
            
            print(f"  处理第 {idx + 1} 行: {text[:50]}{'...' if len(text) > 50 else ''}")
            
            # 生成图片
            if create_text_image(text, output_path):
                print(f"    ✓ 已保存: {output_path.name}")
                success_count += 1
            else:
                print(f"    ❌ 生成失败")
                fail_count += 1
        
        print()
        print(f"处理完成: 成功 {success_count} 个, 失败 {fail_count} 个")
        return success_count > 0
        
    except Exception as e:
        print(f"❌ 处理Excel文件失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主函数"""
    print("=" * 60)
    print("Excel文本生成图片工具（白底黑字）")
    print("=" * 60)
    
    # 获取基础目录
    base_dir = Path(get_base_dir())
    excel_path = base_dir / "zhezhao1.xlsx"
    output_dir = base_dir / "遮罩图片"
    
    # 检查Excel文件是否存在
    if not excel_path.exists():
        print(f"❌ 错误：Excel文件不存在: {excel_path}")
        print("   请确保在程序目录下存在'zhezhao1.xlsx'文件")
        input("按回车键退出...")
        return
    
    # 创建输出目录
    output_dir.mkdir(exist_ok=True)
    print(f"输出目录: {output_dir}")
    print()
    
    # 处理Excel文件
    success = process_excel_file(excel_path, output_dir)
    
    if success:
        print()
        print("=" * 60)
        print("✅ 图片已保存到'遮罩图片'文件夹")
        print("=" * 60)
    else:
        print()
        print("=" * 60)
        print("❌ 处理失败")
        print("=" * 60)
    
    input("按回车键退出...")


if __name__ == "__main__":
    main()
