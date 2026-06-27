import json
import tkinter as tk
import os
import sys

def parse_color(hex_str):
    """将 #RRGGBB 转换为 tkinter 可用的颜色值（直接保留）"""
    hex_str = hex_str.strip()
    if not hex_str.startswith('#'):
        hex_str = '#' + hex_str
    return hex_str

def create_preview_window(json_path='tag_color.json'):
    # 读取颜色配置文件
    if not os.path.exists(json_path):
        print(f"❌ 找不到文件: {json_path}")
        sys.exit(1)

    with open(json_path, 'r', encoding='utf-8') as f:
        try:
            tag_colors = json.load(f)
        except json.JSONDecodeError as e:
            print(f"❌ JSON 格式错误: {e}")
            sys.exit(1)

    if not isinstance(tag_colors, dict):
        print("❌ JSON 内容应该是一个对象（键值对）。")
        sys.exit(1)

    # 创建主窗口
    root = tk.Tk()
    root.title("标签颜色预览")
    root.configure(bg='#F7F5F0')

    # 标题
    title_frame = tk.Frame(root, bg='#F7F5F0')
    title_frame.pack(pady=20)
    tk.Label(title_frame, text="标签样式预览", font=('微软雅黑', 16, 'bold'),
             bg='#F7F5F0', fg='#4A4A4A').pack()

    # 主体滚动框架
    canvas = tk.Canvas(root, bg='#F7F5F0', highlightthickness=0)
    scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg='#F7F5F0')
    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True, padx=20)
    scrollbar.pack(side="right", fill="y")

    # 为每个标签创建一行（两个按钮：普通和选中）
    for tag, color in tag_colors.items():
        hex_color = parse_color(color)
        frame = tk.Frame(scroll_frame, bg='#F7F5F0')
        frame.pack(pady=10)

        # 左侧标签名
        tk.Label(frame, text=tag, font=('微软雅黑', 11),
                 bg='#F7F5F0', fg='#555').pack(side=tk.LEFT, padx=(0, 15))

        # 普通按钮
        btn1 = tk.Button(frame, text=f"  {tag}  ", font=('微软雅黑', 10, 'bold'),
                         bg=hex_color, fg='white', relief='flat', bd=0,
                         activebackground=hex_color, activeforeground='white',
                         padx=20, pady=8, borderwidth=0, highlightthickness=0,
                         cursor='hand2')
        btn1.pack(side=tk.LEFT, padx=5)

        # 选中按钮（带有勾）
        btn2 = tk.Button(frame, text=f"  ✓ {tag}  ", font=('微软雅黑', 10, 'bold'),
                         bg=hex_color, fg='white', relief='flat', bd=0,
                         activebackground=hex_color, activeforeground='white',
                         padx=20, pady=8, borderwidth=0, highlightthickness=0,
                         cursor='hand2')
        btn2.pack(side=tk.LEFT, padx=5)

    root.geometry("600x500")
    root.mainloop()

if __name__ == '__main__':
    create_preview_window()