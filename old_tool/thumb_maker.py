#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
专业头像裁剪工具 - 微信风格
功能：固定比例裁剪框 + 拖动缩放 + 自动压缩
"""

import os
import sys
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkfont

class AvatarCropper:
    def __init__(self, root):
        self.root = root
        self.root.title("🐾 专业头像裁剪工具 - 微信风格")
        self.root.geometry("1000x750")
        self.root.configure(bg='#f5f5f5')
        
        # 变量
        self.image_path = None
        self.original_image = None
        self.displayed_image = None   # 缩放后用于显示的图片
        self.tk_image = None
        self.crop_box = None          # (x1, y1, x2, y2) 在原图上的坐标
        self.display_crop_box = None  # 在显示图片上的坐标
        self.dragging = False
        self.resizing = False
        self.resize_corner = None
        self.offset_x = 0
        self.offset_y = 0
        self.crop_w = 0               # 移动时保存的裁剪框原始宽度（原图像素）
        self.crop_h = 0               # 移动时保存的裁剪框原始高度（原图像素）
        
        # 裁剪参数
        self.crop_ratio = 1.0         # 1:1 正方形
        self.output_size = 300        # 压缩后的尺寸
        self.min_crop_size = 100      # 最小裁剪尺寸（显示图上）
        
        # 创建界面
        self.create_widgets()
        
        # 绑定窗口大小变化事件
        self.root.bind('<Configure>', self.on_resize)
    
    def create_widgets(self):
        """创建界面组件"""
        # ========== 顶部工具栏 ==========
        toolbar = tk.Frame(self.root, bg='#2196F3', height=60)
        toolbar.pack(fill=tk.X)
        toolbar.pack_propagate(False)
        
        title_font = tkfont.Font(family='Arial', size=14, weight='bold')
        title_label = tk.Label(toolbar, text="🐱 专业头像裁剪工具", 
                              font=title_font, bg='#2196F3', fg='white')
        title_label.pack(side=tk.LEFT, padx=20, pady=15)
        
        # 按钮组
        btn_frame = tk.Frame(toolbar, bg='#2196F3')
        btn_frame.pack(side=tk.RIGHT, padx=20)
        
        select_btn = tk.Button(btn_frame, text="📁 选择图片", 
                              command=self.select_image,
                              font=('Arial', 10, 'bold'), bg='#4CAF50', fg='white',
                              padx=20, pady=8, relief=tk.FLAT, cursor='hand2')
        select_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = tk.Button(btn_frame, text="💾 裁剪并保存", 
                                 command=self.crop_and_save,
                                 font=('Arial', 10, 'bold'), bg='#FF9800', fg='white',
                                 padx=20, pady=8, relief=tk.FLAT, 
                                 cursor='hand2', state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # ========== 状态栏 ==========
        status_frame = tk.Frame(self.root, bg='#e0e0e0', height=40)
        status_frame.pack(fill=tk.X)
        status_frame.pack_propagate(False)
        
        self.status_label = tk.Label(status_frame, text="💡 提示：拖动裁剪框移动位置，拖动边缘调整大小", 
                                    font=('Arial', 9), bg='#e0e0e0', fg='#666')
        self.status_label.pack(side=tk.LEFT, padx=20)
        
        self.info_label = tk.Label(status_frame, text="", 
                                  font=('Arial', 9, 'bold'), bg='#e0e0e0', fg='#2196F3')
        self.info_label.pack(side=tk.RIGHT, padx=20)
        
        # ========== 主工作区 ==========
        main_frame = tk.Frame(self.root, bg='#f5f5f5')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        # 左侧：图片显示区
        left_frame = tk.Frame(main_frame, bg='white', relief=tk.SUNKEN, borderwidth=2)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(left_frame, bg='#f0f0f0', cursor='fleur')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # 绑定鼠标事件
        self.canvas.bind('<ButtonPress-1>', self.on_mouse_down)
        self.canvas.bind('<B1-Motion>', self.on_mouse_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_up)
        self.canvas.bind('<Motion>', self.on_mouse_move)
        
        # 右侧：参数设置区
        right_frame = tk.LabelFrame(main_frame, text="⚙️ 裁剪参数", 
                                   font=('Arial', 10, 'bold'), bg='#f5f5f5',
                                   padx=15, pady=15, width=280)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))
        right_frame.pack_propagate(False)
        
        # 输出尺寸设置
        size_label = tk.Label(right_frame, text="输出尺寸:", 
                             font=('Arial', 9, 'bold'), bg='#f5f5f5', anchor='w')
        size_label.pack(fill=tk.X, pady=(0, 5))
        
        size_frame = tk.Frame(right_frame, bg='#f5f5f5')
        size_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.size_var = tk.IntVar(value=300)
        size_entry = tk.Entry(size_frame, textvariable=self.size_var, 
                             font=('Arial', 9), width=8)
        size_entry.pack(side=tk.LEFT)
        
        size_unit = tk.Label(size_frame, text="× 300 像素", 
                           font=('Arial', 9), bg='#f5f5f5')
        size_unit.pack(side=tk.LEFT, padx=5)
        
        # 快捷尺寸按钮
        preset_frame = tk.Frame(right_frame, bg='#f5f5f5')
        preset_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(preset_frame, text="常用尺寸:", 
                font=('Arial', 9), bg='#f5f5f5').pack(anchor='w', pady=(0, 5))
        
        presets = [(150, "小 (150×150)"), (300, "中 (300×300)"), (500, "大 (500×500)")]
        for size, text in presets:
            btn = tk.Button(preset_frame, text=text,
                          command=lambda s=size: self.set_preset_size(s),
                          font=('Arial', 8), bg='#e3f2fd', relief=tk.RAISED,
                          padx=10, pady=3, cursor='hand2')
            btn.pack(fill=tk.X, pady=2)
        
        # 裁剪框控制按钮
        control_label = tk.Label(right_frame, text="裁剪框控制:", 
                               font=('Arial', 9, 'bold'), bg='#f5f5f5', anchor='w')
        control_label.pack(fill=tk.X, pady=(15, 5))
        
        control_frame = tk.Frame(right_frame, bg='#f5f5f5')
        control_frame.pack(fill=tk.X, pady=(0, 15))
        
        reset_btn = tk.Button(control_frame, text="↺ 重置裁剪框", 
                            command=self.reset_crop_box,
                            font=('Arial', 9), bg='#ffeb3b', relief=tk.RAISED,
                            padx=10, pady=5, cursor='hand2')
        reset_btn.pack(fill=tk.X)
        
        center_btn = tk.Button(control_frame, text="🎯 居中裁剪框", 
                             command=self.center_crop_box,
                             font=('Arial', 9), bg='#2196F3', fg='white', relief=tk.RAISED,
                             padx=10, pady=5, cursor='hand2')
        center_btn.pack(fill=tk.X, pady=(5, 0))
        
        # 使用说明
        help_label = tk.Label(right_frame, text="📋 使用说明:", 
                            font=('Arial', 9, 'bold'), bg='#f5f5f5', anchor='w')
        help_label.pack(fill=tk.X, pady=(15, 5))
        
        help_text = tk.Text(right_frame, height=8, font=('Arial', 8), 
                          bg='#e8f5e9', relief=tk.SOLID, borderwidth=1)
        help_text.pack(fill=tk.X)
        help_text.insert(tk.END, "1. 点击「选择图片」\n")
        help_text.insert(tk.END, "2. 拖动蓝色框移动位置\n")
        help_text.insert(tk.END, "3. 拖动边缘调整框大小\n")
        help_text.insert(tk.END, "4. 点击「裁剪并保存」\n")
        help_text.insert(tk.END, "5. 自动保存为「原图名_thumb」\n")
        help_text.config(state=tk.DISABLED)
        
        # 底部版权
        footer = tk.Label(right_frame, text="© 2026 猫咪档案馆", 
                         font=('Arial', 7), bg='#f5f5f5', fg='#999')
        footer.pack(side=tk.BOTTOM, pady=(10, 0))
    
    def select_image(self):
        """选择图片文件"""
        filetypes = [
            ('图片文件', '*.jpg *.jpeg *.png *.bmp *.gif'),
            ('所有文件', '*.*')
        ]
        
        filename = filedialog.askopenfilename(
            title='选择猫咪图片',
            filetypes=filetypes,
            initialdir=os.path.expanduser('~')
        )
        
        if filename:
            try:
                self.image_path = filename
                self.original_image = Image.open(filename)
                
                file_name = os.path.basename(filename)
                self.status_label.config(text=f"✅ 已加载: {file_name}")
                
                self.init_crop_box()
                self.display_image()
                self.save_btn.config(state=tk.NORMAL)
                self.update_info_label()
                
            except Exception as e:
                messagebox.showerror("错误", f"无法打开图片:\n{str(e)}")
    
    def init_crop_box(self):
        """初始化裁剪框（居中，适中大小）"""
        if not self.original_image:
            return
        
        img_w, img_h = self.original_image.size
        crop_size = int(min(img_w, img_h) * 0.6)
        crop_size = max(crop_size, 100)
        
        x1 = (img_w - crop_size) // 2
        y1 = (img_h - crop_size) // 2
        x2 = x1 + crop_size
        y2 = y1 + crop_size
        
        self.crop_box = (x1, y1, x2, y2)
        self.display_crop_box = None
    
    def display_image(self):
        """在画布上显示图片并绘制裁剪框"""
        if not self.original_image:
            return
        
        canvas_w = self.canvas.winfo_width() or 800
        canvas_h = self.canvas.winfo_height() or 600
        
        img_w, img_h = self.original_image.size
        scale_w = canvas_w / img_w
        scale_h = canvas_h / img_h
        self.display_scale = min(scale_w, scale_h, 1.0)
        
        display_w = int(img_w * self.display_scale)
        display_h = int(img_h * self.display_scale)
        
        # 缩放后的图片
        self.displayed_image = self.original_image.resize(
            (display_w, display_h), 
            Image.LANCZOS
        )
        
        self.tk_image = ImageTk.PhotoImage(self.displayed_image)
        self.canvas.delete('all')
        
        self.img_x = (canvas_w - display_w) // 2
        self.img_y = (canvas_h - display_h) // 2
        
        self.canvas.create_image(
            self.img_x, self.img_y,
            anchor=tk.NW,
            image=self.tk_image,
            tags='image'
        )
        
        self.draw_crop_box()
    
    def draw_crop_box(self):
        """绘制裁剪框"""
        if not self.crop_box or not self.displayed_image:
            return
        
        x1, y1, x2, y2 = self.crop_box
        dx1 = self.img_x + int(x1 * self.display_scale)
        dy1 = self.img_y + int(y1 * self.display_scale)
        dx2 = self.img_x + int(x2 * self.display_scale)
        dy2 = self.img_y + int(y2 * self.display_scale)
        
        self.display_crop_box = (dx1, dy1, dx2, dy2)
        
        self.canvas.delete('crop_box')
        self.canvas.delete('crop_overlay')
        self.canvas.delete('crop_handle')
        
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        
        overlay_color = '#000000'
        # 四周半透明遮罩
        self.canvas.create_rectangle(0, 0, dx1, canvas_h,
                                     fill=overlay_color, stipple='gray50', tags='crop_overlay')
        self.canvas.create_rectangle(dx2, 0, canvas_w, canvas_h,
                                     fill=overlay_color, stipple='gray50', tags='crop_overlay')
        self.canvas.create_rectangle(dx1, 0, dx2, dy1,
                                     fill=overlay_color, stipple='gray50', tags='crop_overlay')
        self.canvas.create_rectangle(dx1, dy2, dx2, canvas_h,
                                     fill=overlay_color, stipple='gray50', tags='crop_overlay')
        
        # 裁剪框边框
        self.canvas.create_rectangle(dx1, dy1, dx2, dy2,
                                     outline='#2196F3', width=3, tags='crop_box')
        
        # 四个角标
        handle_size = 8
        handles = [(dx1, dy1), (dx2, dy1), (dx1, dy2), (dx2, dy2)]
        for hx, hy in handles:
            self.canvas.create_oval(
                hx - handle_size, hy - handle_size,
                hx + handle_size, hy + handle_size,
                fill='#FF5722', outline='white', width=2,
                tags='crop_handle'
            )
    
    def on_mouse_down(self, event):
        """鼠标按下"""
        if not self.display_crop_box:
            return
        
        dx1, dy1, dx2, dy2 = self.display_crop_box
        handle_size = 12
        
        # 检测是否按在角标上（调整大小）
        corners = {
            'nw': (dx1, dy1),
            'ne': (dx2, dy1),
            'sw': (dx1, dy2),
            'se': (dx2, dy2)
        }
        for corner, (cx, cy) in corners.items():
            if abs(event.x - cx) < handle_size and abs(event.y - cy) < handle_size:
                self.resizing = True
                self.resize_corner = corner
                self.offset_x = event.x
                self.offset_y = event.y
                self.canvas.config(cursor='fleur')
                return
        
        # 检测是否在框内（移动）
        if dx1 <= event.x <= dx2 and dy1 <= event.y <= dy2:
            self.dragging = True
            self.offset_x = event.x - dx1
            self.offset_y = event.y - dy1
            # 记录移动前的原图宽高，确保移动时尺寸不变
            self.crop_w = self.crop_box[2] - self.crop_box[0]
            self.crop_h = self.crop_box[3] - self.crop_box[1]
            self.canvas.config(cursor='fleur')
    
    def on_mouse_drag(self, event):
        """鼠标拖拽"""
        if self.dragging:
            self.move_crop_box(event)
        elif self.resizing:
            self.resize_crop_box(event)
    
    def on_mouse_up(self, event):
        """鼠标释放"""
        self.dragging = False
        self.resizing = False
        self.resize_corner = None
        self.canvas.config(cursor='arrow')
        self.update_info_label()
    
    def on_mouse_move(self, event):
        """鼠标移动（改变光标形状）"""
        if not self.display_crop_box:
            return
        
        dx1, dy1, dx2, dy2 = self.display_crop_box
        handle_size = 12
        
        corners = [(dx1, dy1), (dx2, dy1), (dx1, dy2), (dx2, dy2)]
        for cx, cy in corners:
            if abs(event.x - cx) < handle_size and abs(event.y - cy) < handle_size:
                self.canvas.config(cursor='fleur')
                return
        
        if dx1 <= event.x <= dx2 and dy1 <= event.y <= dy2:
            self.canvas.config(cursor='fleur')
        else:
            self.canvas.config(cursor='arrow')
    
    def move_crop_box(self, event):
        """移动裁剪框（保持尺寸完全不变，仅限制位置）"""
        if not self.crop_box or not self.displayed_image:
            return

        crop_w = self.crop_w
        crop_h = self.crop_h

        new_dx1 = event.x - self.offset_x
        new_dy1 = event.y - self.offset_y

        x1_float = (new_dx1 - self.img_x) / self.display_scale
        y1_float = (new_dy1 - self.img_y) / self.display_scale

        img_w, img_h = self.original_image.size

        x1 = max(0, min(int(x1_float), img_w - crop_w))
        y1 = max(0, min(int(y1_float), img_h - crop_h))

        x2 = x1 + crop_w
        y2 = y1 + crop_h

        self.crop_box = (x1, y1, x2, y2)
        self.draw_crop_box()
    
    def resize_crop_box(self, event):
        """调整裁剪框大小（拖动角点，对点固定，保持正方形）"""
        if not self.crop_box or not self.resize_corner:
            return

        x1, y1, x2, y2 = self.crop_box
        img_w, img_h = self.original_image.size

        # 鼠标移动增量（原图坐标）
        delta_x = (event.x - self.offset_x) / self.display_scale
        delta_y = (event.y - self.offset_y) / self.display_scale

        # 更新拖动的角点坐标
        if self.resize_corner == 'nw':
            x1 += delta_x
            y1 += delta_y
        elif self.resize_corner == 'ne':
            x2 += delta_x
            y1 += delta_y
        elif self.resize_corner == 'sw':
            x1 += delta_x
            y2 += delta_y
        elif self.resize_corner == 'se':
            x2 += delta_x
            y2 += delta_y

        # 边界钳制（不能超出图片）
        x1 = max(0, x1)
        y1 = max(0, y1)
        x2 = min(img_w, x2)
        y2 = min(img_h, y2)

        # 确定固定点（对角的那个点）
        if self.resize_corner == 'se':
            fixed_x, fixed_y = x1, y1
        elif self.resize_corner == 'sw':
            fixed_x, fixed_y = x2, y1
        elif self.resize_corner == 'ne':
            fixed_x, fixed_y = x1, y2
        elif self.resize_corner == 'nw':
            fixed_x, fixed_y = x2, y2

        # 计算当前宽高，并取较小值作为正方形边长（同时确保不小于最小尺寸）
        w = x2 - x1
        h = y2 - y1
        side = max(min(w, h), self.min_crop_size / self.display_scale)

        # 根据固定点和角点方向重新计算正方形坐标
        if self.resize_corner == 'se':
            # 固定左上角
            x1, y1 = fixed_x, fixed_y
            x2 = x1 + side
            y2 = y1 + side
        elif self.resize_corner == 'sw':
            # 固定右上角
            x2, y1 = fixed_x, fixed_y
            x1 = x2 - side
            y2 = y1 + side
        elif self.resize_corner == 'ne':
            # 固定左下角
            x1, y2 = fixed_x, fixed_y
            x2 = x1 + side
            y1 = y2 - side
        elif self.resize_corner == 'nw':
            # 固定右下角
            x2, y2 = fixed_x, fixed_y
            x1 = x2 - side
            y1 = y2 - side

        # 最终边界保护（避免计算误差越界）
        x1 = max(0, x1)
        y1 = max(0, y1)
        x2 = min(img_w, x2)
        y2 = min(img_h, y2)

        self.crop_box = (int(x1), int(y1), int(x2), int(y2))
        self.draw_crop_box()

        # 更新鼠标偏移，防止跳跃
        self.offset_x = event.x
        self.offset_y = event.y
    
    def reset_crop_box(self):
        """重置裁剪框"""
        if not self.original_image:
            return
        self.init_crop_box()
        self.display_image()
        self.update_info_label()
        messagebox.showinfo("提示", "裁剪框已重置为默认大小和位置！")
    
    def center_crop_box(self):
        """居中裁剪框"""
        if not self.crop_box or not self.original_image:
            return
        x1, y1, x2, y2 = self.crop_box
        size = x2 - x1
        img_w, img_h = self.original_image.size
        new_x1 = (img_w - size) // 2
        new_y1 = (img_h - size) // 2
        self.crop_box = (new_x1, new_y1, new_x1 + size, new_y1 + size)
        self.display_image()
        self.update_info_label()
        messagebox.showinfo("提示", "裁剪框已居中！")
    
    def set_preset_size(self, size):
        self.size_var.set(size)
        messagebox.showinfo("提示", f"输出尺寸已设置为 {size}×{size} 像素")
    
    def update_info_label(self):
        if not self.crop_box:
            return
        x1, y1, x2, y2 = self.crop_box
        crop_size = x2 - x1
        img_w, img_h = self.original_image.size
        self.info_label.config(text=f"裁剪: {crop_size}×{crop_size} | 原图: {img_w}×{img_h}")
    
    def crop_and_save(self):
        """裁剪并保存"""
        if not self.crop_box or not self.original_image:
            messagebox.showwarning("警告", "请先选择图片并调整裁剪框！")
            return
        
        try:
            cropped = self.original_image.crop(self.crop_box)
            output_size = self.size_var.get() or 300
            
            thumbnail = cropped.resize((output_size, output_size), Image.LANCZOS)
            
            dir_name = os.path.dirname(self.image_path)
            base_name = os.path.basename(self.image_path)
            name, ext = os.path.splitext(base_name)
            new_filename = f"{name}_thumb{ext}"
            new_filepath = os.path.join(dir_name, new_filename)
            
            save_kwargs = {}
            if ext.lower() in ['.jpg', '.jpeg']:
                save_kwargs['quality'] = 90
                save_kwargs['optimize'] = True
            elif ext.lower() == '.png':
                save_kwargs['optimize'] = True
            thumbnail.save(new_filepath, **save_kwargs)
            
            original_size = os.path.getsize(self.image_path) / 1024
            new_size = os.path.getsize(new_filepath) / 1024
            ratio = (1 - new_size / original_size) * 100 if original_size > 0 else 0
            
            messagebox.showinfo("✅ 裁剪成功",
                                f"📁 文件: {new_filename}\n\n"
                                f"📐 尺寸: {output_size}×{output_size} 像素\n"
                                f"💾 大小: {new_size:.1f} KB\n"
                                f"⚡ 压缩率: {ratio:.1f}%\n\n"
                                f"📍 保存路径:\n{new_filepath}")
            self.status_label.config(text="💡 裁剪完成！可以继续处理下一张图片")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败:\n{str(e)}")
    
    def on_resize(self, event):
        if self.original_image:
            self.display_image()


def main():
    try:
        from PIL import Image, ImageTk
    except ImportError:
        print("❌ 错误：未安装Pillow库")
        print("📦 请运行：pip install Pillow")
        sys.exit(1)
    
    root = tk.Tk()
    app = AvatarCropper(root)
    root.mainloop()


if __name__ == '__main__':
    main()