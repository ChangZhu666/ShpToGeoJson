import os
import sys
import json
import geopandas as gpd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
from tkinterdnd2 import DND_FILES, TkinterDnD
from pathlib import Path
import matplotlib
import pyproj

matplotlib.use("Agg")  # 使用Agg后端，不显示图形窗口


class SHPGUI(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("SHP文件转JSON转换器")
        self.geometry("800x800")  # 页面整体调高
        self.configure(bg="#f0f0f0")

        # 设置窗口图标
        self.set_window_icon()

        # 设置中文字体支持
        self.option_add("*Font", "SimHei 10")

        # 初始化变量
        self.shp_path = tk.StringVar()
        self.json_path = tk.StringVar()
        self.projection_info = tk.StringVar()
        self.projection_info.set("未选择文件")

        # 为每个文件选择器维护独立的目录历史
        self.shp_last_dir = ""
        self.json_last_dir = ""

        # 投影选项
        self.embed_crs_option = tk.BooleanVar()
        self.embed_crs_option.set(True)  # 默认嵌入CRS信息（如果可用）

        # 创建界面
        self.create_widgets()

    def set_window_icon(self):
        """设置窗口图标"""
        try:
            # 获取当前脚本所在目录
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(script_dir, "icon.ico")

            # 设置窗口图标
            self.iconbitmap(icon_path)
        except Exception as e:
            print(f"无法设置窗口图标: {e}")
            # 不影响程序运行，继续执行

    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.pack(fill=tk.X, pady=5)

        # 创建自定义样式以增加输入框高度
        style = ttk.Style()
        style.configure("TEntry", padding=(10, 8))  # 通过padding增加高度

        # 使用grid布局设置行列权重
        file_frame.columnconfigure(1, weight=1)  # 让输入框列可以扩展

        # 创建一个子框架来包含输入控件
        input_frame = ttk.Frame(file_frame)
        input_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))

        # SHP文件选择
        ttk.Label(input_frame, text="SHP文件:").grid(row=0, column=0, sticky=tk.W, pady=5)

        shp_entry = ttk.Entry(input_frame, textvariable=self.shp_path, width=38, style="TEntry")
        shp_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        # 使Entry支持拖放
        shp_entry.drop_target_register(DND_FILES)
        shp_entry.dnd_bind('<<Drop>>', self.on_drop)

        browse_btn = ttk.Button(input_frame, text="浏览", command=self.browse_shp)
        browse_btn.grid(row=0, column=2, padx=5, pady=5)

        # JSON文件保存位置
        ttk.Label(input_frame, text="保存位置:").grid(row=1, column=0, sticky=tk.W, pady=5)

        json_entry = ttk.Entry(input_frame, textvariable=self.json_path, width=38, style="TEntry")
        json_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        save_btn = ttk.Button(input_frame, text="选择路径", command=self.select_save_path)
        save_btn.grid(row=1, column=2, padx=5, pady=5)

        # 投影选项
        proj_option_frame = ttk.Frame(file_frame)
        proj_option_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=5)

        ttk.Checkbutton(
            proj_option_frame,
            text="嵌入坐标系统信息 (仅当明确匹配到EPSG代码时)",
            variable=self.embed_crs_option
        ).pack(anchor=tk.W)

        # 转换按钮 - 使用grid放置在第0行第3列，跨越2行，并设置sticky为NSEW使其填充高度
        convert_btn = ttk.Button(file_frame, text="转换为JSON", command=self.convert_to_json)
        convert_btn.grid(row=0, column=3, rowspan=3, padx=5, pady=5, sticky=(tk.N, tk.S, tk.E))

        # 设置按钮样式使其更大
        style.configure("Big.TButton", font=("SimHei", 10), padding=(15, 5))
        convert_btn.configure(style="Big.TButton")

        # 投影信息区域 - 使用固定大小的ScrolledText
        proj_frame = ttk.LabelFrame(main_frame, text="投影信息", padding="10")
        proj_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 创建一个只读的ScrolledText组件
        self.projection_text = scrolledtext.ScrolledText(
            proj_frame,
            wrap=tk.WORD,
            width=80,
            height=15,  # 固定高度，约15行文本
            font=("SimHei", 10)
        )
        self.projection_text.pack(fill=tk.BOTH, expand=True)
        self.projection_text.insert(tk.END, "未选择文件")
        self.projection_text.config(state=tk.DISABLED)  # 设置为只读

        # 状态标签
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def on_drop(self, event):
        """处理拖放事件"""
        file_path = event.data.strip('{}')
        if file_path.lower().endswith('.shp'):
            self.shp_path.set(file_path)
            self.shp_last_dir = os.path.dirname(file_path)  # 更新SHP文件的目录历史
            self.update_projection_info()
        else:
            messagebox.showerror("错误", "请拖放SHP文件")

    def browse_shp(self):
        """浏览并选择SHP文件"""
        # 使用SHP文件的目录历史作为初始目录
        initial_dir = self.shp_last_dir if self.shp_last_dir else os.getcwd()

        file_path = filedialog.askopenfilename(
            title="选择SHP文件",
            initialdir=initial_dir,
            filetypes=[("SHP文件", "*.shp")]
        )

        if file_path:
            self.shp_path.set(file_path)
            self.shp_last_dir = os.path.dirname(file_path)  # 更新SHP文件的目录历史
            self.update_projection_info()

    def select_save_path(self):
        """选择JSON文件保存位置"""
        # 使用JSON文件的目录历史作为初始目录
        initial_dir = self.json_last_dir if self.json_last_dir else (
            os.path.dirname(self.shp_path.get()) if self.shp_path.get() else os.getcwd()
        )

        initial_file = os.path.splitext(os.path.basename(self.shp_path.get()))[
                           0] + ".json" if self.shp_path.get() else "output.json"

        file_path = filedialog.asksaveasfilename(
            title="保存JSON文件",
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json")]
        )

        if file_path:
            self.json_path.set(file_path)
            self.json_last_dir = os.path.dirname(file_path)  # 更新JSON文件的目录历史

    def update_projection_info(self):
        """更新投影信息"""
        try:
            if not self.shp_path.get():
                return

            self.status_var.set("正在读取投影信息...")
            self.update()

            # 读取SHP文件获取投影信息
            gdf = gpd.read_file(self.shp_path.get())
            crs = gdf.crs

            if crs:
                # 尝试更智能地获取EPSG代码
                epsg = self._get_epsg_code(crs)

                if crs.is_projected:
                    proj_type = "投影坐标系统"
                else:
                    proj_type = "地理坐标系统"

                proj_text = f"{proj_type}: {crs.name}\n" \
                            f"EPSG代码: {epsg}\n" \
                            f"WKT: {crs.to_wkt(pretty=True)}"
            else:
                proj_text = "无法获取投影信息"

            # 更新文本框内容
            self.projection_text.config(state=tk.NORMAL)
            self.projection_text.delete(1.0, tk.END)
            self.projection_text.insert(tk.END, proj_text)
            self.projection_text.config(state=tk.DISABLED)

            self.status_var.set("就绪")
        except Exception as e:
            # 更新文本框内容
            self.projection_text.config(state=tk.NORMAL)
            self.projection_text.delete(1.0, tk.END)
            self.projection_text.insert(tk.END, f"读取投影信息时出错: {str(e)}")
            self.projection_text.config(state=tk.DISABLED)

            self.status_var.set("就绪")

    def _get_epsg_code(self, crs):
        """尝试多种方法获取EPSG代码"""
        # 首先尝试直接获取EPSG
        epsg = crs.to_epsg()
        if epsg:
            return epsg

        # 如果直接获取失败，尝试使用pyproj进行匹配
        try:
            # 创建pyproj CRS对象
            pyproj_crs = pyproj.CRS(crs.to_wkt())

            # 尝试查找最近似的EPSG，要求至少50%的置信度
            epsg = pyproj_crs.to_epsg(min_confidence=50)
            if epsg:
                return epsg
            else:
                return None
        except:
            return None

    def convert_to_json(self):
        """将SHP文件转换为JSON"""
        try:
            shp_file = self.shp_path.get()
            json_file = self.json_path.get()

            if not shp_file:
                messagebox.showerror("错误", "请选择SHP文件")
                return

            if not json_file:
                messagebox.showerror("错误", "请选择保存位置")
                return

            if not os.path.exists(shp_file):
                messagebox.showerror("错误", "SHP文件不存在")
                return

            self.status_var.set("正在转换...")
            self.update()

            # 读取SHP文件
            gdf = gpd.read_file(shp_file)
            crs = gdf.crs

            # 保留原始投影
            geojson_data = json.loads(gdf.to_json())

            # 获取EPSG代码
            epsg = self._get_epsg_code(crs)

            # 如果用户选择嵌入CRS信息并且我们有明确的EPSG代码
            if self.embed_crs_option.get() and epsg:
                # 嵌入CRS信息 (非标准但被广泛支持)
                geojson_data["crs"] = {
                    "type": "name",
                    "properties": {
                        "name": f"urn:ogc:def:crs:EPSG::{epsg}"
                    }
                }
                crs_info = f"{crs.name} (EPSG:{epsg})"
            else:
                # 不嵌入CRS信息，使用GeoJSON默认的WGS84
                crs_info = f"{crs.name} (未嵌入CRS信息，使用GeoJSON默认的WGS84)"

            # 保存为JSON文件
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(geojson_data, f, ensure_ascii=False, indent=2)

            self.status_var.set("转换完成")
            messagebox.showinfo("成功", f"文件已成功转换并保存到:\n{json_file}\n\n投影方式: {crs_info}")

        except Exception as e:
            self.status_var.set("转换失败")
            messagebox.showerror("错误", f"转换过程中出错: {str(e)}")


if __name__ == "__main__":
    # 确保中文显示正常
    if sys.platform.startswith('win'):
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(2)

    app = SHPGUI()
    app.mainloop()