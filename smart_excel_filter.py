import os
import pandas as pd
from datetime import datetime
import glob
import threading
from tkinter import *
from tkinter import ttk, filedialog, scrolledtext


# ==================== 全局缓存列名 ====================
CACHED_COLUMNS = set()

# ==================== 条件评估函数（支持 为空/不为空）====================
def evaluate_condition(series, operator, value):
    try:
        if series.isna().all():
            return pd.Series([False] * len(series), index=series.index)

        # 新增：为空 / 不为空
        if operator == "为空":
            return series.isna() | (series.astype(str).str.strip() == '')
        if operator == "不为空":
            return ~(series.isna() | (series.astype(str).str.strip() == ''))

        if operator in ["包含", "不包含", "等于", "不等于"]:
            series = series.astype(str).str.strip().fillna('')
            value = str(value).strip()

            if operator == "包含":
                return series.str.contains(value, na=False, case=False)
            elif operator == "不包含":
                return ~series.str.contains(value, na=False, case=False)
            elif operator == "等于":
                return series.str.lower() == value.lower()
            elif operator == "不等于":
                return series.str.lower() != value.lower()

        else:
            try:
                float_value = float(value)
                series = pd.to_numeric(series, errors='coerce')
                if operator == ">":
                    return (series > float_value)
                elif operator == ">=":
                    return (series >= float_value)
                elif operator == "<":
                    return (series < float_value)
                elif operator == "<=":
                    return (series <= float_value)
                elif operator == "==":
                    return (series == float_value)
                elif operator == "!=":
                    return (series != float_value)
            except (ValueError, TypeError):
                return pd.Series([False] * len(series), index=series.index)

        return pd.Series([False] * len(series), index=series.index)
    except Exception:
        return pd.Series([False] * len(series), index=series.index)


# ==================== 核心处理逻辑：每个组独立输出文件 ====================
def process_data(source_dir, output_dir, condition_groups, log_callback):
    """
    condition_groups: 列表，但通常只传一个组 [group_data]
    每次只处理一个组，输出一个文件
    """
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    pattern = os.path.join(source_dir, "*.xls*")
    excel_files = glob.glob(pattern)

    if not excel_files:
        log_callback("[WARN] ⚠️ 未在数据源目录找到 Excel 文件！请检查路径。")
        return

    log_callback(f"📦 发现 {len(excel_files)} 个 Excel 文件，开始处理...")

    for group_data in condition_groups:
        group_name = group_data['name'].strip() or f"结果组{group_data['index']}"
        log_callback(f"🔍 处理独立筛选组: [{group_name}]")

        conditions = group_data['conditions']
        if not conditions:
            log_callback(f"🟡 跳过: 组 '{group_name}' 无有效条件")
            continue

        collected_dfs = []

        for file in excel_files:
            try:
                filename = os.path.basename(file)
                engine = 'openpyxl' if file.endswith('.xlsx') else 'xlrd'
                df = pd.read_excel(file, engine=engine)

                if df.empty:
                    log_callback(f"🟡 跳过: {filename} 为空表")
                    continue

                str_cols = df.select_dtypes(include='object').columns
                df[str_cols] = df[str_cols].astype(str).fillna('').apply(lambda x: x.str.strip())

                final_mask = None

                for cond in conditions:
                    col, op, val = cond['col'], cond['op'], cond['val']

                    if not col:
                        continue

                    if col not in df.columns:
                        log_callback(f"  ⚠️ 列 '{col}' 不存在于 {filename}")
                        continue

                    mask = evaluate_condition(df[col], op, val)
                    if final_mask is None:
                        final_mask = mask
                    else:
                        logic = cond['logic']
                        if logic == "AND":
                            final_mask &= mask
                        else:
                            final_mask |= mask

                if final_mask is not None and final_mask.any():
                    matched_df = df[final_mask].copy()
                    collected_dfs.append(matched_df)
                    log_callback(f"  ✅ {filename} → 匹配 {len(matched_df)} 行")

            except Exception as e:
                log_callback(f"  ❌ 读取失败: {filename} | {str(e)[:100]}...")

        if collected_dfs:
            try:
                result_df = pd.concat(collected_dfs, ignore_index=True)
                result_df.drop_duplicates(inplace=True)

                # 清理文件名
                safe_name = "".join(c for c in group_name if c.isalnum() or c in " _-")
                output_filename = f"{safe_name}_{timestamp}.xlsx"
                output_path = os.path.join(output_dir, output_filename)

                result_df.to_excel(output_path, index=False)
                log_callback(f"🎉 已保存 {len(result_df)} 行 → {output_filename}")
            except Exception as e:
                log_callback(f"[FAIL] 保存失败: {output_filename} | {e}")
        else:
            log_callback(f"🟡 组 '{group_name}' 未匹配到任何数据")

    log_callback(f"✅ 筛选组 [{group_name}] 处理完成。")


# ==================== 扫描列名（修改版）====================
def scan_columns(source_dir, callback):
    global CACHED_COLUMNS
    CACHED_COLUMNS = set()
    pattern = os.path.join(source_dir, "*.xls*")
    files = glob.glob(pattern)
    if not files:
        callback("⚠️ 未找到 Excel 文件")
        return
    # 只处理第一个文件
    first_file = files[0]
    try:
        engine = 'openpyxl' if first_file.endswith('.xlsx') else 'xlrd'
        df = pd.read_excel(first_file, engine=engine, nrows=1)
        CACHED_COLUMNS.update(df.columns.tolist())
        callback(f"✅ 成功从 {os.path.basename(first_file)} 提取 {len(CACHED_COLUMNS)} 个唯一列名")
    except Exception as e:
        callback(f"  ⚠️ 读取 {os.path.basename(first_file)} 失败: {str(e)}")

# ==================== GUI 主程序（最终版）====================
class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎯 智能 Excel 筛选器 · 独立输出版")
        self.root.geometry("950x750")
        self.root.minsize(800, 600)

        self.default_source = "datasource"
        self.default_output = "output"
        os.makedirs(self.default_source, exist_ok=True)
        os.makedirs(self.default_output, exist_ok=True)

        self.condition_groups = []
        self.create_widgets()

    def create_widgets(self):
        main_container = Frame(self.root)
        main_container.pack(fill=BOTH, expand=True, padx=10, pady=5)

        canvas = Canvas(main_container)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # ========= 上层：路径设置 + 列名扫描 =========
        top_frame = LabelFrame(scrollable_frame, text=" 📁 数据路径 & 列名 ", font=("微软雅黑", 10, "bold"))
        top_frame.pack(fill=X, pady=5)

        # 数据源
        Label(top_frame, text="数据源:", font=("微软雅黑", 9)).grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.src_var = StringVar(value=os.path.abspath(self.default_source))
        Entry(top_frame, textvariable=self.src_var, width=45).grid(row=0, column=1, padx=5)
        Button(top_frame, text="📁", command=self.browse_source, width=3).grid(row=0, column=2, padx=2)

        # 扫描列名
        Button(top_frame, text="🔄 扫描列名", command=self.scan_columns_now,
               bg="#17a2b8", fg="white", width=10).grid(row=0, column=3, padx=5)

        # 输出目录
        Label(top_frame, text="输出目录:", font=("微软雅黑", 9)).grid(row=1, column=0, sticky=W, padx=5, pady=5)
        self.out_var = StringVar(value=os.path.abspath(self.default_output))
        Entry(top_frame, textvariable=self.out_var, width=45).grid(row=1, column=1, padx=5)
        Button(top_frame, text="📁", command=self.browse_output, width=3).grid(row=1, column=2, padx=2)

        # ========= 中层：条件组容器 =========
        self.conditions_parent = Frame(scrollable_frame)
        self.conditions_parent.pack(fill=BOTH, expand=True, pady=5)

        # ✅ 先创建“添加条件组”按钮，再添加第一个组
        btn_add_group = Button(scrollable_frame, text="➕ 添加条件组", command=self.add_condition_group,
                               font=("微软雅黑", 9), bg="#28a745", fg="white")
        btn_add_group.pack(fill=X, pady=5)

        self.add_condition_group()  # 添加第一个组

        # ========= 下层：日志和执行按钮 =========
        bottom_frame = Frame(self.root)
        bottom_frame.pack(side=BOTTOM, fill=X, padx=10, pady=5)

        log_frame = Frame(bottom_frame)
        log_frame.pack(fill=BOTH, expand=True)
        self.log_area = scrolledtext.ScrolledText(log_frame, height=12, font=("Consolas", 9), bg="#f8f9fa")
        self.log_area.pack(fill=BOTH, expand=True)

        btn_frame = Frame(bottom_frame)
        btn_frame.pack(fill=X, pady=5)
        Button(btn_frame, text="🚀 执行", command=self.start_processing,
               bg="#007BFF", fg="white", width=12, font=("微软雅黑", 10)).pack(side=LEFT, padx=5)
        Button(btn_frame, text="🗑️ 清空日志", command=lambda: self.log_area.delete(1.0, END),
               width=10, font=("微软雅黑", 10)).pack(side=RIGHT, padx=5)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.bind_mousewheel(canvas)

    def bind_mousewheel(self, canvas):
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _on_enter(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _on_leave(event):
            canvas.unbind_all("<MouseWheel>")
        for widget in [canvas, self.log_area]:
            widget.bind("<Enter>", _on_enter)
            widget.bind("<Leave>", _on_leave)

    def browse_source(self):
        path = filedialog.askdirectory()
        if path:
            self.src_var.set(path)

    def browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.out_var.set(path)

    def scan_columns_now(self):
        source_dir = self.src_var.get().strip()
        if not source_dir or not os.path.exists(source_dir):
            self.log("❌ 请先设置有效的数据源目录！")
            return

        def task():
            scan_columns(source_dir, self.log)

        threading.Thread(target=task, daemon=True).start()

    def add_condition_group(self):
        index = len(self.condition_groups) + 1
        group_frame = LabelFrame(self.conditions_parent, text=f" 筛选组 {index} ", padx=10, pady=5)
        group_frame.pack(fill=X, pady=5)

        name_frame = Frame(group_frame)
        name_frame.pack(fill=X, pady=2)
        Label(name_frame, text="命名:", width=6).pack(side=LEFT)
        name_var = StringVar(value=f"结果{index}")
        Entry(name_frame, textvariable=name_var, width=20).pack(side=LEFT, padx=5)

        conditions_list = Frame(group_frame)
        conditions_list.pack(fill=X, pady=5)

        # 初始化第一行条件
        self.add_condition_row(conditions_list)

        btn_frame = Frame(group_frame)
        btn_frame.pack(fill=X, pady=2)
        # 删除组按钮
        Button(btn_frame, text="🗑️ 删除本组", command=lambda: self.remove_group(group_frame),
               fg="red", font=("Arial", 9)).pack(side=RIGHT)

        self.condition_groups.append({
            'frame': group_frame,
            'name_var': name_var,
            'conditions_list': conditions_list,
            'index': index
        })

    def add_condition_row(self, parent):
        row_frame = Frame(parent)
        row_frame.pack(fill=X, pady=1)

        # 字段名 → 下拉+输入
        col_var = StringVar(value="字段名")
        col_combo = ttk.Combobox(row_frame, textvariable=col_var, values=sorted(CACHED_COLUMNS), width=14)
        col_combo.pack(side=LEFT, padx=2)
        col_combo.bind('<FocusIn>', lambda e: col_combo.configure(values=sorted(CACHED_COLUMNS)))

        # 操作符
        op_var = StringVar(value="包含")
        ops = ["包含", "不包含", "等于", "不等于", ">", ">=", "<", "<=", "为空", "不为空"]
        op_menu = OptionMenu(row_frame, op_var, *ops)
        op_menu.config(width=8)
        op_menu.pack(side=LEFT, padx=2)

        # 值输入
        val_var = StringVar(value="关键词")
        val_entry = Entry(row_frame, textvariable=val_var, width=18)
        val_entry.pack(side=LEFT, padx=2)

        # 逻辑连接符（组内条件）
        logic_var = StringVar(value="AND")
        logic_menu = OptionMenu(row_frame, logic_var, "AND", "OR")
        logic_menu.config(width=5)
        logic_menu.pack(side=LEFT, padx=2)

        # ➕ 按钮：在当前行下方插入新条件
        Button(row_frame, text="➕", width=2, command=lambda: self.add_condition_row(parent)).pack(side=LEFT, padx=1)

        # ➖ 按钮：删除当前行
        Button(row_frame, text="➖", width=2, command=row_frame.destroy).pack(side=LEFT, padx=1)

        def on_op_change(*_):
            if op_var.get() in ["为空", "不为空"]:
                val_entry.config(state=DISABLED)
            else:
                val_entry.config(state=NORMAL)

        op_var.trace('w', on_op_change)
        on_op_change()

    def remove_group(self, frame):
        if len(self.condition_groups) <= 1:
            return
        frame.destroy()
        self.condition_groups = [g for g in self.condition_groups if g['frame'] != frame]

    def log(self, msg):
        self.log_area.insert(END, msg + "\n")
        self.log_area.see(END)
        self.root.update_idletasks()

    def start_processing(self):
        for widget in self.root.winfo_children():
            if isinstance(widget, Button):
                widget.config(state=DISABLED)
        threading.Thread(target=self.run_process, daemon=True).start()

    def run_process(self):
        try:
            self.log("⏳ 正在启动...")
            source_dir = self.src_var.get().strip()
            output_dir = self.out_var.get().strip()

            if not source_dir or not os.path.exists(source_dir):
                self.log("[FAIL] ❌ 数据源目录不存在！")
                return
            if not output_dir:
                self.log("[FAIL] ❌ 请设置输出目录！")
                return

            valid_groups = []
            for i, group in enumerate(self.condition_groups):
                group_name = group['name_var'].get().strip()
                conditions = []

                for row in group['conditions_list'].winfo_children():
                    if not isinstance(row, Frame):
                        continue
                    widgets = row.winfo_children()
                    if len(widgets) < 6:
                        continue

                    col_widget = widgets[0]
                    op_widget = widgets[1]
                    val_widget = widgets[2]
                    logic_widget = widgets[3]

                    col = col_widget.get() if hasattr(col_widget, 'get') else ''
                    op = op_widget.cget("text") if isinstance(op_widget, Menubutton) else op_widget.get()
                    val = val_widget.get() if isinstance(val_widget, Entry) else ''
                    logic = logic_widget.cget("text") if isinstance(logic_widget, Menubutton) else logic_widget.get()

                    if col and col != "字段名":
                        if op in ["为空", "不为空"]:
                            conditions.append({'col': col, 'op': op, 'val': '', 'logic': logic})
                        elif val and val != "关键词":
                            conditions.append({'col': col, 'op': op, 'val': val, 'logic': logic})

                if conditions:
                    valid_groups.append({
                        'name': group_name,
                        'conditions': conditions,
                        'index': i + 1
                    })
                else:
                    self.log(f"🟡 跳过: 组 '{group_name}' 无有效条件")

            if not valid_groups:
                self.log("[WARN] ⚠️ 所有组均无有效条件，请检查输入！")
                return

            # ✅ 每个组独立执行
            for group_data in valid_groups:
                process_data(source_dir, output_dir, [group_data], self.log)

            self.log("🎉 所有独立筛选任务已完成！")

        except Exception as e:
            self.log(f"[FAIL] 💥 系统异常: {str(e)}")
        finally:
            for widget in self.root.winfo_children():
                if isinstance(widget, Button):
                    widget.config(state=NORMAL)


# ==================== 启动 ====================
if __name__ == "__main__":
    root = Tk()
    app = ExcelFilterApp(root)
    root.mainloop()