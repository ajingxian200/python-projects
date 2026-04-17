import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
import os


def validate_columns(df, required_cols, file_label):
    """验证 DataFrame 包含必需列，缺失时抛出 ValueError。"""
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"{file_label} 缺少必需列: {', '.join(missing)}")


def process_csv_data(df_yxd, df_csr, df_hyd):
    """
    纯数据处理函数，可独立于 UI 测试。

    1. 对 CSR 按"下单用户id"去重（保留第一条）
    2. 对会员店按"酒店ID"去重（保留第一条）
    3. 关联键统一转为 str 类型并 strip()
    4. 意向单 LEFT JOIN CSR（on user_id = 下单用户id），取"来源线索"→"下INQ渠道"
    5. 意向单 LEFT JOIN 会员店（on 酒店ID = 酒店ID），取"是否会员店"→"是否协议酒店"
    6. 选择并排序输出列（17列）
    7. 返回结果 DataFrame
    """
    # 去重
    df_csr = df_csr.drop_duplicates(subset=["下单用户id"], keep="first")
    df_hyd = df_hyd.drop_duplicates(subset=["酒店ID"], keep="first")

    # 关联键统一转为 str 并 strip
    df_yxd = df_yxd.copy()
    df_yxd["user_id"] = df_yxd["user_id"].astype(str).str.strip()
    df_yxd["酒店ID"] = df_yxd["酒店ID"].astype(str).str.strip()

    df_csr = df_csr.copy()
    df_csr["下单用户id"] = df_csr["下单用户id"].astype(str).str.strip()

    df_hyd = df_hyd.copy()
    df_hyd["酒店ID"] = df_hyd["酒店ID"].astype(str).str.strip()

    # LEFT JOIN 意向单 with CSR on user_id = 下单用户id, 取"来源线索"列
    csr_subset = df_csr[["下单用户id", "来源线索"]].rename(columns={"来源线索": "下INQ渠道"})
    result = df_yxd.merge(csr_subset, left_on="user_id", right_on="下单用户id", how="left")
    result.drop(columns=["下单用户id"], inplace=True)

    # LEFT JOIN 意向单 with 会员店 on 酒店ID = 酒店ID, 取"是否会员店"列
    hyd_subset = df_hyd[["酒店ID", "是否会员店"]].rename(columns={"是否会员店": "是否协议酒店"})
    result = result.merge(hyd_subset, on="酒店ID", how="left")

    # 选择并排序输出列
    output_columns = [
        "number", "省", "城市", "等级", "user_type", "user_id", "name",
        "客户手机号", "supplier_type", "supplier_id", "酒店ID", "酒店名称",
        "上架状态", "source", "created_at", "下INQ渠道", "是否协议酒店",
    ]
    result = result[output_columns]

    return result


class ExcelAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 工具箱")
        self.root.geometry("620x580")
        self.root.resizable(False, False)

        # Mode selection
        self.app_mode = tk.StringVar(value="excel_compare")

        # Excel compare mode vars
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.key_col = tk.StringVar()
        self.analysis_type = tk.StringVar(value="merge")

        # CSV to Excel mode vars
        self.csv_yxd_path = tk.StringVar()
        self.csv_csr_path = tk.StringVar()
        self.csv_hyd_path = tk.StringVar()
        self.csv_output_path = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # --- 模式选择区 ---
        frame_mode = ttk.LabelFrame(self.root, text="功能模式", padding=10)
        frame_mode.pack(fill="x", **pad)
        ttk.Radiobutton(
            frame_mode, text="Excel 对比分析", variable=self.app_mode,
            value="excel_compare", command=self._on_mode_change,
        ).pack(side="left", padx=20)
        ttk.Radiobutton(
            frame_mode, text="CSV 转 Excel 生成", variable=self.app_mode,
            value="csv_to_excel", command=self._on_mode_change,
        ).pack(side="left", padx=20)

        # === Excel 对比分析模式 frame ===
        self.frame_excel_mode = ttk.Frame(self.root)

        frame_files = ttk.LabelFrame(self.frame_excel_mode, text="文件选择", padding=10)
        frame_files.pack(fill="x", **pad)

        for label_text, var, row in [
            ("Excel 文件 1:", self.file1_path, 0),
            ("Excel 文件 2:", self.file2_path, 1),
            ("输出文件:", self.output_path, 2),
        ]:
            ttk.Label(frame_files, text=label_text).grid(row=row, column=0, sticky="w")
            ttk.Entry(frame_files, textvariable=var, width=50).grid(row=row, column=1, padx=5)
            is_save = row == 2
            cmd = (lambda v=var, s=is_save: self._browse(v, save=s))
            ttk.Button(frame_files, text="浏览...", command=cmd).grid(row=row, column=2)

        frame_settings = ttk.LabelFrame(self.frame_excel_mode, text="分析设置", padding=10)
        frame_settings.pack(fill="x", **pad)

        ttk.Label(frame_settings, text="分析类型:").grid(row=0, column=0, sticky="w")
        types = [
            ("按列合并 (类似VLOOKUP)", "merge"),
            ("数据对比 (找差异)", "diff"),
            ("相关性分析 (数值列)", "corr"),
        ]
        for i, (text, val) in enumerate(types):
            ttk.Radiobutton(frame_settings, text=text, variable=self.analysis_type, value=val).grid(
                row=i + 1, column=0, columnspan=3, sticky="w", padx=20
            )

        ttk.Label(frame_settings, text="关联列名 (合并/对比时必填):").grid(row=4, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(frame_settings, textvariable=self.key_col, width=30).grid(row=4, column=1, sticky="w", pady=(10, 0))

        self.btn_run = ttk.Button(self.frame_excel_mode, text="开始分析", command=self._run_analysis)
        self.btn_run.pack(pady=10)

        self.frame_excel_mode.pack(fill="x")

        # === CSV 转 Excel 模式 frame ===
        self.frame_csv_mode = ttk.Frame(self.root)

        frame_csv_files = ttk.LabelFrame(self.frame_csv_mode, text="CSV 文件选择", padding=10)
        frame_csv_files.pack(fill="x", **pad)

        csv_file_rows = [
            ("意向单 CSV:", self.csv_yxd_path, 0, False),
            ("CSR CSV:", self.csv_csr_path, 1, False),
            ("会员店 CSV:", self.csv_hyd_path, 2, False),
            ("输出文件:", self.csv_output_path, 3, True),
        ]
        for label_text, var, row, is_save in csv_file_rows:
            ttk.Label(frame_csv_files, text=label_text).grid(row=row, column=0, sticky="w")
            ttk.Entry(frame_csv_files, textvariable=var, width=50).grid(row=row, column=1, padx=5)
            cmd = (lambda v=var, s=is_save: self._browse_csv(v, save=s))
            ttk.Button(frame_csv_files, text="浏览...", command=cmd).grid(row=row, column=2)

        self.btn_csv_run = ttk.Button(self.frame_csv_mode, text="开始生成", command=self._run_csv_to_excel)
        self.btn_csv_run.pack(pady=10)

        # CSV mode frame is hidden by default (excel_compare is default)

        # --- 进度与日志 (shared) ---
        self.progress = ttk.Progressbar(self.root, mode="indeterminate", length=580)
        self.progress.pack(**pad)

        self.log_text = tk.Text(self.root, height=8, state="disabled", wrap="word")
        self.log_text.pack(fill="both", **pad)

    def _browse(self, var, save=False):
        ftypes = [("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        if save:
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=ftypes)
        else:
            path = filedialog.askopenfilename(filetypes=ftypes)
        if path:
            var.set(path)

    def _browse_csv(self, var, save=False):
        if save:
            ftypes = [("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=ftypes)
        else:
            ftypes = [("CSV 文件", "*.csv"), ("所有文件", "*.*")]
            path = filedialog.askopenfilename(filetypes=ftypes)
        if path:
            var.set(path)

    def _on_mode_change(self):
        if self.app_mode.get() == "excel_compare":
            self.frame_csv_mode.pack_forget()
            self.frame_excel_mode.pack(fill="x", before=self.progress)
        else:
            self.frame_excel_mode.pack_forget()
            self.frame_csv_mode.pack(fill="x", before=self.progress)

    def _run_csv_to_excel(self):
        yxd = self.csv_yxd_path.get()
        csr = self.csv_csr_path.get()
        hyd = self.csv_hyd_path.get()
        out = self.csv_output_path.get()
        if not all([yxd, csr, hyd]):
            messagebox.showwarning("提示", "请选择所有 CSV 文件路径")
            return
        if not out:
            messagebox.showwarning("提示", "请选择输出文件路径")
            return

        self.btn_csv_run.config(state="disabled")
        self.progress.start()
        threading.Thread(target=self._do_csv_to_excel, args=(yxd, csr, hyd, out), daemon=True).start()

    def _do_csv_to_excel(self, yxd_path, csr_path, hyd_path, output_path):
        try:
            self._log(f"读取意向单 CSV: {os.path.basename(yxd_path)}")
            try:
                df_yxd = pd.read_csv(yxd_path, encoding="utf-8")
            except UnicodeDecodeError:
                df_yxd = pd.read_csv(yxd_path, encoding="gbk")
            self._log(f"  -> {df_yxd.shape[0]} 行, {df_yxd.shape[1]} 列")

            self._log(f"读取 CSR CSV: {os.path.basename(csr_path)}")
            try:
                df_csr = pd.read_csv(csr_path, encoding="utf-8")
            except UnicodeDecodeError:
                df_csr = pd.read_csv(csr_path, encoding="gbk")
            self._log(f"  -> {df_csr.shape[0]} 行, {df_csr.shape[1]} 列")

            self._log(f"读取会员店 CSV: {os.path.basename(hyd_path)}")
            try:
                df_hyd = pd.read_csv(hyd_path, encoding="utf-8")
            except UnicodeDecodeError:
                df_hyd = pd.read_csv(hyd_path, encoding="gbk")
            self._log(f"  -> {df_hyd.shape[0]} 行, {df_hyd.shape[1]} 列")

            self._log("验证必需列...")
            validate_columns(df_yxd, ["user_id", "酒店ID"], "意向单CSV")
            validate_columns(df_csr, ["下单用户id", "来源线索"], "CSR CSV")
            validate_columns(df_hyd, ["酒店ID", "是否会员店"], "会员店CSV")

            self._log("执行数据关联处理...")
            result = process_csv_data(df_yxd, df_csr, df_hyd)
            self._log(f"  -> 关联完成，结果 {result.shape[0]} 行, {result.shape[1]} 列")

            self._log(f"写入 Excel: {os.path.basename(output_path)}")
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                result.to_excel(writer, sheet_name="意向单详细数据", index=False)

            self._log(f"✅ 生成完成，结果已保存到: {os.path.basename(output_path)}")
            self.root.after(0, lambda: messagebox.showinfo("完成", f"结果已保存到:\n{output_path}"))

        except Exception as e:
            self._log(f"❌ 错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            self.root.after(0, self._stop_csv_progress)

    def _stop_csv_progress(self):
        self.progress.stop()
        self.btn_csv_run.config(state="normal")

    def _log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _run_analysis(self):
        f1, f2, out = self.file1_path.get(), self.file2_path.get(), self.output_path.get()
        if not all([f1, f2, out]):
            messagebox.showwarning("提示", "请选择所有文件路径")
            return

        atype = self.analysis_type.get()
        key = self.key_col.get().strip()
        if atype in ("merge", "diff") and not key:
            messagebox.showwarning("提示", "合并/对比模式需要填写关联列名")
            return

        self.btn_run.config(state="disabled")
        self.progress.start()
        threading.Thread(target=self._do_analysis, args=(f1, f2, out, atype, key), daemon=True).start()

    def _do_analysis(self, f1, f2, out, atype, key):
        try:
            self._log(f"读取文件 1: {os.path.basename(f1)}")
            df1 = pd.read_excel(f1)
            self._log(f"  -> {df1.shape[0]} 行, {df1.shape[1]} 列")

            self._log(f"读取文件 2: {os.path.basename(f2)}")
            df2 = pd.read_excel(f2)
            self._log(f"  -> {df2.shape[0]} 行, {df2.shape[1]} 列")

            if atype == "merge":
                result = self._analyze_merge(df1, df2, key)
            elif atype == "diff":
                result = self._analyze_diff(df1, df2, key)
            else:
                result = self._analyze_corr(df1, df2)

            if isinstance(result, dict):
                with pd.ExcelWriter(out, engine="openpyxl") as writer:
                    for sheet_name, df in result.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                result.to_excel(out, index=False, engine="openpyxl")

            self._log(f"✅ 分析完成，结果已保存到: {os.path.basename(out)}")
            self.root.after(0, lambda: messagebox.showinfo("完成", f"结果已保存到:\n{out}"))

        except Exception as e:
            self._log(f"❌ 错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            self.root.after(0, self._stop_progress)

    def _stop_progress(self):
        self.progress.stop()
        self.btn_run.config(state="normal")

    def _analyze_merge(self, df1, df2, key):
        self._log(f"执行合并分析，关联列: {key}")
        merged = pd.merge(df1, df2, on=key, how="outer", suffixes=("_文件1", "_文件2"), indicator=True)
        merged["_merge"] = merged["_merge"].map({
            "left_only": "仅文件1", "right_only": "仅文件2", "both": "两者都有"
        })
        merged.rename(columns={"_merge": "来源"}, inplace=True)

        summary = merged["来源"].value_counts().reset_index()
        summary.columns = ["来源", "数量"]

        self._log(f"  合并结果: {merged.shape[0]} 行")
        return {"合并结果": merged, "汇总": summary}

    def _analyze_diff(self, df1, df2, key):
        self._log(f"执行对比分析，关联列: {key}")
        common_cols = [c for c in df1.columns if c in df2.columns and c != key]
        merged = pd.merge(df1, df2, on=key, how="inner", suffixes=("_文件1", "_文件2"))

        diffs = []
        for _, row in merged.iterrows():
            for col in common_cols:
                v1, v2 = row.get(f"{col}_文件1"), row.get(f"{col}_文件2")
                if pd.isna(v1) and pd.isna(v2):
                    continue
                if v1 != v2:
                    diffs.append({
                        key: row[key], "列名": col,
                        "文件1值": v1, "文件2值": v2
                    })

        result = pd.DataFrame(diffs) if diffs else pd.DataFrame(columns=[key, "列名", "文件1值", "文件2值"])
        self._log(f"  发现 {len(diffs)} 处差异")

        only1 = df1[~df1[key].isin(df2[key])]
        only2 = df2[~df2[key].isin(df1[key])]
        return {"差异明细": result, "仅文件1有": only1, "仅文件2有": only2}

    def _analyze_corr(self, df1, df2):
        self._log("执行相关性分析 (数值列)")
        num1 = df1.select_dtypes(include="number")
        num2 = df2.select_dtypes(include="number")

        combined = pd.concat([num1.add_suffix("_文件1"), num2.add_suffix("_文件2")], axis=1)
        corr = combined.corr()

        rows = []
        for c1 in num1.columns:
            for c2 in num2.columns:
                k1, k2 = f"{c1}_文件1", f"{c2}_文件2"
                if k1 in corr.columns and k2 in corr.columns:
                    rows.append({
                        "文件1列": c1, "文件2列": c2,
                        "相关系数": round(corr.loc[k1, k2], 4)
                    })

        result = pd.DataFrame(rows).sort_values("相关系数", key=abs, ascending=False, ignore_index=True)
        self._log(f"  计算了 {len(rows)} 对相关系数")

        stats1 = df1.describe().T.add_suffix("_文件1").reset_index().rename(columns={"index": "列名"})
        stats2 = df2.describe().T.add_suffix("_文件2").reset_index().rename(columns={"index": "列名"})
        return {"相关性": result, "文件1统计": stats1, "文件2统计": stats2}


if __name__ == "__main__":
    root = tk.Tk()
    ExcelAnalyzerApp(root)
    root.mainloop()
