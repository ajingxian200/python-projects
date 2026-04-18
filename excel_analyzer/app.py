import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading
import os
from datetime import date, timedelta


def validate_columns(df, required_cols, file_label):
    """验证 DataFrame 包含必需列，缺失时抛出 ValueError。"""
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"{file_label} 缺少必需列: {', '.join(missing)}")


def process_csv_data(df_yxd, df_csr, df_hyd):
    """
    纯数据处理函数，可独立于 UI 测试。

    1. 先从原始 CSR 筛选"意向单自动下INQ"数据（去重前）
    2. 对 CSR 按"下单用户id"去重（保留第一条）
    3. 对会员店按"酒店ID"去重（保留第一条）
    4. 关联键统一转为 str 类型并 strip()
    5. 意向单 LEFT JOIN CSR（on user_id = 下单用户id），取"来源平台"→"下INQ渠道"
    6. 意向单 LEFT JOIN 会员店（on 酒店ID = 酒店ID），取"是否会员店"→"是否协议酒店"
    7. 选择并排序输出列（17列）
    8. 返回 (结果DataFrame, 自动下INQ DataFrame)
    """
    # 先用原始 CSR（去重前）筛选"意向单自动下INQ"数据
    valid_sources = ["通过意向单创建", "通过引导弹窗下单"]
    df_auto_inq = df_csr[
        df_csr["来源线索"].isin(valid_sources) & df_csr["运营"].notna() & (df_csr["运营"].astype(str).str.strip() != "")
    ].copy()

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

    # LEFT JOIN 意向单 with CSR on user_id = 下单用户id, 取"来源平台"列
    csr_subset = df_csr[["下单用户id", "来源平台"]].rename(columns={"来源平台": "下INQ渠道"})
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

    return result, df_auto_inq


def get_week_range():
    """返回上周五到本周四的日期范围字符串，如 '0410-0416'。"""
    today = date.today()
    # 本周四: weekday() 中 Thursday=3
    days_since_thu = (today.weekday() - 3) % 7
    this_thursday = today - timedelta(days=days_since_thu)
    last_friday = this_thursday - timedelta(days=6)
    return last_friday.strftime("%m%d") + "-" + this_thursday.strftime("%m%d")


def build_city_summary(result, df_auto_inq):
    """
    根据第一个 sheet (result) 和第二个 sheet (df_auto_inq) 生成分城市汇总。

    - 未下INQ线索数: 下INQ渠道为空时，按城市计数
    - 未下INQ用户数: 下INQ渠道为空时，按城市对 user_id 去重计数
    - 意向单转为INQ订单数: df_auto_inq 中来源线索='通过意向单创建'，按城市计数
    - 意向单转为RFP订单数: df_auto_inq 中来源线索='通过意向单创建' 且 审核状态='已通过审核'，按城市计数
    """
    # 下INQ渠道为空的记录
    no_inq = result[result["下INQ渠道"].isna() | (result["下INQ渠道"].astype(str).str.strip() == "")]

    # 未下INQ线索数（按城市计数）
    col1 = no_inq.groupby("城市").size().reset_index(name="未下INQ线索数")

    # 未下INQ用户数（按城市对 user_id 去重计数）
    col2 = no_inq.groupby("城市")["user_id"].nunique().reset_index(name="未下INQ用户数")

    # 意向单转为INQ订单数
    inq_created = df_auto_inq[df_auto_inq["来源线索"] == "通过意向单创建"]
    col3 = inq_created.groupby("城市").size().reset_index(name="意向单转为INQ订单数")

    # 意向单转为RFP订单数
    rfp = inq_created[inq_created["审核状态"] == "已通过审核"]
    col4 = rfp.groupby("城市").size().reset_index(name="意向单转为RFP订单数")

    # 合并
    summary = col1.merge(col2, on="城市", how="outer") \
                   .merge(col3, on="城市", how="outer") \
                   .merge(col4, on="城市", how="outer")
    summary = summary.fillna(0)
    for c in ["未下INQ线索数", "未下INQ用户数", "意向单转为INQ订单数", "意向单转为RFP订单数"]:
        summary[c] = summary[c].astype(int)
    summary = summary.sort_values("未下INQ线索数", ascending=False, ignore_index=True)

    return summary


# 数据环比情况的列名（含重复列名，用列表保持顺序）
WEEKLY_COLUMNS = [
    "周", "累计线索", "累计用户", "未下INQ线索", "未下INQ用户",
    "自动下INQ", "RFP", "转化率", "潜客线索", "RFP", "转化率",
    "引导弹窗", "RFP", "转化率",
]


def build_weekly_comparison(result, df_auto_inq, history_path):
    """
    生成"数据环比情况"sheet。

    1. 从历史数据读取第3-5行（即去掉标题后的前3行数据）作为前3行
    2. 根据当前数据计算本周行
    3. 追加环比行
    """
    # --- 读取历史数据 ---
    ext = os.path.splitext(history_path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df_hist = pd.read_excel(history_path, sheet_name="数据环比情况", header=None)
    else:
        # CSV / TSV
        try:
            df_hist = pd.read_csv(history_path, header=None, encoding="utf-8", sep=None, engine="python")
        except UnicodeDecodeError:
            df_hist = pd.read_csv(history_path, header=None, encoding="gbk", sep=None, engine="python")

    # 取第3-5行（0-indexed: row 0=标题, row 1~3=数据行, row 4=环比）
    # 即 iloc[2:5] 对应原文件第3、4、5行
    hist_rows = df_hist.iloc[2:5].reset_index(drop=True)

    # --- 计算本周数据 ---
    week_range = get_week_range()

    # 累计线索 = 意向单详细数据总行数
    total_leads = len(result)

    # 累计用户 = user_id 去重数
    total_users = result["user_id"].nunique()

    # 下INQ渠道为空的记录
    no_inq = result[result["下INQ渠道"].isna() | (result["下INQ渠道"].astype(str).str.strip() == "")]
    no_inq_leads = len(no_inq)
    no_inq_users = no_inq["user_id"].nunique()

    # 自动下INQ: 来源线索='通过意向单创建' 且 运营不为空
    auto_inq_created = df_auto_inq[df_auto_inq["来源线索"] == "通过意向单创建"]
    auto_inq_count = len(auto_inq_created)

    # RFP(自动下INQ): 以上 + 审核状态='已通过审核'
    auto_rfp_count = len(auto_inq_created[auto_inq_created["审核状态"] == "已通过审核"])
    auto_rate = f"{auto_rfp_count / auto_inq_count * 100:.2f}%" if auto_inq_count > 0 else "0.00%"

    # 潜客线索: 来源线索='通过意向单创建' 且 source_tags='intent,potential_transfer'
    potential = auto_inq_created[
        auto_inq_created["source_tags"].astype(str).str.strip() == "intent,potential_transfer"
    ]
    potential_count = len(potential)
    potential_rfp = len(potential[potential["审核状态"] == "已通过审核"])
    potential_rate = f"{potential_rfp / potential_count * 100:.2f}%" if potential_count > 0 else "0.00%"

    # 引导弹窗: 来源线索='通过引导弹窗下单'
    popup = df_auto_inq[df_auto_inq["来源线索"] == "通过引导弹窗下单"]
    popup_count = len(popup)
    popup_rfp = len(popup[popup["审核状态"] == "已通过审核"])
    popup_rate = f"{popup_rfp / popup_count * 100:.2f}%" if popup_count > 0 else "0.00%"

    current_row = [
        week_range, total_leads, total_users, no_inq_leads, no_inq_users,
        auto_inq_count, auto_rfp_count, auto_rate,
        potential_count, potential_rfp, potential_rate,
        popup_count, popup_rfp, popup_rate,
    ]

    # --- 组装 DataFrame ---
    # 标题行
    header_row = WEEKLY_COLUMNS
    # 历史3行 + 本周行
    data_rows = []
    for i in range(len(hist_rows)):
        row = hist_rows.iloc[i].tolist()[:14]
        # 确保转化率列（索引 7, 10, 13）显示为百分比字符串
        for idx in (7, 10, 13):
            if idx < len(row):
                val = row[idx]
                if isinstance(val, float) and val < 1:
                    row[idx] = f"{val * 100:.2f}%"
                elif isinstance(val, str) and "%" not in val:
                    try:
                        row[idx] = f"{float(val) * 100:.2f}%"
                    except ValueError:
                        pass
        data_rows.append(row)
    data_rows.append(current_row)

    # --- 计算环比行 ---
    # 环比 = (本周行 - 上周行) / 上周行，针对数值列
    prev_row = data_rows[-2]  # 上一行（历史最后一行）
    curr_row = data_rows[-1]  # 本周行
    ratio_row = ["环比"]
    for j in range(1, 14):
        prev_val = _to_number(prev_row[j])
        curr_val = _to_number(curr_row[j])
        if prev_val and prev_val != 0:
            ratio_row.append(f"{(curr_val - prev_val) / prev_val * 100:.2f}%")
        else:
            ratio_row.append("0.00%")
    data_rows.append(ratio_row)

    # 构建最终 DataFrame（不设 columns 以避免重复列名问题）
    all_rows = [header_row] + data_rows
    df_result = pd.DataFrame(all_rows)

    return df_result


def _to_number(val):
    """将值转为数字，百分比字符串转为小数值，无法转换返回 0。"""
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace("%", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0


class ExcelAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("报表分析工具")
        self.root.geometry("620x440")
        self.root.resizable(False, False)

        # CSV to Excel mode vars
        self.csv_yxd_path = tk.StringVar()
        self.csv_csr_path = tk.StringVar()
        self.csv_hyd_path = tk.StringVar()
        self.csv_output_path = tk.StringVar()
        self.history_path = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # --- CSV 文件选择 ---
        frame_csv_files = ttk.LabelFrame(self.root, text="数据文件选择", padding=10)
        frame_csv_files.pack(fill="x", **pad)

        csv_file_rows = [
            ("意向单 CSV:", self.csv_yxd_path, 0, False),
            ("CSR CSV:", self.csv_csr_path, 1, False),
            ("会员店 CSV:", self.csv_hyd_path, 2, False),
            ("历史数据:", self.history_path, 3, False),
            ("输出文件:", self.csv_output_path, 4, True),
        ]
        for label_text, var, row, is_save in csv_file_rows:
            ttk.Label(frame_csv_files, text=label_text).grid(row=row, column=0, sticky="w")
            ttk.Entry(frame_csv_files, textvariable=var, width=50).grid(row=row, column=1, padx=5)
            # 历史数据行用 Excel 浏览，其余不变
            if label_text == "历史数据:":
                cmd = (lambda v=var: self._browse_excel(v))
            else:
                cmd = (lambda v=var, s=is_save: self._browse_csv(v, save=s))
            ttk.Button(frame_csv_files, text="浏览...", command=cmd).grid(row=row, column=2)

        self.btn_csv_run = ttk.Button(self.root, text="开始生成", command=self._run_csv_to_excel)
        self.btn_csv_run.pack(pady=10)

        # --- 进度与日志 ---
        self.progress = ttk.Progressbar(self.root, mode="indeterminate", length=580)
        self.progress.pack(**pad)

        self.log_text = tk.Text(self.root, height=8, state="disabled", wrap="word")
        self.log_text.pack(fill="both", **pad)

    def _browse_csv(self, var, save=False):
        if save:
            ftypes = [("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=ftypes)
        else:
            ftypes = [("CSV 文件", "*.csv"), ("所有文件", "*.*")]
            path = filedialog.askopenfilename(filetypes=ftypes)
        if path:
            var.set(path)

    def _browse_excel(self, var):
        ftypes = [("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        path = filedialog.askopenfilename(filetypes=ftypes)
        if path:
            var.set(path)

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
            validate_columns(df_csr, ["下单用户id", "来源平台", "来源线索", "运营", "城市", "审核状态", "source_tags"], "CSR CSV")
            validate_columns(df_hyd, ["酒店ID", "是否会员店"], "会员店CSV")

            self._log("执行数据关联处理...")
            result, df_auto_inq = process_csv_data(df_yxd, df_csr, df_hyd)
            self._log(f"  -> 关联完成，结果 {result.shape[0]} 行, {result.shape[1]} 列")
            self._log(f"  -> 意向单自动下INQ数据 {df_auto_inq.shape[0]} 行")

            self._log("生成分城市汇总...")
            df_city_summary = build_city_summary(result, df_auto_inq)
            week_range = get_week_range()
            sheet3_name = f"意向单分城市({week_range})"
            self._log(f"  -> 分城市汇总 {df_city_summary.shape[0]} 行, 日期范围: {week_range}")

            # 生成数据环比情况（第四个 sheet）
            hist = self.history_path.get()
            df_weekly = None
            if hist:
                self._log(f"读取历史数据: {os.path.basename(hist)}")
                df_weekly = build_weekly_comparison(result, df_auto_inq, hist)
                self._log(f"  -> 数据环比情况生成完成，共 {df_weekly.shape[0]} 行")

            self._log(f"写入 Excel: {os.path.basename(output_path)}")
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                result.to_excel(writer, sheet_name="意向单详细数据", index=False)
                df_auto_inq.to_excel(writer, sheet_name="意向单自动下INQ数据", index=False)
                df_city_summary.to_excel(writer, sheet_name=sheet3_name, index=False)
                if df_weekly is not None:
                    df_weekly.to_excel(writer, sheet_name="数据环比情况", index=False, header=False)

            self._log(f"✅ 生成完成，结果已保存到: {os.path.basename(output_path)}")
            self.root.after(0, lambda: messagebox.showinfo("完成", f"结果已保存到:\n{output_path}"))

        except Exception as e:
            self._log(f"❌ 错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            self.root.after(0, self._stop_progress)

    def _stop_progress(self):
        self.progress.stop()
        self.btn_csv_run.config(state="normal")

    def _log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")


if __name__ == "__main__":
    root = tk.Tk()
    ExcelAnalyzerApp(root)
    root.mainloop()
