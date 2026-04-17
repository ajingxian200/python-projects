# 实施计划：CSV 转 Excel 生成

## 概述

在现有 `excel_analyzer/app.py` 的 `ExcelAnalyzerApp` 类中扩展 CSV 转 Excel 功能。通过模式切换（RadioButton）在"Excel 对比分析"和"CSV 转 Excel 生成"之间切换 UI，复用现有日志、进度条、线程处理模式。数据处理逻辑提取为可独立测试的纯函数。

## 任务

- [x] 1. 提取纯数据处理函数并添加列验证逻辑
  - [x] 1.1 在 `app.py` 中新增 `validate_columns(df, required_cols, file_label)` 纯函数
    - 接收 DataFrame、必需列列表、文件标签
    - 缺失列时抛出 ValueError，错误信息包含所有缺失列名
    - _需求: 6.3, 6.4_
  - [x] 1.2 在 `app.py` 中新增 `process_csv_data(df_yxd, df_csr, df_hyd)` 纯函数
    - 对 CSR 按"下单用户id"去重（保留第一条）
    - 对会员店按"酒店ID"去重（保留第一条）
    - 关联键统一转为 `str` 类型并 `strip()`
    - 意向单 LEFT JOIN CSR（on user_id = 下单用户id），取"来源线索"列重命名为"下INQ渠道"
    - 意向单 LEFT JOIN 会员店（on 酒店ID = 酒店ID），取"是否会员店"列重命名为"是否协议酒店"
    - 选择并排序输出列（17列，顺序按设计文档）
    - 返回结果 DataFrame
    - _需求: 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 4.3, 4.4_
  - [ ]* 1.3 编写属性测试：LEFT JOIN 保持行数不变
    - **属性 1: LEFT JOIN 保持行数不变**
    - 使用 hypothesis 生成随机行数的意向单/CSR/会员店 DataFrame，ID 值随机（含部分重叠）
    - 验证 `process_csv_data` 返回行数等于意向单行数
    - **验证需求: 3.2, 3.3**
  - [ ]* 1.4 编写属性测试：列值映射正确性
    - **属性 2: 列值映射正确性**
    - 生成随机数据，确保部分 ID 匹配、部分不匹配
    - 验证匹配行的"下INQ渠道"等于 CSR 对应"来源线索"，"是否协议酒店"等于会员店对应"是否会员店"
    - **验证需求: 3.4, 3.5, 3.6, 3.7**
  - [ ]* 1.5 编写属性测试：输出列结构与顺序
    - **属性 3: 输出列结构与顺序**
    - 生成随机行数的有效 DataFrame
    - 验证结果恰好包含 17 列，且列顺序正确
    - **验证需求: 4.3, 4.4**
  - [ ]* 1.6 编写属性测试：列验证拒绝缺失必需列
    - **属性 4: 列验证拒绝缺失必需列**
    - 生成随机列名集合（随机移除必需列）
    - 验证 `validate_columns` 抛出 ValueError，且错误信息包含所有缺失列名
    - **验证需求: 6.3, 6.4**

- [x] 2. 检查点 - 确保数据处理逻辑和属性测试通过
  - 确保所有测试通过，如有问题请向用户确认。

- [-] 3. 实现模式切换 UI 和 CSV 文件选择控件
  - [-] 3.1 在 `_build_ui()` 方法顶部新增模式选择区域
    - 添加 `self.app_mode = tk.StringVar(value="excel_compare")` 模式变量
    - 添加两个 RadioButton："Excel 对比分析"和"CSV 转 Excel 生成"
    - _需求: 1.1_
  - [-] 3.2 实现 `_on_mode_change()` 方法
    - 根据 `self.app_mode` 值切换显示/隐藏两组 UI 控件（`pack()`/`pack_forget()`）
    - "excel_compare" 模式显示现有文件选择和分析设置控件
    - "csv_to_excel" 模式显示 CSV 文件选择控件
    - _需求: 1.2, 1.3_
  - [-] 3.3 新增 CSV 文件选择 UI 控件
    - 添加 `self.csv_yxd_path`、`self.csv_csr_path`、`self.csv_hyd_path`、`self.csv_output_path` 四个 StringVar
    - 创建三个 CSV 文件选择行（标签 + 输入框 + 浏览按钮），文件筛选器为 `[("CSV 文件", "*.csv"), ("所有文件", "*.*")]`
    - 创建输出文件路径选择行，默认保存格式为 .xlsx
    - 添加"开始生成"按钮，绑定 `_run_csv_to_excel()`
    - _需求: 2.1, 2.2, 2.3, 2.4_
  - [-] 3.4 将现有文件选择和分析设置控件包装为可切换的 frame
    - 将现有 `frame_files`、`frame_settings`、`btn_run` 包装到一个父 frame 中
    - CSV 模式控件同样包装到一个父 frame 中
    - 确保模式切换时正确显示/隐藏
    - _需求: 7.1, 7.2_

- [-] 4. 实现 CSV 转 Excel 执行逻辑与错误处理
  - [-] 4.1 实现 `_run_csv_to_excel()` 方法
    - 验证四个路径均非空，否则弹出警告对话框
    - 禁用按钮、启动进度条、启动后台线程调用 `_do_csv_to_excel()`
    - _需求: 6.1, 6.2, 5.2_
  - [x] 4.2 实现 `_do_csv_to_excel()` 方法
    - 读取三个 CSV 文件（UTF-8），每步记录日志
    - 调用 `validate_columns()` 验证必需列
    - 调用 `process_csv_data()` 执行数据关联
    - 使用 `pd.ExcelWriter` + openpyxl 写入结果，sheet 名为"意向单详细数据"
    - 成功时日志显示成功消息并弹出成功对话框
    - 异常时日志显示错误并弹出错误对话框
    - finally 中恢复按钮状态和停止进度条
    - _需求: 3.1, 4.1, 4.2, 5.1, 5.3, 5.4_
  - [ ]* 4.3 编写单元测试
    - 使用 `excel_analyzer/sample/` 目录下的示例文件进行端到端测试
    - 验证生成的 Excel 文件包含"意向单详细数据" sheet
    - 验证输出列数为 17 且列顺序正确
    - 验证现有 Excel 对比分析功能不受影响（回归测试）
    - _需求: 4.2, 4.3, 4.4, 7.1_

- [ ] 5. 添加 hypothesis 依赖并整合测试
  - [~] 5.1 在 `excel_analyzer/requirements.txt` 中添加 `hypothesis` 和 `pytest` 依赖
    - _需求: 无（测试基础设施）_
  - [ ]* 5.2 创建 `excel_analyzer/tests/__init__.py` 和测试文件 `excel_analyzer/tests/test_csv_to_excel.py`
    - 将任务 1.3-1.6 的属性测试和任务 4.3 的单元测试整合到此文件中
    - 确保所有测试可通过 `pytest excel_analyzer/tests/` 运行
    - _需求: 无（测试基础设施）_

- [ ] 6. 最终检查点 - 确保所有测试通过
  - 确保所有测试通过，如有问题请向用户确认。

## 备注

- 标记 `*` 的任务为可选任务，可跳过以加快 MVP 进度
- 每个任务引用了具体需求编号以确保可追溯性
- 检查点确保增量验证
- 属性测试验证通用正确性属性，单元测试验证具体示例和边界情况
- 数据处理逻辑提取为纯函数（`process_csv_data`、`validate_columns`），便于独立测试
