# 需求文档

## 简介

修改现有的 `app.py`（基于 tkinter 的桌面应用），新增"CSV 转 Excel"功能模块。用户上传三个 CSV 文件（意向单.csv、csr.csv、会员店251110.csv），系统通过关联查询生成一个包含增强数据的 Excel 结果文件。结果文件包含一个名为"意向单详细数据"的 sheet，其数据来源于意向单.csv，并通过关联 csr.csv 和会员店251110.csv 补充"下INQ渠道"和"是否协议酒店"两列。

## 术语表

- **App**: 基于 tkinter 的桌面应用程序（excel_analyzer/app.py）
- **意向单CSV**: 包含意向单数据的 CSV 文件，列包括 number、省、城市、等级、user_type、user_id、name、客户手机号、supplier_type、supplier_id、酒店ID、酒店名称、上架状态、source、created_at
- **CSR_CSV**: 包含 CSR 订单数据的 CSV 文件，列包括 城市、单号、下单时间、下单用户类型、下单用户id、审核状态、来源平台、来源线索、source_tags 等
- **会员店CSV**: 包含会员店数据的 CSV 文件，列包括 大区、省份、城市、行政区、酒店ID、酒店名称、酒店类型、上架状态、上架时间、是否会员店
- **结果Excel**: 生成的 .xlsx 文件，包含一个名为"意向单详细数据"的 sheet
- **下INQ渠道**: 结果中新增的列，值来源于 CSR_CSV 的"来源线索"列，通过 user_id 关联查询
- **是否协议酒店**: 结果中新增的列，值来源于会员店CSV 的"是否会员店"列，通过酒店ID 关联查询

## 需求

### 需求 1：CSV 转 Excel 功能入口

**用户故事：** 作为用户，我想在应用中选择"CSV 转 Excel"功能模式，以便使用新的 CSV 数据合并生成功能。

#### 验收标准

1. WHEN 用户启动 App，THE App SHALL 提供"CSV 转 Excel 生成"功能模式的选择入口，与现有的"Excel 对比分析"功能并列显示
2. WHEN 用户选择"CSV 转 Excel 生成"模式，THE App SHALL 显示三个 CSV 文件上传区域和一个输出文件路径选择区域
3. WHEN 用户选择"CSV 转 Excel 生成"模式，THE App SHALL 隐藏与"Excel 对比分析"相关的设置控件（分析类型、关联列名）

### 需求 2：CSV 文件上传

**用户故事：** 作为用户，我想分别上传意向单.csv、csr.csv 和会员店251110.csv 三个文件，以便系统进行数据关联处理。

#### 验收标准

1. THE App SHALL 提供三个独立的文件选择控件，分别标注为"意向单 CSV"、"CSR CSV"和"会员店 CSV"
2. WHEN 用户点击文件选择按钮，THE App SHALL 打开文件对话框，默认筛选 CSV 文件类型（*.csv）
3. WHEN 用户选择文件后，THE App SHALL 在对应的输入框中显示所选文件的完整路径
4. THE App SHALL 提供一个输出文件路径选择控件，默认保存格式为 .xlsx

### 需求 3：数据关联与合并处理

**用户故事：** 作为用户，我想让系统自动将意向单数据与 CSR 数据和会员店数据进行关联，以便获得包含"下INQ渠道"和"是否协议酒店"信息的完整数据。

#### 验收标准

1. WHEN 用户点击"开始生成"按钮，THE App SHALL 使用 pandas 读取意向单CSV、CSR_CSV 和会员店CSV 三个文件，编码格式为 UTF-8
2. WHEN 执行关联查询时，THE App SHALL 将意向单CSV 的 user_id 列与 CSR_CSV 的"下单用户id"列统一转换为字符串类型后进行左连接（left join）
3. WHEN 执行关联查询时，THE App SHALL 将意向单CSV 的"酒店ID"列与会员店CSV 的"酒店ID"列统一转换为字符串类型后进行左连接（left join）
4. WHEN 关联完成后，THE App SHALL 将 CSR_CSV 中匹配到的"来源线索"列值填入结果的"下INQ渠道"列
5. WHEN 关联完成后，THE App SHALL 将会员店CSV 中匹配到的"是否会员店"列值填入结果的"是否协议酒店"列
6. WHEN 意向单CSV 中某行的 user_id 在 CSR_CSV 中无匹配记录时，THE App SHALL 将该行的"下INQ渠道"列值设为空
7. WHEN 意向单CSV 中某行的酒店ID 在会员店CSV 中无匹配记录时，THE App SHALL 将该行的"是否协议酒店"列值设为空

### 需求 4：Excel 结果文件生成

**用户故事：** 作为用户，我想将合并后的数据导出为一个格式规范的 Excel 文件，以便后续分析使用。

#### 验收标准

1. WHEN 数据关联处理完成后，THE App SHALL 生成一个 .xlsx 格式的 Excel 文件，写入引擎为 openpyxl
2. THE 结果Excel SHALL 包含一个名为"意向单详细数据"的 sheet
3. THE "意向单详细数据" sheet SHALL 包含意向单CSV 的全部原始列，并在末尾追加"下INQ渠道"和"是否协议酒店"两列
4. THE "意向单详细数据" sheet 中的列顺序 SHALL 为：number、省、城市、等级、user_type、user_id、name、客户手机号、supplier_type、supplier_id、酒店ID、酒店名称、上架状态、source、created_at、下INQ渠道、是否协议酒店

### 需求 5：处理进度与反馈

**用户故事：** 作为用户，我想在数据处理过程中看到进度和状态信息，以便了解处理是否正常进行。

#### 验收标准

1. WHEN 用户点击"开始生成"按钮，THE App SHALL 在日志区域显示每个处理步骤的状态信息（读取文件、关联数据、写入结果）
2. WHILE 数据处理正在进行，THE App SHALL 显示进度条动画，并禁用"开始生成"按钮以防止重复提交
3. WHEN 数据处理成功完成，THE App SHALL 在日志区域显示成功消息，包含输出文件名称，并弹出成功提示对话框
4. IF 处理过程中发生错误（如文件格式不正确、列名缺失），THEN THE App SHALL 在日志区域显示错误信息，并弹出错误提示对话框

### 需求 6：输入验证

**用户故事：** 作为用户，我想在开始处理前得到输入校验反馈，以便避免因文件缺失或格式错误导致处理失败。

#### 验收标准

1. WHEN 用户点击"开始生成"按钮且任一 CSV 文件路径为空，THE App SHALL 弹出警告对话框提示用户选择所有文件
2. WHEN 用户点击"开始生成"按钮且输出文件路径为空，THE App SHALL 弹出警告对话框提示用户选择输出路径
3. WHEN 读取 CSV 文件后，THE App SHALL 验证意向单CSV 包含 user_id 和酒店ID 列，CSR_CSV 包含"下单用户id"和"来源线索"列，会员店CSV 包含"酒店ID"和"是否会员店"列
4. IF 任一 CSV 文件缺少必需的关联列，THEN THE App SHALL 在日志区域显示具体缺失的列名，并弹出错误提示对话框终止处理

### 需求 7：保留现有功能

**用户故事：** 作为用户，我想在使用新功能的同时保留原有的 Excel 对比分析功能，以便根据需要切换使用。

#### 验收标准

1. THE App SHALL 保留现有的"按列合并"、"数据对比"和"相关性分析"三种分析功能，功能逻辑不受修改影响
2. WHEN 用户选择"Excel 对比分析"模式，THE App SHALL 显示原有的两个 Excel 文件选择控件和分析设置控件
