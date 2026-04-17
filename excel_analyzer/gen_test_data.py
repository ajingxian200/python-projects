"""生成测试用 Excel 文件"""
import pandas as pd

df1 = pd.DataFrame({
    "员工ID": [1, 2, 3, 4, 5],
    "姓名": ["张三", "李四", "王五", "赵六", "钱七"],
    "部门": ["技术", "销售", "技术", "人事", "销售"],
    "薪资": [15000, 12000, 18000, 11000, 13000],
})

df2 = pd.DataFrame({
    "员工ID": [1, 2, 3, 6, 7],
    "姓名": ["张三", "李四", "王五", "孙八", "周九"],
    "部门": ["技术", "市场", "技术", "财务", "技术"],  # 李四部门不同
    "薪资": [16000, 12000, 19000, 14000, 15000],  # 张三、王五薪资不同
})

df1.to_excel("test_file1.xlsx", index=False)
df2.to_excel("test_file2.xlsx", index=False)
print("测试文件已生成: test_file1.xlsx, test_file2.xlsx")
