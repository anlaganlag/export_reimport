# 测试文件目录

此目录用于存放用于验证的测试文件样本。

## 所需文件

1. **采购装箱单 (original_packing_list.xlsx)**
   
   这是最主要的输入文件，包含需要处理的采购信息。该文件应包含以下特征：
   - 表头应包含"采购装箱单"或"装箱单"字样
   - 应有明确的装箱单编号
   - 应有中英文字段名（通常在第2-3行）
   - 必须字段：序号、零件号、描述、数量、单位、净重、毛重、单价、总金额

2. **政策文件 (policy.xlsx)**
   
   包含处理采购装箱单所需的政策参数，如汇率、加价率等。该文件应包含：
   - 与装箱单匹配的编号
   - 汇率（保留4位小数）
   - 公司和银行信息
   - 加价率、保险费率等参数

## 文件格式要求

### 采购装箱单格式

采购装箱单应遵循以下格式：
- 第1行：标题和编号
- 第2行：中文字段名
- 第3行：英文字段名
- 第4行及以后：数据行
- 最后部分：汇总行（包含"合计"或"Total"）

### 政策文件格式

政策文件应包含以下内容：
- 装箱单编号
- 汇率设置
- 加价比例
- 保险费率
- 运费设置
- 公司信息
- 银行账户信息

## 示例数据

如果您没有真实数据进行测试，可以使用以下命令创建示例数据：

```bash
python validation_program/create_sample_data.py
```

该脚本将在此目录创建用于测试的示例文件。

## 文件位置

验证程序默认会查找以下文件路径：

```
testfiles/original_packing_list.xlsx
testfiles/policy.xlsx
```

您也可以在运行验证程序时通过参数指定其他文件路径：

```bash
python validation_program/run_validation.py --packing-list path/to/your/packing_list.xlsx --policy-file path/to/your/policy.xlsx
``` 