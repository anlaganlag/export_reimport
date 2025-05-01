# 进出口文件生成系统验收报告

## 验证说明

本报告包含对进出口文件生成系统的完整验证结果。验证过程包括以下几个部分：

1. **输入文件验证** - 检查装箱单和政策文件的格式和内容是否符合要求
2. **处理逻辑验证** - 验证文件处理逻辑是否正确
3. **输出文件验证** - 验证生成的输出文件是否符合要求

如果任何验证项目失败，整体验收结果将被标记为"不通过"。请查看详细测试结果部分以找出具体问题。

## 验证文件信息

### 输入文件
- 采购装箱单: `D:\project\export_reimport\testfiles\original_packing_list.xlsx`
- 政策文件: `D:\project\export_reimport\testfiles\policy.xlsx`

### 输出目录
- 输出路径: `D:\project\export_reimport\outputs`
- 模板目录: `D:\project\export_reimport\Template`
- 报告路径: `D:\project\export_reimport\validation_program\reports\validation_report.md`

### 处理参数
- 跳过文件处理: 否

### 检测到的输出文件
- `backup_reimport_invoice.xlsx`
- `cif_original_invoice.xlsx`
- `export_invoice.xlsx`
- `pl_original_invoice.xlsx`
- `reimport_invoice.xlsx`
- `reimport_工厂_1.0.xlsx`
- `reimport_工厂_10.0.xlsx`
- `reimport_工厂_11.0.xlsx`
- `reimport_工厂_12.0.xlsx`
- `reimport_工厂_13.0.xlsx`
- `reimport_工厂_14.0.xlsx`
- `reimport_工厂_15.0.xlsx`
- `reimport_工厂_16.0.xlsx`
- `reimport_工厂_17.0.xlsx`
- `reimport_工厂_18.0.xlsx`
- `reimport_工厂_19.0.xlsx`
- `reimport_工厂_2.0.xlsx`
- `reimport_工厂_20.0.xlsx`
- `reimport_工厂_21.0.xlsx`
- `reimport_工厂_3.0.xlsx`
- `reimport_工厂_4.0.xlsx`
- `reimport_工厂_5.0.xlsx`
- `reimport_工厂_6.0.xlsx`
- `reimport_工厂_7.0.xlsx`
- `reimport_工厂_8.0.xlsx`
- `reimport_工厂_9.0.xlsx`
- `reimport_工厂_AA.xlsx`
- `reimport_工厂_Daman.xlsx`
- `reimport_工厂_SS.xlsx`
- `reimport_工厂_Silvassa.xlsx`

## 验收标准

### 输入文件标准

1. **采购装箱单标题**
   - 要求: 文件应包含'采购装箱单'或'装箱单'等相关标题文本
   - 格式示例: '采购装箱单 PL-20250418-0001'
   - 参考文档: testfiles/README.md 第12行

2. **字段头格式**
   - 要求: 表头应包含中英文字段名，含必要字段
   - 必要中文字段: 序号、零件号、描述、数量、单位、净重、毛重
   - 必要英文字段: No、Part、Description、Quantity、Unit、Net、Gross
   - 参考文档: testfiles/README.md 第13-14行

3. **政策文件要求**
   - 要求: 政策文件应包含与装箱单匹配的编号，以及完整的参数设置
   - 必要内容: 匹配编号、汇率、加价率、保险费率、公司和银行信息
   - 参考文档: testfiles/README.md 第21-25行

## 验收结果：不通过

## 验证结果统计

### 输入文件验证: 3/7 通过

## 详细测试结果

### 输入文件验证

#### ❌ 失败的测试

##### packing_list_field_headers: ❌ 失败
- 结果: 缺少必要的中文字段: 采购单价不含税。找到的字段: 序号, 料号, 供应商, 项目名称, 工厂地点...。验收标准: 表头应包含所有必要的中文字段(参见文档第14行要求)。

##### packing_list_header: ❌ 失败
- 结果: 表头未包含任何所需标题文本。需要: 采购装箱单，实际值: 'nan...'。验收标准: 文件第1行应包含装箱单或相关标题文本(参见文档第12行要求)。

##### policy_file_id: ❌ 失败
- 结果: 采购装箱单编号未提取到。验收标准: 采购装箱单必须包含可识别的编号。

##### summary_data: ❌ 失败
- 结果: 汇总数据有误: 数量汇总不正确: 显示1, 实际35.0

#### ✅ 通过的测试

##### company_bank_info: ✅ 通过
- 结果: 公司信息和银行信息验证通过

##### exchange_rate_decimal: ✅ 通过
- 结果: 汇率验证通过

##### weights: ✅ 通过
- 结果: 净重毛重验证通过

**统计**: 3/7 测试通过

## 解决方案建议

### 输入文件验证问题修复建议

#### packing_list_header:
- **建议**: 修改文件首行，确保包含'采购装箱单'、'装箱单'或'PACKING LIST'等关键词
- **正确示例**: '采购装箱单 PL-20250418-0001'

#### packing_list_field_headers:
- **建议**: 检查文件表头行，确保包含所有必要的中英文字段
- **中文字段**: 序号、零件号、描述、数量、单位、净重、毛重
- **英文字段**: No、Part No.、Description、Quantity、Unit、Net Weight、Gross Weight

#### summary_data:
- **建议**: 检查表底汇总行的计算，确保数量、体积、净重和毛重的总和正确
- **解决方法**: 重新计算各列的总和，并更新汇总行的值

#### policy_file_id:
- **建议**: 根据错误信息修复问题: 采购装箱单编号未提取到。验收标准: 采购装箱单必须包含可识别的编号。



---
生成时间: 2025-05-01 23:57:11