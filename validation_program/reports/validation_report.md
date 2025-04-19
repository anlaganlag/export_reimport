# 进出口文件生成系统验收报告

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
- `cif_original_invoice.xlsx`
- `combined_invoice.xlsx`
- `deduped_list.xlsx`
- `export_invoice.xlsx`
- `invoice_大华_Silvassa.xlsx`
- `invoice_工厂_Daman.xlsx`
- `invoice_工厂_Silvassa.xlsx`
- `invoice_麦格米特_Silvassa - 副本.xlsx`
- `invoice_麦格米特_Silvassa.xlsx`
- `output1-export.xlsx`
- `output2-reimport.xlsx`
- `pf.xlsx`
- `ph.xlsx`
- `pl_original_invoice.xlsx`
- `reimport_invoice.xlsx`
- `~$export_invoice.xlsx`
- `~$reimport_invoice.xlsx`

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

### 输入文件验证: 3/3 通过

## 详细测试结果

### 输入文件验证

#### packing_list_field_headers: ✅ 通过
- 结果: 表头字段名验证通过。中文字段: 27个, 英文字段: 27个

#### packing_list_header: ✅ 通过
- 结果: 采购装箱单表头验证通过。找到标题: '采购装箱单', 编号: 'PL25001'

#### policy_file_id: ✅ 通过
- 结果: 政策文件编号验证通过

## 解决方案建议



---
生成时间: 2025-04-19 23:37:47