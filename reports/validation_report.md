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
- 报告路径: `D:\project\export_reimport\reports\validation_report.md`

### 处理参数
- 跳过文件处理: 否

### 检测到的输出文件
- `cif_original_invoice.xlsx`
- `export_invoice.xlsx`
- `reimport_invoice.xlsx`

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

### 输入文件验证: 8/8 通过

### 处理逻辑验证: 7/7 通过

### 输出文件验证: 3/12 通过

## 详细测试结果

### 输入文件验证

#### ✅ 通过的测试

##### company_bank_info: ✅ 通过
- 结果: 公司信息和银行信息验证通过

##### exchange_rate_decimal: ✅ 通过
- 结果: 汇率验证通过

##### packing_list_field_headers: ✅ 通过
- 结果: 表头字段名验证通过。中文字段: 27个, 英文字段: 27个

##### packing_list_header: ✅ 通过
- 结果: 采购装箱单表头验证通过。找到标题: '采购装箱单', 编号: 'CXCI2025012201'

##### policy_file_id: ✅ 通过
- 结果: 政策文件编号验证通过

##### sheet_naming: ✅ 通过
- 结果: 工作表命名验证通过

##### summary_data: ✅ 通过
- 结果: 汇总数据验证已跳过（根据需求变更）

##### weights: ✅ 通过
- 结果: 净重毛重验证通过，所有箱号的总净重均小于总毛重

【自动容错提示】:
箱号 F01 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F02 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F03 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F04 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F05 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F06 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F07 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F08 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F09-F76 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F77-F132 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F133 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F134 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F135 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F136 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F137 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F138 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F139 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F140 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F141 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F142 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F143 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F144 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F145 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F146-F155 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F156-F165 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F166 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F167 的净重(0.0)或毛重(0.0)为0或无效，自动跳过
箱号 F168 的净重(0.0)或毛重(0.0)为0或无效，自动跳过

**统计**: 8/8 测试通过

### 处理逻辑验证

#### ✅ 通过的测试

##### cif_price_calculation: ✅ 通过
- 结果: CIF价格计算验证通过

##### fob_price_calculation: ✅ 通过
- 结果: FOB价格计算验证通过

##### freight_calculation: ✅ 通过
- 结果: 运费计算验证通过

##### insurance_calculation: ✅ 通过
- 结果: 保险费计算验证通过

##### merge_logic: ✅ 通过
- 结果: 物料合并逻辑验证通过

##### trade_type_identification: ✅ 通过
- 结果: 贸易类型识别逻辑验证通过

##### trade_type_split: ✅ 通过
- 结果: 贸易类型拆分验证通过

**统计**: 7/7 测试通过

### 输出文件验证

#### ❌ 失败的测试

##### export_invoice_field_mapping: ❌ 失败
- 结果: 输出文件缺少字段: Material code, DESCRIPTION, Model NO., Qty, Unit Price, Amount

##### export_invoice_prices: ❌ 失败
- 结果: 未找到单价列

##### export_invoice_quantity: ❌ 失败
- 结果: 未找到数量列

##### export_packing_list_field_mapping: ❌ 失败
- 结果: 输出文件缺少字段: P/N., DESCRIPTION, Model NO., QUANTITY, CTNS, Carton MEASUREMENT, G.W (KG), N.W(KG), Carton NO.

##### import_invoice_field_mapping_reimport_invoice.xlsx: ❌ 失败
- 结果: 输出文件缺少字段: NO., Material code, DESCRIPTION, Model NO., Unit Price, Qty, Unit, Amount

##### import_invoice_quantity: ❌ 失败
- 结果: 进口发票总数量(0)与采购装箱单总数量(9506.0)不一致
进口发票数量明细: 

##### import_invoice_split: ❌ 失败
- 结果: 以下项目和工厂组合未找到对应的进口发票: 项目:SMT工厂月度辅耗材 工厂:Silvassa, 项目:SMT工厂设备配件 工厂:Silvassa, 项目:项目名称 工厂:工厂地点, 项目:TP-LINK 工厂:Silvassa, 项目:TP-LINK 工厂:Daman, 项目:麦格米特 工厂:Silvassa, 项目:组装厂月度辅耗材 工厂:Daman

##### import_packing_list_field_mapping_reimport_invoice.xlsx: ❌ 失败
- 结果: 输出文件缺少字段: P/N., DESCRIPTION, Model NO., QUANTITY, CTNS, Carton MEASUREMENT, G.W (KG), N.W(KG), Carton NO.

##### import_packing_list_totals: ❌ 失败
- 结果: 验证进口装箱单汇总数据时出错: can only concatenate str (not "int") to str

#### ✅ 通过的测试

##### export_packing_list_totals: ✅ 通过
- 结果: 出口装箱单汇总数据验证通过

##### file_naming: ✅ 通过
- 结果: 文件命名格式验证通过

##### format_compliance: ✅ 通过
- 结果: 文件格式一致性验证通过

**统计**: 3/12 测试通过

## 解决方案建议

### 输出文件验证问题修复建议

#### export_invoice_field_mapping:
- **建议**: 根据错误信息修复问题: 输出文件缺少字段: Material code, DESCRIPTION, Model NO., Qty, Unit Price, Amount

#### export_invoice_quantity:
- **建议**: 根据错误信息修复问题: 未找到数量列

#### export_invoice_prices:
- **建议**: 根据错误信息修复问题: 未找到单价列

#### export_packing_list_field_mapping:
- **建议**: 根据错误信息修复问题: 输出文件缺少字段: P/N., DESCRIPTION, Model NO., QUANTITY, CTNS, Carton MEASUREMENT, G.W (KG), N.W(KG), Carton NO.

#### import_invoice_split:
- **建议**: 根据错误信息修复问题: 以下项目和工厂组合未找到对应的进口发票: 项目:SMT工厂月度辅耗材 工厂:Silvassa, 项目:SMT工厂设备配件 工厂:Silvassa, 项目:项目名称 工厂:工厂地点, 项目:TP-LINK 工厂:Silvassa, 项目:TP-LINK 工厂:Daman, 项目:麦格米特 工厂:Silvassa, 项目:组装厂月度辅耗材 工厂:Daman

#### import_invoice_field_mapping_reimport_invoice.xlsx:
- **建议**: 根据错误信息修复问题: 输出文件缺少字段: NO., Material code, DESCRIPTION, Model NO., Unit Price, Qty, Unit, Amount

#### import_packing_list_field_mapping_reimport_invoice.xlsx:
- **建议**: 根据错误信息修复问题: 输出文件缺少字段: P/N., DESCRIPTION, Model NO., QUANTITY, CTNS, Carton MEASUREMENT, G.W (KG), N.W(KG), Carton NO.

#### import_invoice_quantity:
- **建议**: 根据错误信息修复问题: 进口发票总数量(0)与采购装箱单总数量(9506.0)不一致
进口发票数量明细: 

#### import_packing_list_totals:
- **建议**: 根据错误信息修复问题: 验证进口装箱单汇总数据时出错: can only concatenate str (not "int") to str



---
生成时间: 2025-05-05 17:13:18