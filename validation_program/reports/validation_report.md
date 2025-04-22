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
- `cif_original_invoice.xlsx`
- `deduped_list.xlsx`
- `export_invoice - 副本 (2).xlsx`
- `export_invoice - 副本 (3).xlsx`
- `export_invoice - 副本 (4).xlsx`
- `export_invoice - 副本.xlsx`
- `export_invoice.xlsx`
- `output1-export.xlsx`
- `output2-reimport.xlsx`
- `pl_original_invoice.xlsx`
- `reimport_invoice - 副本.xlsx`
- `reimport_invoice.xlsx`
- `reimport_工厂_Daman.xlsx`
- `reimport_工厂_Silvass.xlsx`
- `reimport_麦格米特_Silvass.xlsx`
- `temp_export_invoice.xlsx`
- `xreimport_工厂_Daman.xlsx`
- `xreimport_工厂_Silvass.xlsx`
- `xreimport_麦格米特_Silvass.xlsx`

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

### 输入文件验证: 8/10 通过

### 处理逻辑验证: 6/7 通过

### 输出文件验证: 8/20 通过

## 详细测试结果

### 输入文件验证

#### ❌ 失败的测试

##### project_split: ❌ 失败
- 结果: 以下项目未找到对应的进口发票: SMT工厂设备配件, 组装厂月度辅耗材, TP-LINK, SMT工厂月度辅耗材

##### sheet_naming: ❌ 失败
- 结果: 进口文件reimport_工厂_Daman.xlsx中缺少'PL'工作表
进口文件reimport_工厂_Daman.xlsx中的工作表'Sheet1'不符合发票号码命名格式
进口文件reimport_工厂_Silvass.xlsx中缺少'PL'工作表
进口文件reimport_工厂_Silvass.xlsx中的工作表'Sheet1'不符合发票号码命名格式
进口文件reimport_麦格米特_Silvass.xlsx中缺少'PL'工作表
进口文件reimport_麦格米特_Silvass.xlsx中的工作表'Sheet1'不符合发票号码命名格式

#### ✅ 通过的测试

##### company_bank_info: ✅ 通过
- 结果: 公司信息和银行信息验证通过

##### exchange_rate_decimal: ✅ 通过
- 结果: 汇率小数位验证通过

##### factory_split: ✅ 通过
- 结果: 所有工厂都有对应的进口发票文件

##### packing_list_field_headers: ✅ 通过
- 结果: 表头字段名验证通过。中文字段: 27个, 英文字段: 27个

##### packing_list_header: ✅ 通过
- 结果: 采购装箱单表头验证通过。找到标题: '采购装箱单', 编号: 'PL25001'

##### policy_file_id: ✅ 通过
- 结果: 政策文件编号验证通过

##### summary_data: ✅ 通过
- 结果: 汇总数据验证通过

##### weights: ✅ 通过
- 结果: 净重毛重验证通过

**统计**: 8/10 测试通过

### 处理逻辑验证

#### ❌ 失败的测试

##### merge_logic: ❌ 失败
- 结果: 未找到所有需要的列，无法验证合并逻辑

#### ✅ 通过的测试

##### cif_price_calculation: ✅ 通过
- 结果: CIF价格计算验证通过

##### fob_price_calculation: ✅ 通过
- 结果: FOB价格计算验证通过

##### freight_calculation: ✅ 通过
- 结果: 运费计算验证通过

##### insurance_calculation: ✅ 通过
- 结果: 保险费计算验证通过

##### trade_type_identification: ✅ 通过
- 结果: 贸易类型识别逻辑验证通过

##### trade_type_split: ✅ 通过
- 结果: 贸易类型拆分验证通过

**统计**: 6/7 测试通过

### 输出文件验证

#### ❌ 失败的测试

##### export_invoice_field_mapping: ❌ 失败
- 结果: 输出文件缺少字段: NO., Material code, DESCRIPTION, Model NO., Unit Price, Qty, Amount

##### export_invoice_quantity: ❌ 失败
- 结果: 出口发票总数量(19006.0)与采购装箱单总数量(9503.0)不一致
存在差异的物料: C100.C05-032-04-00(原始:497.0, 出口:0, 差异:-497.0), E100.020310008(原始:6.0, 出口:0, 差异:-6.0), E100.020310014(原始:2.0, 出口:0, 差异:-2.0), E100.020310015(原始:1.0, 出口:0, 差异:-1.0), E100.A37-066-02-00(原始:100.0, 出口:0, 差异:-100.0), J100.020715018(原始:137.0, 出口:0, 差异:-137.0), E100.020200009(原始:1.0, 出口:0, 差异:-1.0), E100.020310017(原始:1.0, 出口:0, 差异:-1.0), E100.020310012(原始:3.0, 出口:0, 差异:-3.0), E100.A33-013-03-00(原始:5.0, 出口:0, 差异:-5.0), J100.S07-010-10-00(原始:30.0, 出口:0, 差异:-30.0), J100.S07-010-04-01(原始:500.0, 出口:0, 差异:-500.0), J100.S07-010-06-01(原始:400.0, 出口:0, 差异:-400.0), J100.S07-010-06-02(原始:40.0, 出口:0, 差异:-40.0), J100.S07-010-11-00(原始:40.0, 出口:0, 差异:-40.0), C100.C06-007-01-00(原始:140.0, 出口:0, 差异:-140.0), C100.C06-019-06-00(原始:250.0, 出口:0, 差异:-250.0), E100.A20-001-15-00(原始:388.0, 出口:0, 差异:-388.0), E100.A20-001-16-00(原始:457.0, 出口:0, 差异:-457.0), E100.A20-001-17-00(原始:458.0, 出口:0, 差异:-458.0), E100.A20-001-20-00(原始:287.0, 出口:0, 差异:-287.0), E100.A20-001-52-00(原始:40.0, 出口:0, 差异:-40.0), E100.A20-001-53-00(原始:60.0, 出口:0, 差异:-60.0), E100.A20-001-54-00(原始:50.0, 出口:0, 差异:-50.0), E100.A20-001-22-00(原始:50.0, 出口:0, 差异:-50.0), E100.A20-001-33-00(原始:110.0, 出口:0, 差异:-110.0), C100.C06-006-01-00(原始:2040.0, 出口:0, 差异:-2040.0), C100.C06-006-02-00(原始:2800.0, 出口:0, 差异:-2800.0), E100.020349014(原始:2.0, 出口:0, 差异:-2.0), E100.A37-154-01-00(原始:5.0, 出口:0, 差异:-5.0), E100.0111901002(原始:3.0, 出口:0, 差异:-3.0), J100.031003005(原始:12.0, 出口:0, 差异:-12.0), E100.021335001(原始:9.0, 出口:0, 差异:-9.0), E100.E17-003-01-00(原始:10.0, 出口:0, 差异:-10.0), E100.E17-004-01-00(原始:10.0, 出口:0, 差异:-10.0), E100.020396013(原始:1.0, 出口:0, 差异:-1.0), E100.0203104001(原始:10.0, 出口:0, 差异:-10.0), E100.0203154000(原始:5.0, 出口:0, 差异:-5.0), E100.0203125000(原始:500.0, 出口:0, 差异:-500.0), E100.0203162159(原始:4.0, 出口:0, 差异:-4.0), E100.0203133001(原始:5.0, 出口:0, 差异:-5.0), E100.E00-011-15-01(原始:2.0, 出口:0, 差异:-2.0), E100.020396061(原始:10.0, 出口:0, 差异:-10.0), E100.020396062(原始:10.0, 出口:0, 差异:-10.0), E100.020396047(原始:2.0, 出口:0, 差异:-2.0), E100.020396048(原始:2.0, 出口:0, 差异:-2.0), E100.020396049(原始:2.0, 出口:0, 差异:-2.0), E100.020396050(原始:2.0, 出口:0, 差异:-2.0), E100.020396055(原始:2.0, 出口:0, 差异:-2.0), E100.020396056(原始:2.0, 出口:0, 差异:-2.0)

##### import_invoice_field_mapping_reimport_invoice.xlsx: ❌ 失败
- 结果: 输出文件缺少字段: NO., Material code, DESCRIPTION, Model NO., Unit Price, Qty, Amount

##### import_invoice_field_mapping_reimport_工厂_Daman.xlsx: ❌ 失败
- 结果: 验证字段映射时出错: Worksheet index 1 is invalid, 1 worksheets found

##### import_invoice_field_mapping_reimport_工厂_Silvass.xlsx: ❌ 失败
- 结果: 验证字段映射时出错: Worksheet index 1 is invalid, 1 worksheets found

##### import_invoice_field_mapping_reimport_麦格米特_Silvass.xlsx: ❌ 失败
- 结果: 验证字段映射时出错: Worksheet index 1 is invalid, 1 worksheets found

##### import_invoice_quantity: ❌ 失败
- 结果: 进口发票总数量(0)与采购装箱单总数量(9503.0)不一致
进口发票数量明细: 

##### import_invoice_split: ❌ 失败
- 结果: 以下项目和工厂组合未找到对应的进口发票: 项目:TP-LINK 工厂:Silvass, 项目:TP-LINK 工厂:Daman, 项目:组装厂月度辅耗材 工厂:Daman, 项目:SMT工厂月度辅耗材 工厂:Silvass, 项目:SMT工厂设备配件 工厂:Silvass

##### import_packing_list_field_mapping_reimport_工厂_Daman.xlsx: ❌ 失败
- 结果: 无法读取输出文件: Worksheet named 'PL' not found

##### import_packing_list_field_mapping_reimport_工厂_Silvass.xlsx: ❌ 失败
- 结果: 无法读取输出文件: Worksheet named 'PL' not found

##### import_packing_list_field_mapping_reimport_麦格米特_Silvass.xlsx: ❌ 失败
- 结果: 无法读取输出文件: Worksheet named 'PL' not found

##### import_packing_list_totals: ❌ 失败
- 结果: 进口装箱单汇总数据不一致: Quantity不一致: 进口总计(38018) vs 原始(19006.0); Total Carton Quantity不一致: 进口总计(0) vs 原始(335.988); Total Volume不一致: 进口总计(0) vs 原始(36.63988); Total Gross Weight不一致: 进口总计(1152.6399999999999) vs 原始(6749.6464); Total Net Weight不一致: 进口总计(12228.999999999998) vs 原始(6313.32)

#### ✅ 通过的测试

##### export_invoice_prices: ✅ 通过
- 结果: 出口发票价格验证通过

##### export_packing_list_field_mapping: ✅ 通过
- 结果: 字段映射验证通过

##### export_packing_list_totals: ✅ 通过
- 结果: 出口装箱单汇总数据验证通过

##### file_naming: ✅ 通过
- 结果: 文件命名格式验证通过

##### format_compliance: ✅ 通过
- 结果: 文件格式一致性验证通过

##### import_invoice_field_mapping_reimport_invoice - 副本.xlsx: ✅ 通过
- 结果: 字段映射验证通过

##### import_packing_list_field_mapping_reimport_invoice - 副本.xlsx: ✅ 通过
- 结果: 字段映射验证通过

##### import_packing_list_field_mapping_reimport_invoice.xlsx: ✅ 通过
- 结果: 字段映射验证通过

**统计**: 8/20 测试通过

## 解决方案建议

### 输入文件验证问题修复建议

#### sheet_naming:
- **建议**: 根据错误信息修复问题: 进口文件reimport_工厂_Daman.xlsx中缺少'PL'工作表
进口文件reimport_工厂_Daman.xlsx中的工作表'Sheet1'不符合发票号码命名格式
进口文件reimport_工厂_Silvass.xlsx中缺少'PL'工作表
进口文件reimport_工厂_Silvass.xlsx中的工作表'Sheet1'不符合发票号码命名格式
进口文件reimport_麦格米特_Silvass.xlsx中缺少'PL'工作表
进口文件reimport_麦格米特_Silvass.xlsx中的工作表'Sheet1'不符合发票号码命名格式

#### project_split:
- **建议**: 根据错误信息修复问题: 以下项目未找到对应的进口发票: SMT工厂设备配件, 组装厂月度辅耗材, TP-LINK, SMT工厂月度辅耗材

### 处理逻辑验证问题修复建议

#### merge_logic:
- **建议**: 根据错误信息修复问题: 未找到所有需要的列，无法验证合并逻辑

### 输出文件验证问题修复建议

#### export_invoice_field_mapping:
- **建议**: 根据错误信息修复问题: 输出文件缺少字段: NO., Material code, DESCRIPTION, Model NO., Unit Price, Qty, Amount

#### export_invoice_quantity:
- **建议**: 根据错误信息修复问题: 出口发票总数量(19006.0)与采购装箱单总数量(9503.0)不一致
存在差异的物料: C100.C05-032-04-00(原始:497.0, 出口:0, 差异:-497.0), E100.020310008(原始:6.0, 出口:0, 差异:-6.0), E100.020310014(原始:2.0, 出口:0, 差异:-2.0), E100.020310015(原始:1.0, 出口:0, 差异:-1.0), E100.A37-066-02-00(原始:100.0, 出口:0, 差异:-100.0), J100.020715018(原始:137.0, 出口:0, 差异:-137.0), E100.020200009(原始:1.0, 出口:0, 差异:-1.0), E100.020310017(原始:1.0, 出口:0, 差异:-1.0), E100.020310012(原始:3.0, 出口:0, 差异:-3.0), E100.A33-013-03-00(原始:5.0, 出口:0, 差异:-5.0), J100.S07-010-10-00(原始:30.0, 出口:0, 差异:-30.0), J100.S07-010-04-01(原始:500.0, 出口:0, 差异:-500.0), J100.S07-010-06-01(原始:400.0, 出口:0, 差异:-400.0), J100.S07-010-06-02(原始:40.0, 出口:0, 差异:-40.0), J100.S07-010-11-00(原始:40.0, 出口:0, 差异:-40.0), C100.C06-007-01-00(原始:140.0, 出口:0, 差异:-140.0), C100.C06-019-06-00(原始:250.0, 出口:0, 差异:-250.0), E100.A20-001-15-00(原始:388.0, 出口:0, 差异:-388.0), E100.A20-001-16-00(原始:457.0, 出口:0, 差异:-457.0), E100.A20-001-17-00(原始:458.0, 出口:0, 差异:-458.0), E100.A20-001-20-00(原始:287.0, 出口:0, 差异:-287.0), E100.A20-001-52-00(原始:40.0, 出口:0, 差异:-40.0), E100.A20-001-53-00(原始:60.0, 出口:0, 差异:-60.0), E100.A20-001-54-00(原始:50.0, 出口:0, 差异:-50.0), E100.A20-001-22-00(原始:50.0, 出口:0, 差异:-50.0), E100.A20-001-33-00(原始:110.0, 出口:0, 差异:-110.0), C100.C06-006-01-00(原始:2040.0, 出口:0, 差异:-2040.0), C100.C06-006-02-00(原始:2800.0, 出口:0, 差异:-2800.0), E100.020349014(原始:2.0, 出口:0, 差异:-2.0), E100.A37-154-01-00(原始:5.0, 出口:0, 差异:-5.0), E100.0111901002(原始:3.0, 出口:0, 差异:-3.0), J100.031003005(原始:12.0, 出口:0, 差异:-12.0), E100.021335001(原始:9.0, 出口:0, 差异:-9.0), E100.E17-003-01-00(原始:10.0, 出口:0, 差异:-10.0), E100.E17-004-01-00(原始:10.0, 出口:0, 差异:-10.0), E100.020396013(原始:1.0, 出口:0, 差异:-1.0), E100.0203104001(原始:10.0, 出口:0, 差异:-10.0), E100.0203154000(原始:5.0, 出口:0, 差异:-5.0), E100.0203125000(原始:500.0, 出口:0, 差异:-500.0), E100.0203162159(原始:4.0, 出口:0, 差异:-4.0), E100.0203133001(原始:5.0, 出口:0, 差异:-5.0), E100.E00-011-15-01(原始:2.0, 出口:0, 差异:-2.0), E100.020396061(原始:10.0, 出口:0, 差异:-10.0), E100.020396062(原始:10.0, 出口:0, 差异:-10.0), E100.020396047(原始:2.0, 出口:0, 差异:-2.0), E100.020396048(原始:2.0, 出口:0, 差异:-2.0), E100.020396049(原始:2.0, 出口:0, 差异:-2.0), E100.020396050(原始:2.0, 出口:0, 差异:-2.0), E100.020396055(原始:2.0, 出口:0, 差异:-2.0), E100.020396056(原始:2.0, 出口:0, 差异:-2.0)

#### import_invoice_split:
- **建议**: 根据错误信息修复问题: 以下项目和工厂组合未找到对应的进口发票: 项目:TP-LINK 工厂:Silvass, 项目:TP-LINK 工厂:Daman, 项目:组装厂月度辅耗材 工厂:Daman, 项目:SMT工厂月度辅耗材 工厂:Silvass, 项目:SMT工厂设备配件 工厂:Silvass

#### import_invoice_field_mapping_reimport_invoice.xlsx:
- **建议**: 根据错误信息修复问题: 输出文件缺少字段: NO., Material code, DESCRIPTION, Model NO., Unit Price, Qty, Amount

#### import_invoice_field_mapping_reimport_工厂_Daman.xlsx:
- **建议**: 根据错误信息修复问题: 验证字段映射时出错: Worksheet index 1 is invalid, 1 worksheets found

#### import_packing_list_field_mapping_reimport_工厂_Daman.xlsx:
- **建议**: 根据错误信息修复问题: 无法读取输出文件: Worksheet named 'PL' not found

#### import_invoice_field_mapping_reimport_工厂_Silvass.xlsx:
- **建议**: 根据错误信息修复问题: 验证字段映射时出错: Worksheet index 1 is invalid, 1 worksheets found

#### import_packing_list_field_mapping_reimport_工厂_Silvass.xlsx:
- **建议**: 根据错误信息修复问题: 无法读取输出文件: Worksheet named 'PL' not found

#### import_invoice_field_mapping_reimport_麦格米特_Silvass.xlsx:
- **建议**: 根据错误信息修复问题: 验证字段映射时出错: Worksheet index 1 is invalid, 1 worksheets found

#### import_packing_list_field_mapping_reimport_麦格米特_Silvass.xlsx:
- **建议**: 根据错误信息修复问题: 无法读取输出文件: Worksheet named 'PL' not found

#### import_invoice_quantity:
- **建议**: 根据错误信息修复问题: 进口发票总数量(0)与采购装箱单总数量(9503.0)不一致
进口发票数量明细: 

#### import_packing_list_totals:
- **建议**: 根据错误信息修复问题: 进口装箱单汇总数据不一致: Quantity不一致: 进口总计(38018) vs 原始(19006.0); Total Carton Quantity不一致: 进口总计(0) vs 原始(335.988); Total Volume不一致: 进口总计(0) vs 原始(36.63988); Total Gross Weight不一致: 进口总计(1152.6399999999999) vs 原始(6749.6464); Total Net Weight不一致: 进口总计(12228.999999999998) vs 原始(6313.32)



---
生成时间: 2025-04-22 16:10:20