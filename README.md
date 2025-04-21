# 解决GitHub账号选择问题

如果您在使用`git push`时总是看到账号选择对话框，可以按以下步骤解决：

1. 运行`delete_github_credentials.ps1`脚本清除现有的GitHub凭据:
   ```
   .\delete_github_credentials.ps1
   ```

2. 运行`fix_github_account.ps1`脚本设置Git配置:
   ```
   .\fix_github_account.ps1
   ```

3. 按照脚本提示完成操作，下次推送代码时应该不再出现账号选择对话框

如果问题仍然存在，可以尝试以下手动方法：
- 打开`控制面板` -> `用户账户` -> `凭据管理器`
- 找到并删除所有与GitHub相关的Windows凭据
- 在PowerShell中运行以下命令：
  ```
  git config --global credential.helper store
  git config --global --unset credential.helper manager
  git config --global --unset credential.manager
  ```
- 然后尝试`git push`，输入您想默认使用的账号和密码

---

# 装运清单处理工具

## 简介
这个工具用于处理装运清单数据，计算FOB和CIF价格，并生成出口和复进口发票。根据原始装箱单和政策参数，自动计算FOB价格、运费、保险费以及最终的CIF价格。

## 功能特点
- 读取原始装箱单和政策参数Excel文件
- 自动计算FOB价格 (基于加价率)
- 计算保险费 (基于保险系数和保险费率)
- 计算运费 (基于总运费和物料重量)
- 计算最终CIF价格
- 支持按工厂拆分生成复进口发票
- 美观的Excel格式，包括表头样式、数字格式和合计行

## 使用方法

### 准备文件
1. 将原始装箱单Excel文件放在`testfiles`目录下，命名为`original_packing_list.xlsx`
2. 将政策参数Excel文件放在`testfiles`目录下，命名为`policy.xlsx`

### 政策文件格式
政策文件需要包含以下参数:
- `加价率`: FOB计算的加价百分比
- `保险系数`: 保险费计算的系数
- `保险费率`: 保险费率
- `总运费(RMB)`: 总运费金额（人民币）
- `汇率(RMB/美元)`: 人民币兑美元汇率

### 运行程序
1. 确保已安装Python和必要的依赖包：
   ```
   pip install pandas openpyxl
   ```
2. 运行处理脚本：
   ```
   python process_shipping_list.py
   ```
3. 检查`outputs`目录中生成的文件：
   - `export_invoice.xlsx`: 完整的出口发票
   - `reimport_invoice_factory_*.xlsx`: 按工厂拆分的复进口发票

## 输出文件说明
生成的Excel文件包含以下主要列：
- 物料基本信息（料号、描述、型号等）
- 采购价格信息（采购单价、采购总价）
- FOB价格信息（FOB单价、FOB总价）
- 运费和保险费计算（每公斤摊的运保费、该项对应的运保费）
- CIF价格信息（CIF总价、CIF单价）
- 美元价格（单价USD数值）
- 其他辅助信息（工厂、用途等）

装箱单包含以下主要列：
- 序号、料号、描述、型号
- 数量、件数、箱体尺寸
- 毛重、净重、箱号

## 工作流程图
详细的工作流程请参考`WORKFLOW-CN.md`文件，其中包含了详细的流程图和计算逻辑说明。

## 常见问题
1. **如果Excel文件编码有问题怎么办？**  
   确保Excel文件使用UTF-8编码保存，或者在程序中添加相应的编码参数。

2. **如何修改列的顺序？**  
   在`process_shipping_list.py`文件中的`output_columns`列表中调整列的顺序。

3. **如何自定义Excel样式？**  
   修改`apply_excel_styling`函数中的样式设置，如字体、颜色、对齐方式等。

# 出口转内销验收程序

这个程序用于验证出口转内销文件生成系统的输入和输出文件是否符合要求，并生成详细的验收报告。

## 功能特点

- 完整验证输入文件格式和内容
- 验证处理逻辑的正确性
- 验证输出文件的格式和内容
- 生成详细的验收报告
- 支持多种输出格式
- 可自定义验证规则
- 支持生成HTML、Excel、JSON等格式的报告

## 使用方法

### 1. 基本用法

最简单的使用方法是运行示例脚本：

```bash
python validation_program/run_validation.py
```

这将使用默认路径查找输入文件并执行验证。

### 2. 指定文件路径

可以通过命令行参数指定具体的输入和输出文件路径：

```bash
python validation_program/run_validation.py --packing-list path/to/packing_list.xlsx --policy-file path/to/policy.xlsx --output-dir path/to/outputs
```

### 3. 完整参数说明

```
--packing-list    指定采购装箱单文件路径
--policy-file     指定政策文件路径
--output-dir      指定输出目录
--template-dir    指定模板目录
--report-path     指定报告输出路径
--skip-processing 跳过文件处理，仅验证已生成的文件
```

### 4. 直接使用主程序

也可以直接使用主程序，完全自定义验证选项：

```bash
python validation_program/main.py --packing-list path/to/packing_list.xlsx --policy-file path/to/policy.xlsx --output-dir path/to/outputs --template-dir path/to/templates --report-path path/to/report.md
```

## 配置文件说明

程序使用以下配置文件来控制验证行为：

- `validation_program/config/validation_rules.json` - 验证规则配置
- `validation_program/config/integration_rules.json` - 系统集成配置
- `validation_program/config/file_paths.json` - 文件路径配置
- `validation_program/config/error_messages.json` - 错误消息配置
- `validation_program/config/reporting_config.json` - 报告生成配置

## 验证报告

验证完成后，程序会生成详细的验证报告，内容包括：

1. 验证文件信息
   - 输入文件路径
   - 输出目录路径
   - 检测到的输出文件

2. 验收结果
   - 整体验收结果（通过/不通过）
   - 各类别验证结果统计

3. 详细测试结果
   - 输入文件验证结果
   - 处理逻辑验证结果
   - 输出文件验证结果

4. 错误详情和建议修复方法

## 常见问题

### 找不到输入文件

如果程序无法找到输入文件，它会搜索可能的文件并提供建议：

```
错误: 采购装箱单文件不存在: D:\project\export_reimport\testfiles\original_packing_list.xlsx
请使用 --packing-list 指定正确的文件路径

可能的装箱单文件:
  - D:\project\export_reimport\other_files\packing_list_sample.xlsx
```

### 验证失败

如果验证失败，查看生成的报告了解详细原因：

```
验收结果: 不通过
详细报告已生成: D:\project\export_reimport\reports\validation_report.md
```

## 代码结构

- `validation_program/main.py` - 主程序
- `validation_program/run_validation.py` - 示例运行脚本
- `validation_program/validators/` - 验证器模块
  - `input_validator.py` - 输入文件验证器
  - `output_validator.py` - 输出文件验证器
  - `process_validator.py` - 处理逻辑验证器
  - `utils.py` - 工具函数
- `validation_program/config/` - 配置文件目录