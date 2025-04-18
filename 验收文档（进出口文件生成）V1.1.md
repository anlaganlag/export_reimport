# 程序校验及文档处理与验收规范
## 一、程序校验输入文件合法性
### （一）采购装箱单合法性校验
1. 表头第一行必须有“采购装箱单”字样，且包含编号。
2. 采购装箱单表头字段名分两行，一行中文、一行英文。
3. 每个货物数据行的单件总净重需小于总毛重。
4. 核查采购装箱单表尾的总数量、总体积、总毛重、总净重等汇总数据是否正确。

### （二）政策文件合法性校验
1. 政策文件编号要与采购装箱单编号一致。
2. 汇率保留4位小数。
3. 政策文件需包含公司信息和银行信息。

## 二、程序处理要求
1. 程序开始运行时，用文件浏览对话框让用户选择输入文件的位置，包括采购装箱单和政策文件，材料信息数据库。文件浏览对话框打开时默认文件夹的位置是上一次的输入文件位置。
2. 进出口模板文件固定放在文件夹"Template"中。
3. 程序运行时发生输入文件错误，数据处理异常等运行时错误需要提示用户错误来源，修正方法，然后退出程序。
3. 根据采购装箱单里"Purchasing Company采购公司"的名称，调用对应公司的进出口Commercial Invoice模板和Packing List模板，并将模板表头、表尾应用到相应进出口文件中。
4. 依据采购装箱单中"Purchasing Company采购公司"的名称，调用对应公司的出口报关单模板，把模板表头、表尾应用到出口报关单文件中，并填好表头各栏位数据。
5. 进出口文档的发票号采用"字母+日期+流水号"格式填写采购装箱单编号。创想公司格式为"CXCIyyyymmdd####"，凯旋公司格式为"KXCIyyyymmdd####"，"####"为当日流水码，格式不符时提示用户。
6. 进出口文档日期使用当前日期，格式为"yyyy/mm/dd"。
7. 进口发票和出口发票汇总行的下一行填写发票总金额的英文大写描述。
8. 生成的出口报关单为单独文件。
9. 出口发票和出口装箱单合并为一个文件，出口装箱单页名为"PL"，出口发票页名为发票号码，多张出口发票也放在同一文件内。
10. 进口发票和进口装箱单合并为一个文件，进口装箱单页名为"PL"，进口发票页名为发票号码，多张进口发票同样放在同一文件内。
11. 出口Commercial Invoice单价保留6位小数，总价保留2位小数。
12. 项目合并规则：针对相同物料编号（Part number）的物料，CIF单价四舍五入保留到小数点后4位，物料编号和价格都相同的项目，合并数量为一项。

## 三、出口发票验收要求清单
1. 基于采购装箱单，应输出出口发票、出口Packing List、进口发票、进口Packing List。
2. 出口发票与采购装箱单字段对应关系：
    - Part Number => Part Number料号
    - 名称 => Commercial Invoice Description供应商开票名称
    - Model Number => Model Number型号
    - Quantity => Quantity数量
    - Unit => Unit单位
3. 出口发票的Quantity等于采购装箱单总数量。
4. 出口发票的Unit Price (CIF, USD)大于采购装箱单的Unit Price (Excl. Tax, CNY)采购单价（不含税）。
5. 出口发票的Total Amount (CIF, USD)大于FOB价格+总运费+总保费。

## 四、出口Packing List验收要求清单
1. 出口Packing List与采购装箱单字段对应关系：
    - Part Number 对应 Part Number（料号）
    - 名称 对应 Commercial Invoice Description（供应商开票名称）
    - Model Number 对应 Model Number（型号）
    - Quantity 对应 Quantity（数量）
    - Total Carton Quantity 对应 Total Carton Quantity（总件数）
    - Total Volume (CBM) 对应 Total Volume (CBM)（总体积）
    - Total Gross Weight (kg) 对应 Total Gross Weight (kg)（总毛重）
    - Total Net Weight (kg) 对应 Total Net Weight (kg)（总净重）
    - Carton Number 对应 Carton Number（箱号）
2. 出口Packing List的总数量、总件数、总体积、总毛重、总净重与采购装箱单相应字段值完全一致。

## 五、进口发票验收要求清单
1. 按项目Project项目名称 + Plant Location工厂地点拆分生成多份发票。
2. 进口发票与采购装箱单字段对应关系：
    - Part Number 对应 Part Number（料号）
    - Commodity Description (Customs) 对应 Commodity Description (Customs)（进口清关货描）
    - Quantity 对应 Quantity（数量）
    - Unit 对应 Unit（单位）
    - Total Net Weight (kg) 对应 Total Net Weight (kg)总净重
    - Plant Location 对应 Plant Location（工厂地点）材料信息数据库
3. 进口发票总数量等于采购装箱单总数量。
4. 进口发票的Unit Price (CIF, USD)大于采购装箱单的Unit Price (Excl. Tax, CNY)采购单价（不含税）。
5. 进口发票的Total Amount (CIF, USD)大于FOB价格+总运费+总保费。

## 六、进口装箱单验收需求清单
1. 进口Packing List与采购装箱单字段对应关系：
    - Part Number 对应 Part Number（料号）
    - Commodity Description (Customs) 对应 Commodity Description (Customs)（进口清关货描）
    - Quantity 对应 Quantity（数量）
    - Total Carton Quantity 对应 Total Carton Quantity（总件数）
    - Total Volume (CBM) 对应 Total Volume (CBM)（总体积）
    - Total Gross Weight (kg) 对应 Total Gross Weight (kg)（总毛重）
    - Total Net Weight (kg) 对应 Total Net Weight (kg)（总净重）
    - Carton Number 对应 Carton Number（箱号）
2. 进口Packing List的总数量、总件数、总体积、总毛重、总净重与采购装箱单相应字段值完全一致。

## 七、输出文档要求
1. 输出文档格式与提供的模板一致。
2. 输出文件命名规则：
    - 出口报关单文件名：报关单-（发票号）
    - 出口文档（装箱单和发票）文件名：出口-（发票号）
    - 进口文档（装箱单和发票）文件名：进口-（发票号）

## 八、程序注意事项与解决方案
1. 输入文件格式校验：
   - 注意事项：输入的采购装箱单、政策文件、材料信息数据库等需严格符合预定格式。
   - 解决方案：程序应在读取文件后进行字段名、数据类型、必填项等校验，发现格式不符时及时提示用户并退出。
2. 异常处理：
   - 注意事项：运行过程中可能遇到文件缺失、数据异常、模板不匹配等问题。
   - 解决方案：所有异常需捕获并给出明确的错误来源和修正建议，避免程序无提示崩溃。
3. 模板一致性：
   - 注意事项：生成的文档需严格按照模板格式输出，避免字段遗漏或顺序错误。
   - 解决方案：读取模板结构，自动对齐字段，输出前进行字段完整性和顺序校验。
4. 数据精度与四舍五入：
   - 注意事项：金额、单价、汇率等需按要求保留小数位，防止精度丢失。
   - 解决方案：统一采用高精度数据类型，输出前按规则四舍五入。
5. 文件命名规范：
   - 注意事项：输出文件名需严格遵循命名规则，防止重名或格式错误。
   - 解决方案：生成文件名时自动校验格式，若不符则提示用户。
6. 合并规则与去重：
   - 注意事项：相同物料编号和价格的项目需合并，数量相加，避免重复。
   - 解决方案：处理数据时先按物料编号和价格分组，合并数量，输出前再次校验无重复项。
7. 历史路径记忆：
   - 注意事项：文件浏览对话框应记忆上次输入文件夹路径，提升用户体验。
   - 解决方案：将上次选择的路径保存到本地配置文件，启动时自动读取。
8. 多语言与单位转换：
   - 注意事项：部分字段涉及中英文、单位换算，需确保一致性。
   - 解决方案：建立字段映射表和单位换算函数，自动转换并校验。
9. 汇总校验：
   - 注意事项：表尾汇总数据需与明细数据一致。
   - 解决方案：自动计算明细汇总，与表尾数据比对，不一致时提示。
10. 日志与追溯：
    - 注意事项：关键操作和异常需有日志记录，便于问题追溯。
    - 解决方案：实现日志功能，记录文件名、时间、操作内容和异常信息。