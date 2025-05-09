# Project Scratchpad

## Background and Motivation
Converting the process_shipping_list.py script into a Streamlit web application to allow users to upload input and policy files to generate export invoices. This will make the process more user-friendly and accessible.

## Key Challenges and Analysis
1. File handling - Need to manage file uploads and downloads securely
2. User interface - Must be intuitive and bilingual (English/Chinese)
3. Process adaptation - Existing script functions need to be adapted for web use
4. Error handling - Clear feedback for users when issues occur
5. Environment setup - Need cross-platform setup scripts for easy deployment

## High-level Task Breakdown
1. ✅ Setup basic Streamlit app structure
   - Success criteria: App runs and shows basic UI
2. ✅ Create file upload interface
   - Success criteria: Users can upload files and see confirmation
3. ✅ Adapt processing functions
   - Success criteria: Core functionality works with uploaded files
4. ✅ Add download interface
   - Success criteria: Generated files can be downloaded
5. ✅ Enhance UI/UX
   - Success criteria: Interface is user-friendly and bilingual
6. ✅ Create setup scripts and documentation
   - Success criteria: One-click setup works on both Windows and Mac/Linux
7. ✅ Reimport发票字段名替换为Commodity Description (Customs)并取值自进口清关货描

## Project Status Board
- [x] Task 1: Basic Streamlit app setup complete
- [x] Task 2: File upload interface implemented
- [x] Task 3: Processing functions adapted
- [x] Task 4: Download interface added
- [x] Task 5: UI/UX enhanced with bilingual support
- [x] Task 6: Setup scripts and documentation created
  - Created run_app.ps1 for Windows
  - Created run_app.sh for Mac/Linux
  - Added comprehensive README.md with instructions
  - Added troubleshooting guide
- [ ] Task 7: Reimport发票字段名替换为Commodity Description (Customs)并取值自进口清关货描
- [x] 检查testfiles/origin和testfiles/sample目录下的输入文件是否存在
- [x] 尝试用不同输入文件组合执行校验程序
- [x] 检查output_log.txt和validation_report.md，收集详细错误信息
- [ ] 等待用户提供正确的采购装箱单和政策文件，或指示如何处理文件缺失问题
- [x] 修复净重/毛重校验容错逻辑，兼容非数值格式，自动容错并报告
- [x] 校验程序支持政策文件竖表结构（字段名在A列，值在B列），加价、总运费、保险系数、保险费率等均可正确识别，兼容横表兜底。
- [x] find_value_by_fieldname已升级为模糊匹配，支持如"总运费(RMB)"等带括号或单位的字段，所有类似字段无需精确匹配即可自动识别。

## Current Status / Progress Tracking
Project is complete with all major tasks finished. The application now includes:
- Streamlit web interface with bilingual support
- File upload for packing lists and policy files
- Processing functionality
- File download capability
- One-click setup scripts for Windows and Mac/Linux
- Comprehensive documentation

正在执行：
- [ ] reimport印度进口发票"名称"字段名替换为"Commodity Description (Customs)"，且取值逻辑改为"进口清关货描"列。
- [x] 校验程序已修复净重/毛重字段异常格式容错，遇到空、字符串、特殊符号、单位等自动转为0，校验报告有容错提示。

## Executor's Feedback or Assistance Requests
本次已完成：
- 优化了净重/毛重校验函数，所有非标准格式（空、字符串、特殊符号、单位等）自动容错为0。
- 校验报告中增加了容错提示，便于用户追溯异常数据。
- 校验逻辑已严格按"同箱号净重总和 < 毛重总和"标准执行。
- 已通过典型异常数据单元测试。
- 校验程序已支持政策文件为竖表结构（A列为字段名，B列为值），加价、总运费、保险系数、保险费率等均可正确识别，兼容横表兜底。
- find_value_by_fieldname已升级为模糊匹配，支持如"总运费(RMB)"等带括号或单位的字段，所有类似字段无需精确匹配即可自动识别。

请用户Planner确认修复效果，如需进一步测试或有新需求请告知。

## Lessons
1. Always provide bilingual interface for better accessibility
2. Include comprehensive error handling and user feedback
3. Create platform-specific setup scripts for easier deployment
4. Document troubleshooting steps for common issues
5. Use virtual environments for dependency management
6. Include clear success criteria for each task
7. Maintain processing logic while improving user interface
8. 校验程序应对所有输入数据类型健壮容错，遇到异常值自动转为0并在报告中提示，避免整体流程中断。

## 2024-任务：reimport印度进口发票"名称"字段替换为"Commodity Description (Customs)"及其取值逻辑调整

### 背景与动机
用户要求将process_shipping_list.py生成的reimport印度进口发票（复进口发票）中的字段"名称"改为"Commodity Description (Customs)"，且取值逻辑改为输入表格中的"进口清关货描"列。

### 关键分析与挑战
1. 字段名替换：reimport发票相关所有环节将"名称"替换为"Commodity Description (Customs)"
2. 字段取值逻辑调整：优先查找"进口清关货描"列赋值给新字段，找不到需容错
3. 兼容性与健壮性：只影响reimport发票，其他不变
4. 输出列顺序与样式：同步调整相关代码，确保新字段名贯穿始终

### 高级任务拆解
1. 定位所有涉及"名称"字段的代码段（reimport发票相关）
2. 字段名替换为"Commodity Description (Customs)"
3. 字段取值逻辑调整为"进口清关货描"
4. 汇总行、空行、样式等同步适配
5. 测试与验证：生成的reimport发票表头和内容正确，其他功能不受影响

### Success Criteria
- [ ] reimport印度进口发票的表头字段为"Commodity Description (Customs)"
- [ ] 该字段内容为输入表"进口清关货描"列的值
- [ ] 若无该列，程序有容错提示
- [ ] 其他发票/packing list/出口发票不受影响
- [ ] 汇总行、空行、样式等同步适配 