# 装运清单处理工作流程

本文档概述了处理装运清单并生成出口和复进口收据的完整工作流程。

## 流程图

### CIIBER原始装箱单到出口/在进口发票工作流程

```mermaid
graph TD
    Title[CIIBER 原始装箱单
     到出口/在进口发票流程图]
```

```mermaid
graph TD
    %% 输入文件
    OPL[原始装箱单] --> CFOB[计算单个物料FOB价格]
    
    %% FOB计算
    PF[政策文件] --> |加价百分比| CFOB
    
    %% FOB到CIF转换
    PF --> |单批次总净重|CCIF[计算CIF价格]
    PF --> |单批次总运费|CCIF
    PF --> |汇率|CCIF
    PF --> |保险费率|CCIF
    CFOB --> |FOB原始发票|CCIF
    
    %% CIF到最终发票处理
    CCIF --> |合并CIF原始发票|MERGE[合并CIF原始发票同类项]
    MERGE --> EXP[最终出口发票]

    %% 拆分原始发票
    CCIF --> |CIF原始发票|SI[按目的地拆分发票]
    OPL --> |目的地工厂|SI
    SI --> RIMP1[最终复进口发票1]
    SI --> RIMP2[最终复进口发票2]

    %% 样式
    classDef input fill:#e1f5fe,stroke:#01579b
    classDef process fill:#fff3e0,stroke:#e65100
    classDef output fill:#e8f5e9,stroke:#1b5e20
    
    class OPL,PF input
    class CFOB,CCIF,SI,CFEI,MIL2,MIL3,MIL4,MERGE process
    class EXP,EXP1,RIMP1,RIMP2 output
```

### 单个物料FOB价格计算详情

```mermaid
graph LR
    %% 输入值
    OPL[原始装箱单] --> |单价|FOBUP["计算FOB单价
    =
    单价 × (1 + 加价%)"]

    PF[政策文件] --> |"加价百分比"|FOBUP

    
    %% 计算过程
    FOBUP -- "FOB单价" --> FOBT_RESULT[单个物料FOB单价]

    
    %% 样式
    classDef input fill:#e1f5fe,stroke:#01579b
    classDef process fill:#fff3e0,stroke:#e65100
    classDef output fill:#e8f5e9,stroke:#1b5e20
    
    class OPL,PF input
    class FOBUP,FOBT process
    class FOBT_RESULT output
```

### 单个物料保险费计算详情

```mermaid
graph LR
    %% 输入值
    FOBUP[单个物料FOB价格计算] --> |FOB单价|PIA["计算单个物料保险费
    =
    FOB单价
    x
    保险系数
    ×
    保险费率"]
    
    PF[政策文件] --> |保险系数|PIA
    PF[政策文件] --> |保险费率|PIA
    
    %% 计算过程
    PIA --> PIA_RESULT[单个物料保险费]
    
    %% 样式
    classDef input fill:#e1f5fe,stroke:#01579b
    classDef process fill:#fff3e0,stroke:#e65100
    classDef output fill:#e8f5e9,stroke:#1b5e20
    
    class FOBUP,PF input
    class PIA process
    class PIA_RESULT output
```

### 单个物料运费计算详情

```mermaid
graph LR
    %% 单位运费率的输入值
    User[政策文件] --> |总运费金额|UFR["计算单位运费率
    =
    总运费金额
    ÷
    总净重
    (人民币/公斤)"]
    OPL[原始装箱单] --> |总净重|UFR
    
    %% 单个物料运费的输入值
    OPL[原始装箱单] --> |单个物料净重|PFC["计算单个物料运费
    =
    单个物料净重
    ×
    单公斤运费率
    (人民币/单个物料)"]
    OPL[原始装箱单] --> |单个物料数量|PFC
    
    %% 计算流程
    UFR --> |"单位运费率"|PFC
    PFC --> PFC_RESULT[单个物料运费]
    
    %% 样式
    classDef input fill:#e1f5fe,stroke:#01579b
    classDef process fill:#fff3e0,stroke:#e65100
    classDef output fill:#e8f5e9,stroke:#1b5e20
    
    class User,OPL input
    class UFR,PFC process
    class PFC_RESULT output
```

### 单个物料CIF价格计算详情

```mermaid
graph LR
    %% 输入组件
    FOBUP[单个物料FOB单价计算] --> |单个物料单个物料FOB单价|CIF["计算CIF价格
    =
    单个物料FOB单价
    +
    单个物料运费
    +
    单个物料保险费"]
    FC[单个物料运费计算] --> |单个物料运费|CIF
    IC[单个物料保险费计算] --> |单个物料保险费|CIF
    
    %% 最终CIF价格
    CIF --> CIF_RESULT[单个物料CIF价格]
    
    %% 样式
    classDef input fill:#e1f5fe,stroke:#01579b
    classDef process fill:#fff3e0,stroke:#e65100
    classDef output fill:#e8f5e9,stroke:#1b5e20
    
    class FOBUP,FC,IC input
    class CIF process
    class CIF_RESULT output
```

## 测试文件和输出

### 输入文件

系统使用以下Excel文件作为测试输入：

1. **原始装箱单 (testfiles/original_packing_list.xlsx)**
   - 包含物料编号、描述、单价、数量、净重
   - 每个物料的目的地工厂信息
   - 物料规格和包装信息

2. **政策文件 (testfiles/policy.xlsx)**
   - 加价百分比设置
   - 保险费率和保险系数
   - 汇率信息
   - 运费总金额
   - 其他计算参数

### 输出文件

系统处理后生成以下Excel输出文件：

1. **最终出口发票 (outputs/export_invoice.xlsx)**
   - 合并后的CIF价格发票
   - 包含所有物料的汇总信息
   - 总金额和计算明细

2. **最终复进口发票 (outputs/reimport_invoice_factory_*.xlsx)**
   - 按目的地工厂拆分的发票
   - 每个目的地工厂对应一个独立的发票文件
   - 包含该工厂相关物料的CIF价格和明细

### 数据处理流程

```mermaid
graph TD
    %% 输入文件
    PL[testfiles/original_packing_list.xlsx] --> Process[处理系统]
    PF[testfiles/policy.xlsx] --> Process
    
    %% 处理步骤
    Process --> FOB[FOB价格计算]
    Process --> CIF[CIF价格计算]
    Process --> Split[发票拆分]
    
    %% 输出文件
    FOB --> Export[outputs/export_invoice.xlsx]
    CIF --> Export
    Split --> Reimport1[outputs/reimport_invoice_factory_A.xlsx]
    Split --> Reimport2[outputs/reimport_invoice_factory_B.xlsx]
    
    %% 样式
    classDef input fill:#e1f5fe,stroke:#01579b
    classDef process fill:#fff3e0,stroke:#e65100
    classDef output fill:#e8f5e9,stroke:#1b5e20
    
    class PL,PF input
    class Process,FOB,CIF,Split process
    class Export,Reimport1,Reimport2,ReimportN output
```

### 测试文件格式要求

#### 原始装箱单 (original_packing_list.xlsx)
- **物料信息表**: 包含物料编号、描述、规格等基本信息
- **价格信息表**: 包含单价、数量、净重等计算所需数据
- **目的地信息表**: 标明每个物料的目的地工厂代码

#### 政策文件 (policy.xlsx)
- **加价政策表**: 不同类型物料的加价百分比
- **费率表**: 保险费率、运费计算参数
- **汇率表**: 不同货币的汇率信息

### 输出文件格式说明

#### 出口发票 (export_invoice.xlsx)
- 包含所有物料的FOB和CIF价格
- 汇总的总金额和各项费用明细
- 按物料类型分类的统计信息

#### 复进口发票 (reimport_invoice_factory_*.xlsx)
- 每个目的地工厂的专属发票
- 仅包含该工厂相关的物料信息
- 该工厂物料的CIF价格和费用明细
