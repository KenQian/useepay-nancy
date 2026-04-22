# Generate the Excel file for the Foreign Exchange for the channels

## 1. Task Objective
* Generate a python script which accept a path where it contains all the required files.

## The main logic of this python script:
* Look for 各通道需换汇情况汇总-<date>.xlsx, where the <date> usually a format of `yyyy-mm-dd` from the past day.
* Validate all the required files, and then get the new date. 
* Make a copy of 各通道需换汇情况汇总-<date>.xlsx with name 各通道需换汇情况汇总-<new date>-wip.xlsx
* Following the details rules mentioned below to update this file.
* Rename above file by removing `-wip` after complete the processing of the file.

## Validation
* There will be following files under the folder
  * 各通道需换汇情况汇总-<date>.xlsx
  * 渠道订单.xls
  * 消费-商户本金(清算后)xls.
  * 退款-商户本金(清算后).xls

## 2. Some basic rules
* Do not skip hidden, filtered, or grouped rows.
* 

## 3. Copy data to `各通道需换汇情况汇总-<date>.xlsx`
* **TargetSheet:**: `账户流水`
  * Do not skip hidden, filtered, or grouped rows.
  * Remove all the content from A to T except for the header 
  * Copy `消费-商户本金(清算后)xls.xls` Sheet `交易详情` content to above sheet's column A-Q
  * Copy `退款-商户本金(清算后).xls` Sheet `交易详情` and append to above sheet's column A-Q
  * Add formula for the following columns and use the formula to generate new value to replace original value. 
    * Column M(`余额`): M = K - L
    * Column R: R = B & E
    * Column S: S = G
* **TargetSheet:** `渠道订单`
  * **SourceSheet-1:**: `渠道订单.xls` -> `交易详情`. Follow the rules to copy data from source to target
    * Skip the rows in SourceSheet-1 where column J(`交易类型`) values in (`预授权申请`, `预授权撤销`)
    * Copy column A-D from SourceSheet-1 to TargetSheet column A-D
    * Copy column F-AH from SourceSheet-1 to TargetSheet column E-AG
    * Copy column E from SourceSheet-1 to TargetSheet column AH
  * **SourceSheet-2:** `各通道需换汇情况汇总-<date>.xlsx` Sheet `特殊的渠道订单`. Following the rules to copy data from source to target
    Notice: After copying `账户流水` in above steps, the column AJ's data may become available instead of #N/A 
    * Copy column A-AH from SourceSheet-2 to TargetSheet column A-AH where SourceSheet-2 column AJ value has valid value (non #N/A)  
* **Create formula and values for the following columns in TargetSheet**:
  * Column AI: AI = D & E
  * Column AJ: AJ = N
  * Column AK: AK = IF(I="退款", -M*(1-3.2%), M*(1-3.2%))
  * Column AL: AL = XLOOKUP(AR,打款币种!$G:$G,打款币种!$F:$F)
  * Column AM: AM = VLOOKUP(AI,账户流水!R:T,2,0)
  * Column AN: AN = VLOOKUP(AI,账户流水!R:T,3,0)
  * Column AO: AO = IF(AL=AM,"是","否")
  * Column AP: AP = VLOOKUP(F,渠道名称!A:B,2,0)
  * Column AQ:
    * When AP == `2号通道`: AQ = XLOOKUP(AH,'二级商户号映射表-A01'!$A:$A,'二级商户号映射表-A01'!$B:$B)
    * When AP == `A07`: AQ = XLOOKUP(AH,'二级商户号映射表-A07'!$A:$A,'二级商户号映射表-A07'!$C:$C)&$AH
    * When AP == `7号通道`: AQ = AB
    * Others: AQ = <Empty>
  * Column AR: AR = AP&AQ&AJ
* **Post process**
  * After setting the formula, for the cells in column AM whose value is `#N/A`, and if column AP value is not (`PayPal`, `Afterpay直连`) (case-insensitive), 
    then 
    * Copy column A-AH to append to sheet `特殊的渠道订单` column A-AH
    * Set `特殊的渠道订单` formula for the new copied row
      * AI = D&E
      * AJ = XLOOKUP(A,账户流水!R:R,账户流水!R:R)
    * Delete the row from `账户流水`