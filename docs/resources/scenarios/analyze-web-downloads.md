---
title: Office 脚本示例方案：分析 web 下载
description: 在将该信息组织到表中之前，获取 Excel 工作簿中的原始 internet 流量数据并确定原始位置的示例。
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 9ee12c8d4ca7c191168e3734d7cd9eadc333c165
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700132"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Office 脚本示例方案：分析 web 下载

在这种情况下，您将负责分析来自公司网站的下载报告。 此分析的目标是确定 web 流量是来自世界各地还是其他地方。

您的同事将原始数据上传到工作簿。 每周的数据集都有自己的工作表。 还有一个显示每周趋势的表格和图表的**摘要**工作表。

您将开发一个脚本，用于分析活动工作表中每周的下载数据。 它将分析与每个下载关联的 IP 地址，并确定是否来自美国。 该答案将作为布尔值（"TRUE" 或 "FALSE"）插入到工作表中，条件格式将应用于这些单元格。 将在工作表上汇总 IP 地址位置结果，并将其复制到摘要表。

## <a name="scripting-skills-covered"></a>涵盖的脚本技能

- 文本解析
- 脚本中的 Subfunctions
- 条件格式
- 表

## <a name="demo-video"></a>演示视频

此示例是 demoed 年2月2020的 Office 外接程序开发人员社区呼叫的一部分。

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a>设置说明

1. 将<a href="analyze-web-downloads.xlsx">analyze-web-downloads</a>下载到你的 OneDrive。

2. 使用适用于 web 的 Excel 打开工作簿。

3. 在 "**自动化**" 选项卡上，打开**代码编辑器**。

4. 在 "**代码编辑器**" 任务窗格中，按 "**新建脚本**"，并将以下脚本粘贴到编辑器中。

    ```TypeScript
      async function main(context: Excel.RequestContext) {
        let currentWorksheet = context.workbook.worksheets
          .getActiveWorksheet();
        // Get the values of the active range of the active worksheet.
        let logRange = currentWorksheet.getUsedRange().load("values");

        // Get the Summary worksheet and table.
        let summaryWorksheet = context.workbook.worksheets.getItem("Summary");
        let summaryTable = context.workbook.tables.getItem("Table1");

        // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
        let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1)
          .load("address");

        // Get the values of all the US IP addresses.
        let ipRange = context.workbook.worksheets
          .getItem("USIPAddresses")
          .getUsedRange()
          .load("values");
        await context.sync();

        // Remove the first row.
        let topRow = logRange.values.shift();

        // Create a new array to contain the boolean representing if this is a US IP address.
        let newCol = [[]];

        // Go through each row in worksheet and add Boolean.
        for (let i = 0; i < logRange.values.length; i++) {
          let curRowIP = logRange.values[i][1];
          if (findIP(ipRange.values, ipAddressToInteger(curRowIP)) > 0) {
            newCol.push([true]);
          } else {
            newCol.push([false]);
          }
        }

        // Remove the empty column header and add proper heading.
        newCol.shift();
        newCol.unshift(["Is US IP"]);

        // Write the result to the spreadsheet.
        isUSColumn.values = newCol;
        addSummaryData();
        applyConditionalFormatting();
        currentWorksheet.getUsedRange().format.autofitColumns();

        // Get the calculated summary data.
        let summaryRange = currentWorksheet.getRange("J2:M2").load("values");
        await context.sync();

        // Add the corresponding row to the summary table.
        summaryTable.rows.add(null, summaryRange.values);

        // Function to apply conditional formatting to the new column.
        function applyConditionalFormatting() {
          // Add conditional formatting to the new column.
          let conditionalFormatTrue = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          let conditionalFormatFalse = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          // Set TRUE to light blue and FALSE to light orange.
          conditionalFormatTrue.cellValue.format.fill.color = "#8FA8DB";
          conditionalFormatTrue.cellValue.rule = {
            formula1: "=TRUE",
            operator: "EqualTo"
          };
          conditionalFormatFalse.cellValue.format.fill.color = "#F8CCAD";
          conditionalFormatFalse.cellValue.rule = {
            formula1: "=FALSE",
            operator: "EqualTo"
          };
        }

        // Adds the summary data to the current sheet and to the summary table.
        function addSummaryData() {
          // Add a summary row and table.
          let summaryHeader = [["Year", "Week", "US", "Other"]];
          let countTrueFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=TRUE")/' + (newCol.length - 1);
          let countFalseFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=FALSE")/' + (newCol.length - 1);

          let summaryContent = [
            [
              '=TEXT(A2,"YYYY")',
              '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
              countTrueFormula,
              countFalseFormula
            ]
          ];
          let summaryHeaderRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J1:M1");
          let summaryContentRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J2:M2");
          summaryHeaderRow.values = summaryHeader;
          summaryContentRow.values = summaryContent;
          let formats = [[".000", ".000"]];
          summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).numberFormat = formats;
        }
      }

      // Translate an IP address into an integer.
      function ipAddressToInteger(ipAddress: string) {
        // Split the IP address into octets.
        let octets = ipAddress.split(".");

        // Create a number for each octet and do the math to create the integer value of the IP address.
        let fullNum =
          // Define an arbitrary number for the last octet.
          111 +
          parseInt(octets[2]) * 256 +
          parseInt(octets[1]) * 65536 +
          parseInt(octets[0]) * 16777216;
        return fullNum;
      }

      // Return the row number where the ip address is found.
      function findIP(ipLookupTable: number[][], n: number) {
        for (let i = 0; i < ipLookupTable.length; i++) {
          if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
            return i;
          }
        }
        return -1;
      }
    ```

5. 重命名脚本以**分析 Web 下载**并保存它。

## <a name="running-the-script"></a>运行脚本

导航到任意**一周\* **工作表，并运行**分析 Web 下载**脚本。 该脚本将应用在当前工作表上标签的条件格式和位置。 它还将更新**摘要**工作表。

### <a name="before-running-the-script"></a>运行脚本之前

![显示原始 web 流量数据的工作表。](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a>运行脚本后

![显示以前的 web 流量行的格式化 IP 位置信息的工作表。](../../images/scenario-analyze-web-downloads-after.png)

![汇总表和图表，其中汇总了运行脚本的工作表。](../../images/scenario-analyze-web-downloads-table.png)
