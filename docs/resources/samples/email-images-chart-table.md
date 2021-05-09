---
title: 通过电子邮件发送图表和Excel图像
description: 了解如何使用脚本Office脚本Power Automate提取图表和Excel图像并通过电子邮件发送。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: f8b52cbf8c19b93c5fc4288fe97775a25e922ab9
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285855"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>使用Office脚本Power Automate脚本和脚本，以通过电子邮件发送图表和表格的图像

此示例使用Office脚本Power Automate创建图表。 然后，它通过电子邮件发送图表及其基表的图像。

## <a name="example-scenario"></a>示例应用场景

* 计算可获取最新结果。
* 创建图表。
* 获取图表和表格图像。
* 使用电子邮件向图像Power Automate。

_输入数据_

:::image type="content" source="../../images/input-data.png" alt-text="显示输入数据表格的工作表":::

_输出图表_

:::image type="content" source="../../images/chart-created.png" alt-text="创建的柱形图显示客户到期金额":::

_通过流收到Power Automate的电子邮件_

:::image type="content" source="../../images/email-received.png" alt-text="由显示在正文中嵌入Excel图表的流发送的电子邮件":::

## <a name="solution"></a>解决方案

此解决方案由两部分组成：

1. [一Office用于计算和提取图表和Excel的脚本](#sample-code-calculate-and-extract-excel-chart-and-table)
1. 一Power Automate调用脚本并通过电子邮件发送结果的流。 有关如何执行此操作的示例，请参阅使用 Power Automate 创建自动化[工作流](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>示例代码：计算并提取Excel图表和表

以下脚本计算并提取图表Excel图表。

下载示例文件 <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> 并使用此脚本尝试一下！

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate流：通过电子邮件发送图表和表格图像

此流运行脚本，并通过电子邮件发送返回的图像。

1. 创建新的即时 **云流**。
1. 选择 **"手动触发流"，** 然后按"**创建"。**
1. 添加一 **个新** 步骤，该步骤使用 **Excel Online (Business)** 连接器和 **运行脚本 (预览)** 操作。 对操作使用以下值：
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件**：工作簿 ([选择器选项选择)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **脚本**：脚本名称

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="已完成的 Excel Online (Business) 连接器Power Automate":::
1. 此示例使用 Outlook 作为电子邮件客户端。 可以使用支持的任何Power Automate连接器，但其余步骤假定你已选择Outlook。 添加一 **个新** 步骤，该步骤使用 **Office 365 Outlook** 连接器和 **V2** (发送) 操作。 对操作使用以下值：
    * **目标**：测试电子邮件帐户 (或个人) 
    * **主题**：请查看报告数据
    * 对于" **正文** "字段，选择"代码视图 `</>` " () 并输入以下内容：

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="已完成的Office 365 Outlook连接器Power Automate":::
1. 保存流并试用。

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>培训视频：提取图表和表格的图像和电子邮件图像

[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/152GJyqc-Kw)。
