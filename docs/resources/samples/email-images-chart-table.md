---
title: 通过电子邮件发送 Excel 图表和表格的图像
description: 了解如何使用 Office 脚本和 Power Automate 提取 Excel 图表和表格的图像，并通过电子邮件发送这些图像。
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de3cf16537cb12db45d4d465d367d797d053afc4
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754808"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>使用 Office 脚本和 Power Automate 通过电子邮件发送图表和表格的图像

此示例使用 Office 脚本和 Power Automate 创建图表。 然后，它通过电子邮件发送图表及其基表的图像。

## <a name="example-scenario"></a>示例应用场景

* 计算可获取最新结果。
* 创建图表。
* 获取图表和表格图像。
* 使用 Power Automate 通过电子邮件发送图像。

_输入数据_

:::image type="content" source="../../images/input-data.png" alt-text="显示输入数据表格的工作表。":::

_输出图表_

:::image type="content" source="../../images/chart-created.png" alt-text="创建的柱形图显示客户到期金额。":::

_通过 Power Automate 流接收的电子邮件_

:::image type="content" source="../../images/email-received.png" alt-text="由显示在正文中嵌入的 Excel 图表的流发送的电子邮件。":::

## <a name="solution"></a>解决方案

此解决方案由两部分组成：

1. [用于计算和提取 Excel 图表和表格的 Office 脚本](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Power Automate 流，用于调用脚本并通过电子邮件发送结果。 有关操作方法的示例，请参阅使用 Power [Automate 创建自动化工作流](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>示例代码：计算和提取 Excel 图表和表

以下脚本计算并提取 Excel 图表和表格。

下载示例文件 <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> 并使用此脚本尝试一下！

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>培训视频：提取图表和表格的图像和电子邮件图像

[![观看分步视频，了解如何提取图表和表格的图像并通过电子邮件发送图像](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "如何提取图表和表格的图像和通过电子邮件发送图像的分步视频")
