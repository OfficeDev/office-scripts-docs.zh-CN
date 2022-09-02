---
title: Email Excel 图表和表的图像
description: 了解如何使用 Office 脚本和 Power Automate 提取 Excel 图表和表的图像并发送电子邮件。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: dbf9135723a735321c99991d94f4b4387d800702
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572463"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>使用 Office 脚本和 Power Automate 通过电子邮件发送图表和表格的图像

此示例使用 Office 脚本和 Power Automate 创建图表。 然后，它会通过电子邮件发送图表及其基表的图像。

## <a name="example-scenario"></a>示例方案

* 计算以获取最新结果。
* 创建图表。
* 获取图表和表格图像。
* 使用 Power Automate Email图像。

_输入数据_

:::image type="content" source="../../images/input-data.png" alt-text="显示输入数据表的工作表。":::

_输出图表_

:::image type="content" source="../../images/chart-created.png" alt-text="创建的柱形图显示客户应付的金额。":::

_通过 Power Automate 流接收的Email_

:::image type="content" source="../../images/email-received.png" alt-text="流发送的电子邮件，其中显示了嵌入在正文中的 Excel 图表。":::

## <a name="solution"></a>解决方案

此解决方案包含两个部分：

1. [用于计算和提取 Excel 图表和表的 Office 脚本](#sample-code-calculate-and-extract-excel-chart-and-table)
1. 用于调用脚本并通过电子邮件发送结果的 Power Automate 流。 有关如何执行此操作的示例，请参阅 [使用 Power Automate 创建自动化工作流](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。

## <a name="sample-excel-file"></a>示例 Excel 文件

下载现成工作簿 [ 的email-chart-table.xlsx](email-chart-table.xlsx) 。 添加以下脚本以自行尝试示例！

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>示例代码：计算和提取 Excel 图表和表

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate 流：Email图表和表图像

此流运行脚本，并向返回的图像发送电子邮件。

1. 创建新的 **即时云流**。
1. 选择 **“手动触发流** ”，然后选择 **“创建**”。
1. 添加使用 **Excel Online (Business)** 连接器和 **运行脚本** 操作 **的新步骤**。 对操作使用以下值。
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件**：使用 [文件选择器) 选择了](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control) 工作簿 (
    * **脚本**：脚本名称

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Power Automate 中已完成的 Excel Online (Business) 连接器。":::
1. 此示例使用 Outlook 作为电子邮件客户端。 可以使用 Power Automate 支持的任何电子邮件连接器，但其余步骤假定你选择了 Outlook。 添加使用 **Office 365 Outlook** 连接器和 **发送和电子邮件 (V2)** 操作 **的新步骤**。 对操作使用以下值。
    * **收件** 人：测试电子邮件帐户 (或个人电子邮件) 
    * **主题**：请查看报表数据
    * 对于“ **正文** ”字段，选择“代码视图” (`</>`) 并输入以下内容：

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Power Automate 中已完成的 Office 365 Outlook 连接器。":::
1. 保存流并试用。使用流编辑器页上的 **“测试** ”按钮，或通过“ **我的流** ”选项卡运行流。出现提示时，请务必允许访问。

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>培训视频：提取图表和表格的图像并发送电子邮件

[观看苏迪 · 拉马穆尔西在 YouTube 上浏览这个示例](https://youtu.be/152GJyqc-Kw)。
