---
title: 通过电子邮件发送图表和Excel图像
description: 了解如何使用脚本Office脚本Power Automate提取图表和表格的图像并Excel电子邮件。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 63a4bdb16bdf5923bf49f26fcba163fc3f0b7354
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/15/2021
ms.locfileid: "59335065"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>使用Office脚本Power Automate脚本和脚本来发送图表和表格的电子邮件图像

此示例使用Office脚本Power Automate创建图表。 然后，它通过电子邮件发送图表及其基表的图像。

## <a name="example-scenario"></a>示例应用场景

* 计算可获取最新结果。
* 创建图表。
* 获取图表和表格图像。
* 使用电子邮件向图像Power Automate。

_输入数据_

:::image type="content" source="../../images/input-data.png" alt-text="显示输入数据表格的工作表。":::

_输出图表_

:::image type="content" source="../../images/chart-created.png" alt-text="创建的柱形图显示客户到期金额。":::

_通过流收到Power Automate的电子邮件_

:::image type="content" source="../../images/email-received.png" alt-text="由流发送的电子邮件，Excel嵌入正文中的图表。":::

## <a name="solution"></a>解决方案

此解决方案由两部分组成：

1. [一Office脚本，用于计算和提取Excel图表和表](#sample-code-calculate-and-extract-excel-chart-and-table)
1. 一Power Automate调用脚本并通过电子邮件发送结果的流。 有关操作方法的示例，请参阅使用 Power Automate 创建[自动化工作流](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。

## <a name="sample-excel-file"></a>示例Excel文件

下载 <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> 工作簿的工作簿。 添加以下脚本以自己试用示例！

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>示例代码：计算并提取Excel图表和表

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
1. 选择 **"手动触发流"，** 然后选择"创建 **"。**
1. 通过运行 **脚本操作** 添加使用 **Excel Online (Business)** 连接器的新步骤。  对操作使用以下值。
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件**：工作簿 ([选择器选项选择)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **脚本**：脚本名称

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="已完成的 Excel Online (Business) 连接器Power Automate。":::
1. 此示例使用 Outlook 作为电子邮件客户端。 可以使用支持的任何Power Automate连接器，但其余步骤假定你已选择Outlook。 添加一 **个新** 步骤，该步骤使用 **Office 365 Outlook** 连接器和"发送和电子邮件 (**V2)** 操作。 对操作使用以下值。
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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="已完成的Office 365 Outlook连接器Power Automate。":::
1. 保存流并试用。使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>培训视频：提取图表和表格的图像和电子邮件图像

[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/152GJyqc-Kw)。
