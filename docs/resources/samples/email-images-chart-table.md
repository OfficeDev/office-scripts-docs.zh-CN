---
title: 通过电子邮件发送 Excel 图表和表格的图像
description: 了解如何使用 Office 脚本和 Power Automate 提取 Excel 图表和表格的图像，并通过电子邮件发送这些图像。
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 7eb12526f97d72de31acdc3c9a4228c670875e2b
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571128"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="17161-103">使用 Office 脚本和 Power Automate 通过电子邮件发送图表和表格的图像</span><span class="sxs-lookup"><span data-stu-id="17161-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="17161-104">此示例使用 Office 脚本和 Power Automate 创建图表。</span><span class="sxs-lookup"><span data-stu-id="17161-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="17161-105">然后，它通过电子邮件发送图表及其基表的图像。</span><span class="sxs-lookup"><span data-stu-id="17161-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="17161-106">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="17161-106">Example scenario</span></span>

* <span data-ttu-id="17161-107">计算可获取最新结果。</span><span class="sxs-lookup"><span data-stu-id="17161-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="17161-108">创建图表。</span><span class="sxs-lookup"><span data-stu-id="17161-108">Create chart.</span></span>
* <span data-ttu-id="17161-109">获取图表和表格图像。</span><span class="sxs-lookup"><span data-stu-id="17161-109">Get chart and table images.</span></span>
* <span data-ttu-id="17161-110">使用 Power Automate 通过电子邮件发送图像。</span><span class="sxs-lookup"><span data-stu-id="17161-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="17161-111">_输入数据_</span><span class="sxs-lookup"><span data-stu-id="17161-111">_Input data_</span></span>

![输入数据](../../images/input-data.png)

<span data-ttu-id="17161-113">_输出图表_</span><span class="sxs-lookup"><span data-stu-id="17161-113">_Output chart_</span></span>

![创建的图表](../../images/chart-created.png)

<span data-ttu-id="17161-115">_通过 Power Automate 流接收的电子邮件_</span><span class="sxs-lookup"><span data-stu-id="17161-115">_Email that was received through Power Automate flow_</span></span>

![接收的电子邮件](../../images/email-received.png)

## <a name="solution"></a><span data-ttu-id="17161-117">解决方案</span><span class="sxs-lookup"><span data-stu-id="17161-117">Solution</span></span>

<span data-ttu-id="17161-118">此解决方案由两部分组成：</span><span class="sxs-lookup"><span data-stu-id="17161-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="17161-119">用于计算和提取 Excel 图表和表格的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="17161-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="17161-120">Power Automate 流，用于调用脚本并通过电子邮件发送结果。</span><span class="sxs-lookup"><span data-stu-id="17161-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="17161-121">有关操作方法的示例，请参阅使用 Power [Automate 创建自动化工作流](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="17161-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="17161-122">示例代码：计算和提取 Excel 图表和表</span><span class="sxs-lookup"><span data-stu-id="17161-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="17161-123">以下脚本计算并提取 Excel 图表和表格。</span><span class="sxs-lookup"><span data-stu-id="17161-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="17161-124">下载示例文件 <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> 并使用此脚本尝试一下！</span><span class="sxs-lookup"><span data-stu-id="17161-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="17161-125">培训视频：提取图表和表格的图像和电子邮件图像</span><span class="sxs-lookup"><span data-stu-id="17161-125">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="17161-126">[![观看分步视频，了解如何提取图表和表格的图像并通过电子邮件发送图像](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "如何提取图表和表格的图像和通过电子邮件发送图像的分步视频")</span><span class="sxs-lookup"><span data-stu-id="17161-126">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
