---
title: 通过电子邮件发送图表和Excel图像
description: 了解如何使用脚本Office脚本Power Automate提取图表和Excel图像并通过电子邮件发送。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545772"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="5e299-103">使用Office脚本Power Automate脚本和脚本，以通过电子邮件发送图表和表格的图像</span><span class="sxs-lookup"><span data-stu-id="5e299-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="5e299-104">此示例使用Office脚本Power Automate创建图表。</span><span class="sxs-lookup"><span data-stu-id="5e299-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="5e299-105">然后，它通过电子邮件发送图表及其基表的图像。</span><span class="sxs-lookup"><span data-stu-id="5e299-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="5e299-106">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="5e299-106">Example scenario</span></span>

* <span data-ttu-id="5e299-107">计算可获取最新结果。</span><span class="sxs-lookup"><span data-stu-id="5e299-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="5e299-108">创建图表。</span><span class="sxs-lookup"><span data-stu-id="5e299-108">Create chart.</span></span>
* <span data-ttu-id="5e299-109">获取图表和表格图像。</span><span class="sxs-lookup"><span data-stu-id="5e299-109">Get chart and table images.</span></span>
* <span data-ttu-id="5e299-110">使用电子邮件向图像Power Automate。</span><span class="sxs-lookup"><span data-stu-id="5e299-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="5e299-111">_输入数据_</span><span class="sxs-lookup"><span data-stu-id="5e299-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="显示输入数据表格的工作表":::

<span data-ttu-id="5e299-113">_输出图表_</span><span class="sxs-lookup"><span data-stu-id="5e299-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="创建的柱形图显示客户到期金额":::

<span data-ttu-id="5e299-115">_通过流收到Power Automate的电子邮件_</span><span class="sxs-lookup"><span data-stu-id="5e299-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="由显示在正文中嵌入Excel图表的流发送的电子邮件":::

## <a name="solution"></a><span data-ttu-id="5e299-117">解决方案</span><span class="sxs-lookup"><span data-stu-id="5e299-117">Solution</span></span>

<span data-ttu-id="5e299-118">此解决方案由两部分组成：</span><span class="sxs-lookup"><span data-stu-id="5e299-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="5e299-119">一Office用于计算和提取图表和Excel的脚本</span><span class="sxs-lookup"><span data-stu-id="5e299-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="5e299-120">一Power Automate调用脚本并通过电子邮件发送结果的流。</span><span class="sxs-lookup"><span data-stu-id="5e299-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="5e299-121">有关如何执行此操作的示例，请参阅使用 Power Automate 创建自动化[工作流](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="5e299-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="5e299-122">示例代码：计算并提取Excel图表和表</span><span class="sxs-lookup"><span data-stu-id="5e299-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="5e299-123">以下脚本计算并提取图表Excel图表。</span><span class="sxs-lookup"><span data-stu-id="5e299-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="5e299-124">下载示例文件 <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> 并使用此脚本尝试一下！</span><span class="sxs-lookup"><span data-stu-id="5e299-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="5e299-125">Power Automate流：通过电子邮件发送图表和表格图像</span><span class="sxs-lookup"><span data-stu-id="5e299-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="5e299-126">此流运行脚本，并通过电子邮件发送返回的图像。</span><span class="sxs-lookup"><span data-stu-id="5e299-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="5e299-127">创建新的即时 **云流**。</span><span class="sxs-lookup"><span data-stu-id="5e299-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="5e299-128">选择 **"手动触发流"，** 然后按"**创建"。**</span><span class="sxs-lookup"><span data-stu-id="5e299-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="5e299-129">使用 Run **脚本操作** 添加使用 **Excel Online (Business)** 连接器的新步骤。 </span><span class="sxs-lookup"><span data-stu-id="5e299-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="5e299-130">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="5e299-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="5e299-131">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="5e299-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="5e299-132">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="5e299-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="5e299-133">**文件**：工作簿 ([选择器选项选择)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="5e299-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="5e299-134">**脚本**：脚本名称</span><span class="sxs-lookup"><span data-stu-id="5e299-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="已完成的 Excel Online (Business) 连接器Power Automate":::
1. <span data-ttu-id="5e299-136">此示例使用 Outlook 作为电子邮件客户端。</span><span class="sxs-lookup"><span data-stu-id="5e299-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="5e299-137">可以使用支持的任何Power Automate连接器，但其余步骤假定你已选择Outlook。</span><span class="sxs-lookup"><span data-stu-id="5e299-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="5e299-138">添加一 **个新** 步骤，该步骤使用 **Office 365 Outlook** 连接器和 **V2** (发送) 操作。</span><span class="sxs-lookup"><span data-stu-id="5e299-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="5e299-139">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="5e299-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="5e299-140">**目标**：测试电子邮件帐户 (或个人) </span><span class="sxs-lookup"><span data-stu-id="5e299-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="5e299-141">**主题**：请查看报告数据</span><span class="sxs-lookup"><span data-stu-id="5e299-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="5e299-142">对于" **正文** "字段，选择"代码视图 `</>` " () 并输入以下内容：</span><span class="sxs-lookup"><span data-stu-id="5e299-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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
1. <span data-ttu-id="5e299-144">保存流并试用。</span><span class="sxs-lookup"><span data-stu-id="5e299-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="5e299-145">培训视频：提取图表和表格的图像和电子邮件图像</span><span class="sxs-lookup"><span data-stu-id="5e299-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="5e299-146">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/152GJyqc-Kw)。</span><span class="sxs-lookup"><span data-stu-id="5e299-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
