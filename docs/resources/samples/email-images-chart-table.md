---
title: 通过电子邮件发送图表和Excel图像
description: 了解如何使用脚本Office脚本Power Automate提取图表和Excel图像并通过电子邮件发送。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 50bc65c82df7f5fc68dbebf942c4f607bb6af60a
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313839"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="df6fd-103">使用Office脚本Power Automate脚本和脚本，以通过电子邮件发送图表和表格的图像</span><span class="sxs-lookup"><span data-stu-id="df6fd-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="df6fd-104">此示例使用Office脚本Power Automate创建图表。</span><span class="sxs-lookup"><span data-stu-id="df6fd-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="df6fd-105">然后，它通过电子邮件发送图表及其基表的图像。</span><span class="sxs-lookup"><span data-stu-id="df6fd-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="df6fd-106">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="df6fd-106">Example scenario</span></span>

* <span data-ttu-id="df6fd-107">计算可获取最新结果。</span><span class="sxs-lookup"><span data-stu-id="df6fd-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="df6fd-108">创建图表。</span><span class="sxs-lookup"><span data-stu-id="df6fd-108">Create chart.</span></span>
* <span data-ttu-id="df6fd-109">获取图表和表格图像。</span><span class="sxs-lookup"><span data-stu-id="df6fd-109">Get chart and table images.</span></span>
* <span data-ttu-id="df6fd-110">使用电子邮件向图像Power Automate。</span><span class="sxs-lookup"><span data-stu-id="df6fd-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="df6fd-111">_输入数据_</span><span class="sxs-lookup"><span data-stu-id="df6fd-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="显示输入数据表格的工作表。":::

<span data-ttu-id="df6fd-113">_输出图表_</span><span class="sxs-lookup"><span data-stu-id="df6fd-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="创建的柱形图显示客户到期金额。":::

<span data-ttu-id="df6fd-115">_通过流收到Power Automate的电子邮件_</span><span class="sxs-lookup"><span data-stu-id="df6fd-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="流发送的电子邮件，Excel嵌入正文中的图表。":::

## <a name="solution"></a><span data-ttu-id="df6fd-117">解决方案</span><span class="sxs-lookup"><span data-stu-id="df6fd-117">Solution</span></span>

<span data-ttu-id="df6fd-118">此解决方案由两部分组成：</span><span class="sxs-lookup"><span data-stu-id="df6fd-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="df6fd-119">一Office用于计算和提取图表和Excel的脚本</span><span class="sxs-lookup"><span data-stu-id="df6fd-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="df6fd-120">一Power Automate调用脚本并通过电子邮件发送结果的流。</span><span class="sxs-lookup"><span data-stu-id="df6fd-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="df6fd-121">有关如何执行此操作的示例，请参阅使用 Power Automate 创建自动化[工作流](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="df6fd-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="df6fd-122">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="df6fd-122">Sample Excel file</span></span>

<span data-ttu-id="df6fd-123">下载 <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> 工作簿的工作簿。</span><span class="sxs-lookup"><span data-stu-id="df6fd-123">Download <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="df6fd-124">添加以下脚本以自己试用示例！</span><span class="sxs-lookup"><span data-stu-id="df6fd-124">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="df6fd-125">示例代码：计算并提取Excel图表和表</span><span class="sxs-lookup"><span data-stu-id="df6fd-125">Sample code: Calculate and extract Excel chart and table</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="df6fd-126">Power Automate流：通过电子邮件发送图表和表格图像</span><span class="sxs-lookup"><span data-stu-id="df6fd-126">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="df6fd-127">此流运行脚本，并通过电子邮件发送返回的图像。</span><span class="sxs-lookup"><span data-stu-id="df6fd-127">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="df6fd-128">创建新的即时 **云流**。</span><span class="sxs-lookup"><span data-stu-id="df6fd-128">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="df6fd-129">选择 **"手动触发流"，** 然后选择"创建 **"。**</span><span class="sxs-lookup"><span data-stu-id="df6fd-129">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="df6fd-130">使用 Run **脚本操作** 添加使用 **Excel Online (Business)** 连接器的新步骤。 </span><span class="sxs-lookup"><span data-stu-id="df6fd-130">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="df6fd-131">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="df6fd-131">Use the following values for the action:</span></span>
    * <span data-ttu-id="df6fd-132">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="df6fd-132">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="df6fd-133">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="df6fd-133">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="df6fd-134">**文件**：工作簿 ([选择器选项选择)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="df6fd-134">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="df6fd-135">**脚本**：脚本名称</span><span class="sxs-lookup"><span data-stu-id="df6fd-135">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="已完成的 Excel Online (Business) 连接器Power Automate。":::
1. <span data-ttu-id="df6fd-137">此示例使用 Outlook 作为电子邮件客户端。</span><span class="sxs-lookup"><span data-stu-id="df6fd-137">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="df6fd-138">可以使用支持的任何Power Automate连接器，但其余步骤假定你已选择Outlook。</span><span class="sxs-lookup"><span data-stu-id="df6fd-138">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="df6fd-139">添加一 **个新** 步骤，该步骤使用 **Office 365 Outlook** 连接器和 **V2** (发送) 操作。</span><span class="sxs-lookup"><span data-stu-id="df6fd-139">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="df6fd-140">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="df6fd-140">Use the following values for the action:</span></span>
    * <span data-ttu-id="df6fd-141">**目标**：测试电子邮件帐户 (或个人) </span><span class="sxs-lookup"><span data-stu-id="df6fd-141">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="df6fd-142">**主题**：请查看报告数据</span><span class="sxs-lookup"><span data-stu-id="df6fd-142">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="df6fd-143">对于" **正文** "字段，选择"代码视图 `</>` " () 并输入以下内容：</span><span class="sxs-lookup"><span data-stu-id="df6fd-143">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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
1. <span data-ttu-id="df6fd-145">保存流并试用。使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。</span><span class="sxs-lookup"><span data-stu-id="df6fd-145">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="df6fd-146">培训视频：提取图表和表格的图像和电子邮件图像</span><span class="sxs-lookup"><span data-stu-id="df6fd-146">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="df6fd-147">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/152GJyqc-Kw)。</span><span class="sxs-lookup"><span data-stu-id="df6fd-147">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
