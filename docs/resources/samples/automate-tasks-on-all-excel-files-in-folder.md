---
title: 对文件夹中的所有 Excel 文件运行脚本
description: 了解如何对 OneDrive for Business 上文件夹中的所有 Excel 文件运行OneDrive for Business。
ms.date: 04/02/2021
localization_priority: Normal
ms.openlocfilehash: 6376dcac0eb36c04c2b60b2717d18cd730a0a8ee
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026839"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="d5318-103">对文件夹中的所有 Excel 文件运行脚本</span><span class="sxs-lookup"><span data-stu-id="d5318-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="d5318-104">此项目对位于 OneDrive for Business 上的文件夹中的所有文件执行一组自动化OneDrive for Business。</span><span class="sxs-lookup"><span data-stu-id="d5318-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="d5318-105">它还可用于文件夹SharePoint文件夹。</span><span class="sxs-lookup"><span data-stu-id="d5318-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="d5318-106">该代码对Excel文件执行计算，添加格式，并插入一个[@mentions注释。](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)</span><span class="sxs-lookup"><span data-stu-id="d5318-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="d5318-107">下载文件 <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip，</a>将文件解压缩到本示例中使用的名为 **Sales** 的文件夹，然后自己试用！</span><span class="sxs-lookup"><span data-stu-id="d5318-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="d5318-108">示例代码：添加格式并插入注释</span><span class="sxs-lookup"><span data-stu-id="d5318-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="d5318-109">这是在每个单独的工作簿上运行的脚本。</span><span class="sxs-lookup"><span data-stu-id="d5318-109">This is the script that runs on each individual workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="d5318-110">Power Automate流：对文件夹内每个工作簿运行脚本</span><span class="sxs-lookup"><span data-stu-id="d5318-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="d5318-111">此流对"销售"文件夹中每个工作簿运行脚本。</span><span class="sxs-lookup"><span data-stu-id="d5318-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="d5318-112">创建新的即时 **云流**。</span><span class="sxs-lookup"><span data-stu-id="d5318-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="d5318-113">选择 **"手动触发流"，** 然后按"**创建"。**</span><span class="sxs-lookup"><span data-stu-id="d5318-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="d5318-114">添加一 **个新** 步骤，该步骤使用 **OneDrive for Business** 连接器和 **"在文件夹操作中列出文件**"。</span><span class="sxs-lookup"><span data-stu-id="d5318-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="已完成的OneDrive for Business连接器Power Automate。":::
1. <span data-ttu-id="d5318-116">选择包含提取的工作簿的"Sales"文件夹。</span><span class="sxs-lookup"><span data-stu-id="d5318-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="d5318-117">若要确保仅选择工作簿，请选择"**新建步骤"，\*\*\*\*然后选择"条件**"并设置以下值：</span><span class="sxs-lookup"><span data-stu-id="d5318-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="d5318-118">**文件名** (OneDrive文件名值) </span><span class="sxs-lookup"><span data-stu-id="d5318-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="d5318-119">"ends with"</span><span class="sxs-lookup"><span data-stu-id="d5318-119">"ends with"</span></span>
    1. <span data-ttu-id="d5318-120">"xlsx"。</span><span class="sxs-lookup"><span data-stu-id="d5318-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="The Power Automate condition block that applies subsequent actions to each file.":::
1. <span data-ttu-id="d5318-122">Under the **If yes** branch， add the Excel Online (**Business)** connector with the Run script (**preview)** action.</span><span class="sxs-lookup"><span data-stu-id="d5318-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="d5318-123">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="d5318-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="d5318-124">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="d5318-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="d5318-125">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="d5318-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="d5318-126">**文件\*\*\*\*： (** id OneDrive文件 ID 值) </span><span class="sxs-lookup"><span data-stu-id="d5318-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="d5318-127">**脚本**：脚本名称</span><span class="sxs-lookup"><span data-stu-id="d5318-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="已完成的 Excel Online (Business) 连接器Power Automate。":::
1. <span data-ttu-id="d5318-129">保存流并试用。</span><span class="sxs-lookup"><span data-stu-id="d5318-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="d5318-130">培训视频：对文件夹中的所有Excel文件运行脚本</span><span class="sxs-lookup"><span data-stu-id="d5318-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="d5318-131">[观看分步视频，](https://youtu.be/xMg711o7k6w)了解如何对 Excel 或 SharePoint 文件夹中的所有 OneDrive for Business 文件运行脚本。</span><span class="sxs-lookup"><span data-stu-id="d5318-131">[Watch step-by-step video](https://youtu.be/xMg711o7k6w) on how to run a script on all Excel files in a OneDrive for Business or SharePoint folder.</span></span>
