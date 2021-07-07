---
title: 对文件夹中的所有 Excel 文件运行脚本
description: 了解如何对 OneDrive for Business 上文件夹中的所有 Excel 文件运行OneDrive for Business。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: bf9c0c486dacced5c3017b267ea65dfd215a5197
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313895"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="a7448-103">对文件夹中的所有 Excel 文件运行脚本</span><span class="sxs-lookup"><span data-stu-id="a7448-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="a7448-104">此项目对位于 OneDrive for Business 上的文件夹中的所有文件执行一组自动化OneDrive for Business。</span><span class="sxs-lookup"><span data-stu-id="a7448-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="a7448-105">它还可用于文件夹SharePoint文件夹。</span><span class="sxs-lookup"><span data-stu-id="a7448-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="a7448-106">该代码对Excel文件执行计算，添加格式，并插入一个[@mentions注释。](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)</span><span class="sxs-lookup"><span data-stu-id="a7448-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="a7448-107">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="a7448-107">Sample Excel files</span></span>

<span data-ttu-id="a7448-108">下载 <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> 此示例需要的所有工作簿的工作簿。</span><span class="sxs-lookup"><span data-stu-id="a7448-108">Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> for all the workbooks you'll need for this sample.</span></span> <span data-ttu-id="a7448-109">将这些文件解压缩到名为"销售" **的文件夹中**。</span><span class="sxs-lookup"><span data-stu-id="a7448-109">Extract those files to a folder titled **Sales**.</span></span> <span data-ttu-id="a7448-110">将以下脚本添加到脚本集合，以自己试用示例！</span><span class="sxs-lookup"><span data-stu-id="a7448-110">Add the following script to your script collection to try the sample yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="a7448-111">示例代码：添加格式并插入注释</span><span class="sxs-lookup"><span data-stu-id="a7448-111">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="a7448-112">这是在每个单独的工作簿上运行的脚本。</span><span class="sxs-lookup"><span data-stu-id="a7448-112">This is the script that runs on each individual workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="a7448-113">Power Automate流：对文件夹内每个工作簿运行脚本</span><span class="sxs-lookup"><span data-stu-id="a7448-113">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="a7448-114">此流对"销售"文件夹中每个工作簿运行脚本。</span><span class="sxs-lookup"><span data-stu-id="a7448-114">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="a7448-115">创建新的即时 **云流**。</span><span class="sxs-lookup"><span data-stu-id="a7448-115">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="a7448-116">选择 **"手动触发流"，** 然后选择"创建 **"。**</span><span class="sxs-lookup"><span data-stu-id="a7448-116">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="a7448-117">添加一 **个新** 步骤，该步骤使用 **OneDrive for Business** 连接器和 **"在文件夹操作中列出文件**"。</span><span class="sxs-lookup"><span data-stu-id="a7448-117">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="已完成的OneDrive for Business连接器Power Automate。":::
1. <span data-ttu-id="a7448-119">选择包含提取的工作簿的"Sales"文件夹。</span><span class="sxs-lookup"><span data-stu-id="a7448-119">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="a7448-120">若要确保仅选择工作簿，请选择"**新建步骤"，\*\*\*\*然后选择"条件**"并设置以下值：</span><span class="sxs-lookup"><span data-stu-id="a7448-120">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="a7448-121">**文件名** (OneDrive文件名值) </span><span class="sxs-lookup"><span data-stu-id="a7448-121">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="a7448-122">"ends with"</span><span class="sxs-lookup"><span data-stu-id="a7448-122">"ends with"</span></span>
    1. <span data-ttu-id="a7448-123">"xlsx"。</span><span class="sxs-lookup"><span data-stu-id="a7448-123">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="The Power Automate condition block that applies subsequent actions to each file.":::
1. <span data-ttu-id="a7448-125">Under the **If yes** branch， add the Excel Online (**Business)** connector with the Run **script** action.</span><span class="sxs-lookup"><span data-stu-id="a7448-125">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="a7448-126">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="a7448-126">Use the following values for the action:</span></span>
    1. <span data-ttu-id="a7448-127">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="a7448-127">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="a7448-128">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="a7448-128">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="a7448-129">**文件\*\*\*\*： (** id OneDrive文件 ID 值) </span><span class="sxs-lookup"><span data-stu-id="a7448-129">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="a7448-130">**脚本**：脚本名称</span><span class="sxs-lookup"><span data-stu-id="a7448-130">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="已完成的 Excel Online (Business) 连接器Power Automate。":::
1. <span data-ttu-id="a7448-132">保存流并试用。使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。</span><span class="sxs-lookup"><span data-stu-id="a7448-132">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="a7448-133">培训视频：对文件夹中的所有Excel文件运行脚本</span><span class="sxs-lookup"><span data-stu-id="a7448-133">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="a7448-134">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/xMg711o7k6w)。</span><span class="sxs-lookup"><span data-stu-id="a7448-134">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
