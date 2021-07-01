---
title: 跨引用Excel文件Power Automate
description: 了解如何使用脚本Office脚本Power Automate交叉引用和格式化Excel文件。
ms.date: 06/25/2021
localization_priority: Normal
ms.openlocfilehash: 89c4a5fa5dcff21681fa20cd4118447d39d9b6da
ms.sourcegitcommit: a063b3faf6c1b7c294bd6a73e46845b352f2a22d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/29/2021
ms.locfileid: "53202864"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="f889f-103">跨引用Excel文件Power Automate</span><span class="sxs-lookup"><span data-stu-id="f889f-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="f889f-104">此解决方案演示如何比较两个文件之间的数据Excel查找差异。</span><span class="sxs-lookup"><span data-stu-id="f889f-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="f889f-105">它Office脚本来分析数据，Power Automate工作簿之间进行通信。</span><span class="sxs-lookup"><span data-stu-id="f889f-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="f889f-106">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="f889f-106">Example scenario</span></span>

<span data-ttu-id="f889f-107">你是安排即将召开的会议的演讲者的事件协调人。</span><span class="sxs-lookup"><span data-stu-id="f889f-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="f889f-108">您将事件数据保留在另一个电子表格中，将扬声器注册保留在另一个电子表格中。</span><span class="sxs-lookup"><span data-stu-id="f889f-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="f889f-109">若要确保两个工作簿保持同步，请对脚本Office流来突出显示任何潜在的问题。</span><span class="sxs-lookup"><span data-stu-id="f889f-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="f889f-110">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="f889f-110">Sample Excel files</span></span>

<span data-ttu-id="f889f-111">下载此解决方案中使用的以下文件，以尝试一下！</span><span class="sxs-lookup"><span data-stu-id="f889f-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="f889f-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="f889f-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="f889f-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="f889f-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="f889f-114">示例代码：获取事件数据</span><span class="sxs-lookup"><span data-stu-id="f889f-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="f889f-115">示例代码：验证扬声器注册</span><span class="sxs-lookup"><span data-stu-id="f889f-115">Sample code: Validate speaker registrations</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");
    }
  }

  

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="f889f-116">Power Automate流：检查工作簿之间的不一致情况</span><span class="sxs-lookup"><span data-stu-id="f889f-116">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="f889f-117">此流提取第一个工作簿的事件信息，并使用该数据验证第二个工作簿。</span><span class="sxs-lookup"><span data-stu-id="f889f-117">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="f889f-118">登录到 [Power Automate](https://flow.microsoft.com)并创建新的 **即时云流**。</span><span class="sxs-lookup"><span data-stu-id="f889f-118">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="f889f-119">选择 **"手动触发流"，** 然后按"**创建"。**</span><span class="sxs-lookup"><span data-stu-id="f889f-119">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="f889f-120">使用 Run **脚本操作** 添加使用 **Excel Online (Business)** 连接器的新步骤。 </span><span class="sxs-lookup"><span data-stu-id="f889f-120">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="f889f-121">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="f889f-121">Use the following values for the action:</span></span>
    * <span data-ttu-id="f889f-122">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="f889f-122">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="f889f-123">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="f889f-123">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="f889f-124">**文件**：event-data.xlsx ([文件选择器选项](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)) </span><span class="sxs-lookup"><span data-stu-id="f889f-124">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="f889f-125">**脚本**：获取事件数据</span><span class="sxs-lookup"><span data-stu-id="f889f-125">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="已完成的 Excel Online (Business) 连接器，用于 Power Automate。":::

1. <span data-ttu-id="f889f-127">通过运行脚本 **操作** 添加第二个使用 **Excel Online (Business)** 连接器 **的新** 步骤。</span><span class="sxs-lookup"><span data-stu-id="f889f-127">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="f889f-128">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="f889f-128">Use the following values for the action:</span></span>
    * <span data-ttu-id="f889f-129">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="f889f-129">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="f889f-130">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="f889f-130">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="f889f-131">**文件**：speaker-registration.xlsx ([文件选择器选项](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)) </span><span class="sxs-lookup"><span data-stu-id="f889f-131">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="f889f-132">**脚本**：验证扬声器注册</span><span class="sxs-lookup"><span data-stu-id="f889f-132">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="已完成的 Excel Online (Business) 连接器，用于第二个脚本Power Automate。":::
1. <span data-ttu-id="f889f-134">此示例使用 Outlook 作为电子邮件客户端。</span><span class="sxs-lookup"><span data-stu-id="f889f-134">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="f889f-135">可以使用任何支持的电子邮件Power Automate连接器。</span><span class="sxs-lookup"><span data-stu-id="f889f-135">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="f889f-136">添加一 **个新** 步骤，该步骤使用 **Office 365 Outlook** 连接器和 **V2** (发送) 操作。</span><span class="sxs-lookup"><span data-stu-id="f889f-136">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="f889f-137">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="f889f-137">Use the following values for the action:</span></span>
    * <span data-ttu-id="f889f-138">**目标**：测试电子邮件帐户 (或个人) </span><span class="sxs-lookup"><span data-stu-id="f889f-138">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="f889f-139">**主题**：事件验证结果</span><span class="sxs-lookup"><span data-stu-id="f889f-139">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="f889f-140">**正文**：结果 (_运行脚本 **2 中的**_ 动态) </span><span class="sxs-lookup"><span data-stu-id="f889f-140">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="已完成的Office 365 Outlook连接器Power Automate。":::
1. <span data-ttu-id="f889f-142">保存流，然后选择" **测试** "以试用。你应该收到一封电子邮件，指出"发现不匹配。</span><span class="sxs-lookup"><span data-stu-id="f889f-142">Save the flow, then select **Test** to try it out. You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="f889f-143">数据需要你审查。"</span><span class="sxs-lookup"><span data-stu-id="f889f-143">Data requires your review."</span></span> <span data-ttu-id="f889f-144">这表示行中的行与 **speaker-registrations.xlsx行之间存在\*\*\*\*event-data.xlsx。**</span><span class="sxs-lookup"><span data-stu-id="f889f-144">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="f889f-145">打开 **speaker-registrations.xlsx** 以查看一些突出显示的单元格，其中扬声器注册列表存在潜在问题。</span><span class="sxs-lookup"><span data-stu-id="f889f-145">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
