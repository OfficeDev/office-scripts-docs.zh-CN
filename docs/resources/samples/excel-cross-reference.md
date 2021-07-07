---
title: 跨引用Excel文件Power Automate
description: 了解如何使用脚本Office脚本Power Automate交叉引用和格式化Excel文件。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 0776ce49cacecfa15339cc7c0cd4866daad789ff
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313958"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="e7863-103">跨引用Excel文件Power Automate</span><span class="sxs-lookup"><span data-stu-id="e7863-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="e7863-104">此解决方案演示如何比较两个文件之间的数据Excel查找差异。</span><span class="sxs-lookup"><span data-stu-id="e7863-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="e7863-105">它Office脚本来分析数据，Power Automate工作簿之间进行通信。</span><span class="sxs-lookup"><span data-stu-id="e7863-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="e7863-106">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="e7863-106">Example scenario</span></span>

<span data-ttu-id="e7863-107">你是安排即将召开的会议的演讲者的事件协调人。</span><span class="sxs-lookup"><span data-stu-id="e7863-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="e7863-108">您将事件数据保留在另一个电子表格中，将扬声器注册保留在另一个电子表格中。</span><span class="sxs-lookup"><span data-stu-id="e7863-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="e7863-109">若要确保两个工作簿保持同步，请对脚本Office流来突出显示任何潜在的问题。</span><span class="sxs-lookup"><span data-stu-id="e7863-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="e7863-110">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="e7863-110">Sample Excel files</span></span>

<span data-ttu-id="e7863-111">下载以下文件，获取示例的现成工作簿。</span><span class="sxs-lookup"><span data-stu-id="e7863-111">Download the following files to get ready-to-use workbooks for the sample.</span></span>

1. <span data-ttu-id="e7863-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="e7863-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="e7863-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="e7863-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

<span data-ttu-id="e7863-114">添加以下脚本以自己试用示例！</span><span class="sxs-lookup"><span data-stu-id="e7863-114">Add the following scripts to try the sample yourself!</span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="e7863-115">示例代码：获取事件数据</span><span class="sxs-lookup"><span data-stu-id="e7863-115">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="e7863-116">示例代码：验证扬声器注册</span><span class="sxs-lookup"><span data-stu-id="e7863-116">Sample code: Validate speaker registrations</span></span>

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="e7863-117">Power Automate流：检查工作簿之间的不一致情况</span><span class="sxs-lookup"><span data-stu-id="e7863-117">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="e7863-118">此流提取第一个工作簿的事件信息，并使用该数据验证第二个工作簿。</span><span class="sxs-lookup"><span data-stu-id="e7863-118">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="e7863-119">登录到 [Power Automate](https://flow.microsoft.com)并创建新的 **即时云流**。</span><span class="sxs-lookup"><span data-stu-id="e7863-119">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="e7863-120">选择 **"手动触发流"，** 然后选择"创建 **"。**</span><span class="sxs-lookup"><span data-stu-id="e7863-120">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="e7863-121">使用 Run **脚本操作** 添加使用 **Excel Online (Business)** 连接器的新步骤。 </span><span class="sxs-lookup"><span data-stu-id="e7863-121">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="e7863-122">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="e7863-122">Use the following values for the action:</span></span>
    * <span data-ttu-id="e7863-123">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="e7863-123">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="e7863-124">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="e7863-124">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="e7863-125">**文件**：event-data.xlsx ([文件选择器选项](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)) </span><span class="sxs-lookup"><span data-stu-id="e7863-125">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="e7863-126">**脚本**：获取事件数据</span><span class="sxs-lookup"><span data-stu-id="e7863-126">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="已完成的 Excel Online (Business) 连接器，用于 Power Automate。":::

1. <span data-ttu-id="e7863-128">通过运行脚本 **操作** 添加第二个使用 **Excel Online (Business)** 连接器 **的新** 步骤。</span><span class="sxs-lookup"><span data-stu-id="e7863-128">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="e7863-129">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="e7863-129">Use the following values for the action:</span></span>
    * <span data-ttu-id="e7863-130">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="e7863-130">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="e7863-131">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="e7863-131">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="e7863-132">**文件**：speaker-registration.xlsx ([文件选择器选项](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)) </span><span class="sxs-lookup"><span data-stu-id="e7863-132">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="e7863-133">**脚本**：验证扬声器注册</span><span class="sxs-lookup"><span data-stu-id="e7863-133">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="已完成的 Excel Online (Business) 连接器，用于第二个脚本Power Automate。":::
1. <span data-ttu-id="e7863-135">此示例使用 Outlook 作为电子邮件客户端。</span><span class="sxs-lookup"><span data-stu-id="e7863-135">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="e7863-136">可以使用任何支持的电子邮件Power Automate连接器。</span><span class="sxs-lookup"><span data-stu-id="e7863-136">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="e7863-137">添加一 **个新** 步骤，该步骤使用 **Office 365 Outlook** 连接器和 **V2** (发送) 操作。</span><span class="sxs-lookup"><span data-stu-id="e7863-137">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="e7863-138">对操作使用以下值：</span><span class="sxs-lookup"><span data-stu-id="e7863-138">Use the following values for the action:</span></span>
    * <span data-ttu-id="e7863-139">**目标**：测试电子邮件帐户 (或个人) </span><span class="sxs-lookup"><span data-stu-id="e7863-139">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="e7863-140">**主题**：事件验证结果</span><span class="sxs-lookup"><span data-stu-id="e7863-140">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="e7863-141">**正文**：结果 (_运行脚本 **2 中的**_ 动态) </span><span class="sxs-lookup"><span data-stu-id="e7863-141">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="已完成的Office 365 Outlook连接器Power Automate。":::
1. <span data-ttu-id="e7863-143">保存流。</span><span class="sxs-lookup"><span data-stu-id="e7863-143">Save the flow.</span></span> <span data-ttu-id="e7863-144">使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。</span><span class="sxs-lookup"><span data-stu-id="e7863-144">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>
1. <span data-ttu-id="e7863-145">你应该收到一封电子邮件，指出"发现不匹配。</span><span class="sxs-lookup"><span data-stu-id="e7863-145">You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="e7863-146">数据需要你审查。"</span><span class="sxs-lookup"><span data-stu-id="e7863-146">Data requires your review."</span></span> <span data-ttu-id="e7863-147">这表示行中的行与 **speaker-registrations.xlsx行之间存在\*\*\*\*event-data.xlsx。**</span><span class="sxs-lookup"><span data-stu-id="e7863-147">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="e7863-148">打开 **speaker-registrations.xlsx** 以查看一些突出显示的单元格，其中扬声器注册列表存在潜在问题。</span><span class="sxs-lookup"><span data-stu-id="e7863-148">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
