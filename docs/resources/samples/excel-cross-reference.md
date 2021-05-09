---
title: 交叉引用和格式化Excel文件
description: 了解如何使用脚本Office脚本Power Automate交叉引用和格式化Excel文件。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 7cc10787190e7ba8f5984ddda8b3c770eb0f7d8a
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285904"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="4e5de-103">交叉引用和格式化Excel文件</span><span class="sxs-lookup"><span data-stu-id="4e5de-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="4e5de-104">此解决方案演示如何使用脚本Excel脚本和脚本对两个Office文件进行交叉Power Automate。</span><span class="sxs-lookup"><span data-stu-id="4e5de-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="4e5de-105">项目实现以下目标：</span><span class="sxs-lookup"><span data-stu-id="4e5de-105">The project achieves the following:</span></span>

1. <span data-ttu-id="4e5de-106">使用一个 Run 脚本 <a href="events.xlsx">events.xlsx</a> 事件数据。</span><span class="sxs-lookup"><span data-stu-id="4e5de-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="4e5de-107">将该数据传递给包含事件事务数据的第二Excel文件，并使用该数据对数据进行基本验证，并使用 Office Scripts 对缺失或错误数据进行格式设置。</span><span class="sxs-lookup"><span data-stu-id="4e5de-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="4e5de-108">将结果通过电子邮件发送给审阅者。</span><span class="sxs-lookup"><span data-stu-id="4e5de-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="4e5de-109">有关更多详细信息，请参阅交叉引用和使用脚本[设置两Excel文件Office格式](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)。</span><span class="sxs-lookup"><span data-stu-id="4e5de-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="4e5de-110">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="4e5de-110">Sample Excel files</span></span>

<span data-ttu-id="4e5de-111">下载此解决方案中使用的以下文件，以尝试一下！</span><span class="sxs-lookup"><span data-stu-id="4e5de-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="4e5de-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="4e5de-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="4e5de-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="4e5de-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="4e5de-114">示例代码：获取事件数据</span><span class="sxs-lookup"><span data-stu-id="4e5de-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): EventData[] {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];
  
  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
      let [event, date, location, capacity] = row;
      records.push({
          event: event as string,
          date: date as number, 
          location: location as string,
          capacity: capacity as number
      })
  }

  // Log the event data to the console and return it for a flow.
  console.log(JSON.stringify(records));
  return records;
}

// An interface representing a row of event data.
interface EventData {
  event: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="4e5de-115">示例代码：验证事件事务</span><span class="sxs-lookup"><span data-stu-id="4e5de-115">Sample code: Validate event transactions</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);
    
 // Apply some basic formatting for readability.
  table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
  table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let overallMatch = true;

  // Iterate over every row.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [event, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyObject of keysObject) {
      if (keyObject.event === event) {
        match = true;

        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (keyObject.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (keyObject.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (keyObject.capacity !== capacity) {
          overallMatch = false;
          range.getCell(i, 3).getFormat()
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
  }
  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  event: string
  date: number
  location: string
  capacity: number
}
```

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="4e5de-116">培训视频：交叉引用和格式化Excel文件</span><span class="sxs-lookup"><span data-stu-id="4e5de-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="4e5de-117">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/dVwqBf483qo")。</span><span class="sxs-lookup"><span data-stu-id="4e5de-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/dVwqBf483qo").</span></span>
