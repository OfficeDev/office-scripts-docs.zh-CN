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
# <a name="cross-reference-excel-files-with-power-automate"></a>跨引用Excel文件Power Automate

此解决方案演示如何比较两个文件之间的数据Excel查找差异。 它Office脚本来分析数据，Power Automate工作簿之间进行通信。

## <a name="example-scenario"></a>示例应用场景

你是安排即将召开的会议的演讲者的事件协调人。 您将事件数据保留在另一个电子表格中，将扬声器注册保留在另一个电子表格中。 若要确保两个工作簿保持同步，请对脚本Office流来突出显示任何潜在的问题。

## <a name="sample-excel-files"></a>示例Excel文件

下载以下文件，获取示例的现成工作簿。

1. <a href="event-data.xlsx">event-data.xlsx</a>
1. <a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a>

添加以下脚本以自己试用示例！

## <a name="sample-code-get-event-data"></a>示例代码：获取事件数据

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

## <a name="sample-code-validate-speaker-registrations"></a>示例代码：验证扬声器注册

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Power Automate流：检查工作簿之间的不一致情况

此流提取第一个工作簿的事件信息，并使用该数据验证第二个工作簿。

1. 登录到 [Power Automate](https://flow.microsoft.com)并创建新的 **即时云流**。
1. 选择 **"手动触发流"，** 然后选择"创建 **"。**
1. 使用 Run **脚本操作** 添加使用 **Excel Online (Business)** 连接器的新步骤。  对操作使用以下值：
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件**：event-data.xlsx ([文件选择器选项](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)) 
    * **脚本**：获取事件数据

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="已完成的 Excel Online (Business) 连接器，用于 Power Automate。":::

1. 通过运行脚本 **操作** 添加第二个使用 **Excel Online (Business)** 连接器 **的新** 步骤。 对操作使用以下值：
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件**：speaker-registration.xlsx ([文件选择器选项](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)) 
    * **脚本**：验证扬声器注册

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="已完成的 Excel Online (Business) 连接器，用于第二个脚本Power Automate。":::
1. 此示例使用 Outlook 作为电子邮件客户端。 可以使用任何支持的电子邮件Power Automate连接器。 添加一 **个新** 步骤，该步骤使用 **Office 365 Outlook** 连接器和 **V2** (发送) 操作。 对操作使用以下值：
    * **目标**：测试电子邮件帐户 (或个人) 
    * **主题**：事件验证结果
    * **正文**：结果 (_运行脚本 **2 中的**_ 动态) 

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="已完成的Office 365 Outlook连接器Power Automate。":::
1. 保存流。 使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。
1. 你应该收到一封电子邮件，指出"发现不匹配。 数据需要你审查。" 这表示行中的行与 **speaker-registrations.xlsx行之间存在****event-data.xlsx。** 打开 **speaker-registrations.xlsx** 以查看一些突出显示的单元格，其中扬声器注册列表存在潜在问题。
