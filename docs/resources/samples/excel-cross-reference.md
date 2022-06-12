---
title: 使用Power Automate交叉引用Excel文件
description: 了解如何使用Office脚本和Power Automate交叉引用和设置Excel文件的格式。
ms.date: 06/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 02c06b6376d3726b3e1b44255df14aa64be196ea
ms.sourcegitcommit: f5fc9146d5c096e3a580a3fa8f9714147c548df4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/12/2022
ms.locfileid: "66038670"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>使用Power Automate交叉引用Excel文件

此解决方案演示如何比较两个Excel文件中的数据，以查找差异。 它使用Office脚本分析数据，Power Automate在工作簿之间进行通信。

## <a name="example-scenario"></a>示例方案

你是一名事件协调器，负责为即将召开的会议安排演讲者。 将事件数据保存在一个电子表格中，将说话人注册保存在另一个电子表格中。 若要确保这两个工作簿保持同步，请使用包含Office脚本的流来突出显示任何潜在问题。

## <a name="sample-excel-files"></a>示例Excel文件

下载以下文件以获取示例的即用工作簿。

1. <a href="event-data.xlsx">event-data.xlsx</a>
1. <a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a>

添加以下脚本以自行尝试示例！

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

## <a name="sample-code-validate-speaker-registrations"></a>示例代码：验证说话人注册

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Power Automate流：检查工作簿中的不一致性

此流从第一个工作簿中提取事件信息，并使用该数据来验证第二个工作簿。

1. 登录 [Power Automate](https://flow.microsoft.com)并创建新的 **即时云流**。
1. 选择 **“手动触发流** ”，然后选择 **“创建**”。
1. 添加使用 **Excel Online (Business)** 连接器和 **运行脚本** 操作 **的新步骤**。 对操作使用以下值。
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件**：event-data.xlsx ([使用文件选择器) 选中](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **脚本**：获取事件数据

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Power Automate中第一个脚本的已完成 Excel Online (Business) 连接器。":::

1. 添加第二个新 **步骤**，该步骤使用 **Excel Online (Business)** 连接器和 **运行脚本** 操作。 这会使用 **Get 事件数据** 脚本中返回的值作为 **验证事件数据** 脚本的输入。 对操作使用以下值。
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件**：speaker-registration.xlsx ([使用文件选择器) 选中](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **脚本**：验证说话人注册
    * **键**：运行 _**脚本**)  (动态内容_ 的结果

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Power Automate中第二个脚本的已完成 Excel Online (Business) 连接器。":::
1. 此示例使用Outlook作为电子邮件客户端。 可以使用任何电子邮件连接器Power Automate支持。 添加使用 **Office 365 Outlook** 连接器和 **发送和电子邮件 (V2)** 操作 **的新步骤**。 这会使用 **验证说话人注册** 脚本中返回的值作为电子邮件正文内容。 对操作使用以下值。
    * **收件** 人：测试电子邮件帐户 (或个人电子邮件) 
    * **主题**：事件验证结果
    * **正文**：运行 _**脚本 2** 中 (动态内容_ 的结果) 

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Power Automate中已完成的Office 365 Outlook连接器。":::
1. 保存流。 使用流编辑器页上的 **“测试** ”按钮，或通过“ **我的流** ”选项卡运行流。出现提示时，请务必允许访问。
1. 应收到一封电子邮件，指出“发现不匹配。 数据需要审阅。” 这表明speaker-registrations.xlsx中的行和 **event-data.xlsx** **中的行** 之间存在差异。 打开 **speaker-registrations.xlsx** ，查看几个突出显示的单元格，其中扬声器注册列表存在潜在问题。
