---
title: 在 Excel 中记录日常更改，并使用 Power Automate 流报告这些更改
description: 了解如何使用 Office 脚本和 Power Automate 跟踪工作簿中的值更改
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 083ca08573db060aa4788aea58fc67e50d004a4b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572653"
---
# <a name="record-day-to-day-changes-in-excel-and-report-them-with-a-power-automate-flow"></a>在 Excel 中记录日常更改，并使用 Power Automate 流报告这些更改

Power Automate 和 Office 脚本组合在一起，可为你处理重复性任务。 在此示例中，你的任务是每天在工作簿中录制单个数值读取，并报告自昨天起的更改。 你将生成一个流来获取该读取，将其记录在工作簿中，并通过电子邮件报告更改。

## <a name="sample-excel-file"></a>示例 Excel 文件

下载现成工作簿 [ 的daily-readings.xlsx](daily-readings.xlsx) 。 添加以下脚本以自行尝试示例！

## <a name="sample-code-record-and-report-daily-readings"></a>示例代码：记录和报告每日读数

```TypeScript
function main(workbook: ExcelScript.Workbook, newData: string): string {
  // Get the table by its name.
  const table = workbook.getTable("ReadingTable");

  // Read the current last entry in the Reading column.
  const readingColumn = table.getColumnByName("Reading");
  const readingColumnValues = readingColumn.getRange().getValues();
  const previousValue = readingColumnValues[readingColumnValues.length - 1][0] as number;

  // Add a row with the date, new value, and a formula calculating the difference.
  const currentDate = new Date(Date.now()).toLocaleDateString();
  const newRow = [currentDate, newData, "=[@Reading]-OFFSET([@Reading],-1,0)"];
  table.addRow(-1, newRow,);

  // Return the difference between the newData and the previous entry.
  const difference = Number.parseFloat(newData) - previousValue;
  console.log(difference);
  return difference;
}
```

## <a name="sample-flow-report-day-to-day-changes"></a>示例流：报告日常更改

按照以下步骤为示例生成 [Power Automate](https://powerautomate.microsoft.com/) 流。

1. 创建新的 **计划云流**。
1. 计划每 **1 天重复一次** 流。

    :::image type="content" source="../../images/day-to-day-changes-flow-1.png" alt-text="显示它每天重复的流创建步骤。":::
1. 选择“**创建**”。
1. 在实际流中，你将添加一个获取数据的步骤。 数据可以来自另一个工作簿、Teams 自适应卡片或任何其他源。 若要测试示例，请创建一个测试编号。 使用 **Initialize 变量** 操作添加新步骤。 为其提供以下值。
    1. **名称**：输入
    1. **类型**：整数
    1. **值**：190000

    :::image type="content" source="../../images/day-to-day-changes-flow-2.png" alt-text="使用给定值初始化变量操作。":::
1. 使用“**运行脚本**”操作通过 **Excel Online (Business)** 连接器添加新步骤。 对操作使用以下值。
    1. **位置**：OneDrive for Business
    1. **文档库**：OneDrive
    1. **文件**：通过 *文件浏览器) 选择daily-readings.xlsx (*
    1. **脚本**：脚本名称
    1. **newData**：输入 *(动态内容)*

    :::image type="content" source="../../images/day-to-day-changes-flow-3.png" alt-text="具有给定值的运行脚本操作。":::
1. 该脚本将每日读取差异作为名为“result”的动态内容返回。 对于示例，可以将信息发送给自己。 创建一个新步骤，该步骤将 **Outlook** 连接器与 **发送电子邮件 (V2)** 操作 (或你喜欢) 的任何电子邮件客户端一起使用。 使用以下值完成操作。
    1. **收件** 人：电子邮件地址
    1. **主题**：每日阅读更改
    1. **正文**：“与昨天的区别”结果 *(来自 Excel 的动态内容)*

    :::image type="content" source="../../images/day-to-day-changes-flow-4.png" alt-text="Power Automate 中已完成的 Outlook 连接器。":::
1. 保存流并试用。使用流编辑器页上的 **“测试** ”按钮。 出现提示时，请务必允许访问。
