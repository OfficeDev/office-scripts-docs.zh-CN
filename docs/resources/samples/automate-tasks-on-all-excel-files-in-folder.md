---
title: 对文件夹中的所有 Excel 文件运行脚本
description: 了解如何对 OneDrive for Business 上文件夹中的所有 Excel 文件运行脚本。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: fad1483fbcddaf541874630e8a4e5a06faa784627d44d17ea2ab7ca0af1550a4
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847402"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>对文件夹中的所有 Excel 文件运行脚本

此项目对位于 OneDrive for Business 上的文件夹中的所有文件执行一组自动化OneDrive for Business。 它还可用于文件夹SharePoint文件夹。
该代码对Excel文件执行计算，添加格式，并插入一个[@mentions注释。](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)

## <a name="sample-excel-files"></a>示例Excel文件

下载 <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> 此示例需要的所有工作簿的工作簿。 将这些文件解压缩到名为"销售" **的文件夹中**。 将以下脚本添加到脚本集合，以自己试用示例！

## <a name="sample-code-add-formatting-and-insert-comment"></a>示例代码：添加格式并插入注释

这是在每个单独的工作簿上运行的脚本。

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate流：对文件夹内每个工作簿运行脚本

此流对"销售"文件夹中每个工作簿运行脚本。

1. 创建新的即时 **云流**。
1. 选择 **"手动触发流"，** 然后选择"创建 **"。**
1. 添加一 **个新** 步骤，该步骤使用 **OneDrive for Business** 连接器和 **"在文件夹操作中列出文件**"。

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="已完成的OneDrive for Business连接器Power Automate。":::
1. 选择包含提取的工作簿的"Sales"文件夹。
1. 若要确保仅选择工作簿，请选择"**新建步骤"，** 然后选择"条件 **"。** 对条件使用以下值。
    1. **文件名** (OneDrive文件名值) 
    1. "ends with"
    1. "xlsx"

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="The Power Automate condition block that applies subsequent actions to each file.":::
1. Under the **If yes** branch， add the Excel Online (**Business)** connector with the Run **script** action. 对操作使用以下值。
    1. **位置**：OneDrive for Business
    1. **文档库**：OneDrive
    1. **文件****： (** id OneDrive文件 ID 值) 
    1. **脚本**：脚本名称

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="已完成的 Excel Online (Business) 连接器Power Automate。":::
1. 保存流并试用。使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>培训视频：对文件夹中的所有Excel文件运行脚本

[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/xMg711o7k6w)。
