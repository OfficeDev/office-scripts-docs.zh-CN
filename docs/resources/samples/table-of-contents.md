---
title: 创建工作簿目录
description: 了解如何创建包含指向每个工作表的链接的目录。
ms.date: 01/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: b2d69609514c2e1e87f9c0590ea10152fc7d5e7d
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585518"
---
# <a name="create-a-workbook-table-of-contents"></a>创建工作簿目录

此示例演示如何为工作簿创建目录。 目录的每个条目都是指向工作簿中工作表之一的超链接。

:::image type="content" source="../../images/table-of-contents-sample.png" alt-text="显示指向其他工作表的链接的目录工作表。":::

## <a name="sample-excel-file"></a>示例Excel文件

下载 <a href="table-of-contents.xlsx">table-of-contents.xlsx</a> 工作簿的工作簿。 添加以下脚本，然后自己尝试示例！

## <a name="sample-code-create-a-workbook-table-of-contents"></a>示例代码：创建工作簿目录

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Insert a new worksheet at the beginning of the workbook.
  let tocSheet = workbook.addWorksheet();
  tocSheet.setPosition(0);
  tocSheet.setName("Table of Contents");

  // Give the worksheet a title in the sheet.
  tocSheet.getRange("A1").setValue("Table of Contents");
  tocSheet.getRange("A1").getFormat().getFont().setBold(true);

  // Create the table of contents headers.
  let tocRange = tocSheet.getRange("A2:B2")
  tocRange.setValues([["#", "Name"]]);

  // Get the range for the table of contents entries.
  let worksheets = workbook.getWorksheets();
  tocRange = tocRange.getResizedRange(worksheets.length, 0);

  // Loop through all worksheets in the workbook, except the first one.
  for (let i = 1; i < worksheets.length; i++) {
    // Create a row for each worksheet with its index and linked name.
    tocRange.getCell(i, 0).setValue(i);
    tocRange.getCell(i, 1).setHyperlink({
      textToDisplay: worksheets[i].getName(),
      documentReference: `'${worksheets[i].getName()}'!A1`
    });
  };

  // Activate the table of contents worksheet.
  tocSheet.activate();
}
```
