---
title: 创建工作簿目录
description: 了解如何创建包含每个工作表链接的内容表。
ms.date: 01/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5b158160ecb9ac29df547c6da6552e21c9875be3
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572512"
---
# <a name="create-a-workbook-table-of-contents"></a>创建工作簿目录

此示例演示如何为工作簿创建目录。 目录中的每个条目都是工作簿中某个工作表的超链接。

:::image type="content" source="../../images/table-of-contents-sample.png" alt-text="显示指向其他工作表链接的目录工作表。":::

## <a name="sample-excel-file"></a>示例 Excel 文件

下载现成工作簿 [ 的table-of-contents.xlsx](table-of-contents.xlsx) 。 添加以下脚本并自行试用示例！

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
