---
title: 从工作表的每个单元格中删除Excel超链接
description: 了解如何使用脚本Office工作表中每个单元格删除Excel超链接。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: dc33eb639edac8ada29824a53440031942e59179
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313748"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>从工作表的每个单元格中删除Excel超链接

 本示例清除当前工作表的所有超链接。 它会遍历工作表，如果存在与单元格关联的任何超链接，它会清除超链接，但会保留单元格值。 还记录完成遍历所花的时间。

> [!NOTE]
> 这仅在单元格计数为 10，000<有效。

## <a name="sample-excel-file"></a>示例Excel文件

下载适用于 <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> 工作簿的文件文件。 添加以下脚本以自己试用示例！

## <a name="sample-code-remove-hyperlinks"></a>示例代码：删除超链接

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {
  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);

  // Get the used range to operate on.
  // For large ranges (over 10000 entries), consider splitting the operation into batches for performance.
  const targetRange = sheet.getUsedRange(true);
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);

  // Go through each individual cell looking for a hyperlink. 
  // This allows us to limit the formatting changes to only the cells with hyperlink formatting.
  let clearedCount = 0;
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      const cell = targetRange.getCell(i, j);
      const hyperlink = cell.getHyperlink();
      if (hyperlink) {
        cell.clear(ExcelScript.ClearApplyTo.hyperlinks);
        cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        cell.getFormat().getFont().setColor('Black');
        clearedCount++;
      }
    }
  }

  console.log(`Done. Cleared hyperlinks from ${clearedCount} cells`);
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>培训视频：从工作表中的每个单元格Excel超链接

[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/v20fdinxpHU)。
