---
title: 从 Excel 工作表中的每个单元格中删除超链接
description: 了解如何使用 Office 脚本从 Excel 工作表中的每个单元格中删除超链接。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1445988b1e6a85fcab8914ffeaaef80a07a52f5e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572624"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>从 Excel 工作表中的每个单元格中删除超链接

 此示例清除当前工作表中的所有超链接。 它遍历工作表，如果有任何与单元格关联的超链接，它会清除超链接，但按原样保留单元格值。 另请记录完成遍历所需的时间。

> [!NOTE]
> 仅当单元格计数< 10k 时，这才有效。

## <a name="sample-excel-file"></a>示例 Excel 文件

下载文件 [remove-hyperlinks.xlsx](remove-hyperlinks.xlsx) ，以获取随时可用的工作簿。 添加以下脚本以自行尝试示例！

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>培训视频：从 Excel 工作表中的每个单元格中删除超链接

[观看苏迪 · 拉马穆尔西在 YouTube 上浏览这个示例](https://youtu.be/v20fdinxpHU)。
