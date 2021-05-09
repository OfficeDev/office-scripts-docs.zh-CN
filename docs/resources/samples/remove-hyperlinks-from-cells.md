---
title: 从工作表的每个单元格中删除Excel超链接
description: 了解如何使用脚本Office工作表中每个单元格删除Excel超链接。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 048d01691377a7086bdba9ceb87ca98839cfa4d1
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285799"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="a8718-103">从工作表的每个单元格中删除Excel超链接</span><span class="sxs-lookup"><span data-stu-id="a8718-103">Remove hyperlinks from each cell in an Excel worksheet</span></span>

 <span data-ttu-id="a8718-104">本示例清除当前工作表的所有超链接。</span><span class="sxs-lookup"><span data-stu-id="a8718-104">This sample clears all of the hyperlinks from the current worksheet.</span></span> <span data-ttu-id="a8718-105">它会遍历工作表，如果存在与单元格关联的任何超链接，它会清除超链接，但会保留单元格值。</span><span class="sxs-lookup"><span data-stu-id="a8718-105">It traverses the worksheet and if there is any hyperlink associated with the cell, it clears the hyperlink yet retains the cell value as is.</span></span> <span data-ttu-id="a8718-106">还记录完成遍历所花的时间。</span><span class="sxs-lookup"><span data-stu-id="a8718-106">Also logs the time it takes to complete traversal.</span></span>

> [!NOTE]
> <span data-ttu-id="a8718-107">这仅在单元格计数为 10，000<有效。</span><span class="sxs-lookup"><span data-stu-id="a8718-107">This only works if the cell count is < 10k.</span></span>

## <a name="sample-code-remove-hyperlinks"></a><span data-ttu-id="a8718-108">示例代码：删除超链接</span><span class="sxs-lookup"><span data-stu-id="a8718-108">Sample code: Remove hyperlinks</span></span>

<span data-ttu-id="a8718-109">下载此示例 <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> 使用的文件，然后自己试用！</span><span class="sxs-lookup"><span data-stu-id="a8718-109">Download the file <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> used in this sample and try it out yourself!</span></span>

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="a8718-110">培训视频：从工作表中的每个单元格Excel超链接</span><span class="sxs-lookup"><span data-stu-id="a8718-110">Training video: Remove hyperlinks from each cell in an Excel worksheet</span></span>

<span data-ttu-id="a8718-111">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/v20fdinxpHU)。</span><span class="sxs-lookup"><span data-stu-id="a8718-111">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/v20fdinxpHU).</span></span>
