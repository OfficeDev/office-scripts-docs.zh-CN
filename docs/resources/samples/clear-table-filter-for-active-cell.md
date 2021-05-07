---
title: 基于活动单元格位置清除表格列筛选器
description: 了解如何根据活动单元格位置清除表列筛选器。
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: bbca4adce1de2cfade2c4f84273bf0bc06b5cc4b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232499"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>基于活动单元格位置清除表格列筛选器

本示例根据活动单元格位置清除表格列筛选器。 该脚本检测单元格是否属于表格，确定表格列，并清除应用了表格的任何筛选器。

如果希望了解有关在清除筛选器之前如何保存筛选器 (并稍后重新应用) ，请参阅通过保存筛选器跨表移动行，这是一个更[](move-rows-across-tables.md)高级的示例。

_在清除列筛选器 (，请注意活动单元格)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="清除列筛选器之前的活动单元格":::

_清除列筛选器后_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="清除列筛选器后的活动单元格":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>示例代码：基于活动单元格清除表列筛选器

以下脚本基于活动单元格位置清除表格列筛选器，并可以应用于Excel文件。 为方便起见，你可以 <a href="table-with-filter.xlsx"> 下载并使用 </a>table-with-filter.xlsx。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, return/exit.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get table (since it is already determined that there is only
    // a single table part of the selection).
    const currentTable = tables[0];

    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    const entireCol = cell.getEntireColumn();
    const intersect = entireCol.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get column.
    const col = currentTable.getColumnByName(headerCellValue);

    // Clear filter.
    col.getFilter().clear();
}
```
