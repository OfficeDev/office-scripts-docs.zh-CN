---
title: 基于活动单元格位置清除表格列筛选器
description: 了解如何根据活动单元格位置清除表列筛选器。
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: d6f267b433be9a0ddf44edf53ed92a136eb2ded6
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074436"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>基于活动单元格位置清除表格列筛选器

本示例根据活动单元格位置清除表格列筛选器。 该脚本检测单元格是否属于表格，确定表格列，并清除应用了表格的任何筛选器。

如果希望了解有关在清除筛选器之前如何保存筛选器 (并稍后重新应用) ，请参阅通过保存筛选器跨表移动行，这是一个更[](move-rows-across-tables.md)高级的示例。

_在清除列筛选器 (，请注意活动单元格)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="清除列筛选器之前的活动单元格。":::

_清除列筛选器后_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="清除列筛选器后的活动单元格。":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>示例代码：基于活动单元格清除表列筛选器

以下脚本基于活动单元格位置清除表格列筛选器，并可以应用于Excel文件。 为方便起见，你可以 <a href="table-with-filter.xlsx"> 下载并使用 </a>table-with-filter.xlsx。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, end the script.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get the first table associated with the active cell.
    const currentTable = tables[0];

    // Log key information about the table.
    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    // Get the table header above the current cell by referencing its column.
    const entireColumn = cell.getEntireColumn();
    const intersect = entireColumn.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get the TableColumn object matching that header.
    const tableColumn = currentTable.getColumnByName(headerCellValue);

    // Clear the filter on that table column.
    tableColumn.getFilter().clear();
}
```
