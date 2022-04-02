---
title: 基于活动单元格位置清除表格列筛选器
description: 了解如何根据活动单元格位置清除表列筛选器。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: c52f1a3501318a479744abc6f2aa15cfaf3f9ded
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585581"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>基于活动单元格位置清除表格列筛选器

本示例根据活动单元格位置清除表格列筛选器。 该脚本检测单元格是否属于表格，确定表格列，并清除应用了表格的任何筛选器。

如果希望了解有关在清除筛选器之前如何保存筛选器 (并稍后重新应用) ，请参阅通过保存筛选器跨表移动行，这是一个更高级的示例[](move-rows-across-tables.md)。

_在清除列筛选器 (，请注意活动单元格)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="清除列筛选器之前的活动单元格。":::

_清除列筛选器后_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="清除列筛选器后的活动单元格。":::

## <a name="sample-excel-file"></a>示例Excel文件

下载 <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> 工作簿的工作簿。 添加以下脚本以自己试用示例！

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>示例代码：基于活动单元格清除表列筛选器

以下脚本基于活动单元格位置清除表列筛选器，并可以应用于任何包含Excel文件。

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
