---
title: 删除表格列筛选器
description: 了解如何根据活动单元格位置清除表列筛选器。
ms.date: 07/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: e016f7f2af9e7553229f3b3b19007e011879de8e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572519"
---
# <a name="remove-table-column-filters"></a>删除表格列筛选器

此示例基于活动单元格位置从表列中删除筛选器。 该脚本检测单元格是否是表的一部分，确定表列，并清除对其应用的任何筛选器。

若要详细了解如何在清除筛选器之前保存筛选器 (并在以后) 重新应用，请参阅 [通过保存筛选器跨表移动行](move-rows-across-tables.md)，这是一个更高级的示例。

## <a name="sample-excel-file"></a>示例 Excel 文件

下载现成工作簿 [ 的table-with-filter.xlsx](table-with-filter.xlsx) 。 添加以下脚本以自行尝试示例！

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>示例代码：清除基于活动单元格的表列筛选器

以下脚本根据活动单元格位置清除表列筛选器，并可应用于具有表的任何 Excel 文件。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  const cell = workbook.getActiveCell();

  // Get the tables associated with that cell.
  // Since tables can't overlap, this will be one table at most.
  const currentTable = cell.getTables()[0];

  // If there is no table on the selection, end the script.
  if (!currentTable) {
    console.log("The selection is not in a table.");
    return;
  }

  // Get the table header above the current cell by referencing its column.
  const entireColumn = cell.getEntireColumn();
  const intersect = entireColumn.getIntersection(currentTable.getRange());
  const headerCellValue = intersect.getCell(0, 0).getValue() as string;

  // Get the TableColumn object matching that header.
  const tableColumn = currentTable.getColumnByName(headerCellValue);

  // Clear the filters on that table column.
  tableColumn.getFilter().clear();
}
```

## <a name="before-clearing-column-filter-notice-the-active-cell"></a>在清除列筛选器之前 (请注意活动单元格) 

:::image type="content" source="../../images/before-filter-applied.png" alt-text="清除列筛选器之前的活动单元格。":::

## <a name="after-clearing-column-filter"></a>清除列筛选器后

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="清除列筛选器后的活动单元格。":::
