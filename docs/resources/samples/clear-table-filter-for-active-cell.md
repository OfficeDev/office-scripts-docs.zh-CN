---
title: 基于活动单元格位置清除表格列筛选器
description: 了解如何根据活动单元格位置清除表列筛选器。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: f10e23b4ad948a28c5b749533ddedefe164d7142
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313888"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="89821-103">基于活动单元格位置清除表格列筛选器</span><span class="sxs-lookup"><span data-stu-id="89821-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="89821-104">本示例根据活动单元格位置清除表格列筛选器。</span><span class="sxs-lookup"><span data-stu-id="89821-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="89821-105">该脚本检测单元格是否属于表格，确定表格列，并清除应用了表格的任何筛选器。</span><span class="sxs-lookup"><span data-stu-id="89821-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="89821-106">如果希望了解有关在清除筛选器之前如何保存筛选器 (并稍后重新应用) ，请参阅通过保存筛选器跨表移动行，这是一个更[](move-rows-across-tables.md)高级的示例。</span><span class="sxs-lookup"><span data-stu-id="89821-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="89821-107">_在清除列筛选器 (，请注意活动单元格)_</span><span class="sxs-lookup"><span data-stu-id="89821-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="清除列筛选器之前的活动单元格。":::

<span data-ttu-id="89821-109">_清除列筛选器后_</span><span class="sxs-lookup"><span data-stu-id="89821-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="清除列筛选器后的活动单元格。":::

## <a name="sample-excel-file"></a><span data-ttu-id="89821-111">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="89821-111">Sample Excel file</span></span>

<span data-ttu-id="89821-112">下载 <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> 工作簿的工作簿。</span><span class="sxs-lookup"><span data-stu-id="89821-112">Download <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="89821-113">添加以下脚本以自己试用示例！</span><span class="sxs-lookup"><span data-stu-id="89821-113">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="89821-114">示例代码：基于活动单元格清除表列筛选器</span><span class="sxs-lookup"><span data-stu-id="89821-114">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="89821-115">以下脚本基于活动单元格位置清除表格列筛选器，并可以应用于Excel文件。</span><span class="sxs-lookup"><span data-stu-id="89821-115">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span>

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
