---
title: 基于活动单元格位置清除表格列筛选器
description: 了解如何根据活动单元格位置清除表列筛选器。
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: 5da3cdfe753f254c7f9ec97bea384b77ccaff09e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571122"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="e9482-103">基于活动单元格位置清除表格列筛选器</span><span class="sxs-lookup"><span data-stu-id="e9482-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="e9482-104">本示例根据活动单元格位置清除表格列筛选器。</span><span class="sxs-lookup"><span data-stu-id="e9482-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="e9482-105">该脚本检测单元格是否属于表格，确定表格列，并清除应用了表格的任何筛选器。</span><span class="sxs-lookup"><span data-stu-id="e9482-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="e9482-106">如果希望了解有关在清除筛选器之前如何保存筛选器 (并稍后重新应用) ，请参阅通过保存筛选器跨表移动行，这是一个更[](move-rows-across-tables.md)高级的示例。</span><span class="sxs-lookup"><span data-stu-id="e9482-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="e9482-107">_在清除列筛选器 (，请注意活动单元格)_</span><span class="sxs-lookup"><span data-stu-id="e9482-107">_Before clearing column filter (notice the active cell)_</span></span>

![清除列筛选器之前](../../images/before-filter-applied.png)

<span data-ttu-id="e9482-109">_清除列筛选器后_</span><span class="sxs-lookup"><span data-stu-id="e9482-109">_After clearing column filter_</span></span>

![清除列筛选器后](../../images/after-filter-cleared.png)

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="e9482-111">示例代码：基于活动单元格清除表列筛选器</span><span class="sxs-lookup"><span data-stu-id="e9482-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="e9482-112">以下脚本基于活动单元格位置清除表格列筛选器，并可以应用于任何包含表格的 Excel 文件。</span><span class="sxs-lookup"><span data-stu-id="e9482-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="e9482-113">为方便起见，你可以 <a href="table-with-filter.xlsx"> 下载并使用 </a>table-with-filter.xlsx。</span><span class="sxs-lookup"><span data-stu-id="e9482-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

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

## <a name="training-video-clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="e9482-114">培训视频：基于活动单元格位置清除表格列筛选器</span><span class="sxs-lookup"><span data-stu-id="e9482-114">Training video: Clear table column filter based on active cell location</span></span>

<span data-ttu-id="e9482-115">有关如何使用区域的示例，请参阅 [Range 基础知识培训视频](range-basics.md#training-videos-range-basics)。</span><span class="sxs-lookup"><span data-stu-id="e9482-115">For an example of how to work with ranges, see [Range basics training videos](range-basics.md#training-videos-range-basics).</span></span>