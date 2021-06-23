---
title: 使用脚本跨表Office行
description: 了解如何通过保存筛选器，然后处理和重新应用筛选器来跨表移动行。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: c850ed055457f6733694027469a96a87e74ef66a
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074450"
---
# <a name="move-rows-across-tables-by-saving-filters-then-processing-and-reapplying-the-filters"></a><span data-ttu-id="89273-103">通过保存筛选器，然后处理和重新应用筛选器，跨表移动行</span><span class="sxs-lookup"><span data-stu-id="89273-103">Move rows across tables by saving filters, then processing and reapplying the filters</span></span>

<span data-ttu-id="89273-104">此脚本执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="89273-104">This script does the following:</span></span>

* <span data-ttu-id="89273-105">从源表中选择行，其中列中的值等于 _某个值_。</span><span class="sxs-lookup"><span data-stu-id="89273-105">Selects rows from the source table where the value in a column is equal to _some value_.</span></span>
* <span data-ttu-id="89273-106">将所有选定的行移动到另一 (工作表) 中的目标行。</span><span class="sxs-lookup"><span data-stu-id="89273-106">Moves all selected rows into another (target) table on another worksheet.</span></span>
* <span data-ttu-id="89273-107">重新应用源表上的相关筛选器。</span><span class="sxs-lookup"><span data-stu-id="89273-107">Reapplies the relevant filters on the source table.</span></span>

:::image type="content" source="../../images/table-filter-before-after.png" alt-text="工作簿之前和之后屏幕截图。":::

## <a name="sample-excel-file"></a><span data-ttu-id="89273-109">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="89273-109">Sample Excel file</span></span>

<span data-ttu-id="89273-110">下载此 <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> 中使用的文件，以尝试一下！</span><span class="sxs-lookup"><span data-stu-id="89273-110">Download the file <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> used in this solution to try it out yourself!</span></span>

## <a name="sample-code-move-rows-using-range-values"></a><span data-ttu-id="89273-111">示例代码：使用范围值移动行</span><span class="sxs-lookup"><span data-stu-id="89273-111">Sample code: Move rows using range values</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const IndexOfColumnToFilterOn = 1;
  const NameOfColumnToFilterOn = 'Category';
  const ValueToFilterOn = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }

  // Save the filter criteria.
  const tableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      tableFilters[column.getName()] = colFilterCriteria;
    }
  });

  // Get all the data from the table.
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  // Create variables to hold the rows to be moved and their addresses.
  let rowsToMoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values from the source table.
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][IndexOfColumnToFilterOn] === ValueToFilterOn) {
      rowsToMoveValues.push(dataRows[i]);

      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove.
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }

  // If there are no data rows to process, end the script.
  if (rowsToMoveValues.length < 1) {
    console.log('No rows selected from the source table match the filter criteria.');
    return;
  }

  console.log(`Adding ${rowsToMoveValues.length} rows to target table.`);

  // Insert rows at the end of target table.
  targetTable.addRows(-1, rowsToMoveValues)

  // Remove the rows from the source table.
  const sheet = sourceTable.getWorksheet();

  // Remove all filters before removing rows.
  sourceTable.getAutoFilter().clearCriteria();

  // Important: Remove the rows starting at the bottom of the table.
  // Otherwise, the lower rows change position before they are deleted.
  console.log(`Removing ${rowAddressToRemove.length} rows from the source table.`);
  rowAddressToRemove.reverse().forEach((address) => {
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });

  // Reapply the original filters. 
  Object.keys(tableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a><span data-ttu-id="89273-112">培训视频：跨表移动行</span><span class="sxs-lookup"><span data-stu-id="89273-112">Training video: Move rows across tables</span></span>

<span data-ttu-id="89273-113">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/_3t3Pk4i2L0)。</span><span class="sxs-lookup"><span data-stu-id="89273-113">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/_3t3Pk4i2L0).</span></span> <span data-ttu-id="89273-114">视频解决方案中显示了两个脚本。</span><span class="sxs-lookup"><span data-stu-id="89273-114">There are two scripts shown in the video's solution.</span></span> <span data-ttu-id="89273-115">主要区别是如何选择行。</span><span class="sxs-lookup"><span data-stu-id="89273-115">The main difference is how the rows are selected.</span></span>

* <span data-ttu-id="89273-116">第一个变量中，通过应用表筛选器并读取可见区域来选择行。</span><span class="sxs-lookup"><span data-stu-id="89273-116">In the first variant, the rows are selected by applying the table filter and reading the visible range.</span></span>
* <span data-ttu-id="89273-117">第二步，通过读取值并提取行值来选择行 (此页上的示例使用行) 。</span><span class="sxs-lookup"><span data-stu-id="89273-117">In the second, the rows are selected by reading the values and extracting the row values (which is what the sample on this page uses).</span></span>
