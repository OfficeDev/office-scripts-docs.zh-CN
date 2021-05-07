---
title: 使用脚本跨表Office行
description: 了解如何通过保存筛选器，然后处理和重新应用筛选器来跨表移动行。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 959fb002b0ba485b43f4de7de3004e1074f768a7
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232751"
---
# <a name="move-rows-across-tables-by-saving-filters-then-processing-and-reapplying-the-filters"></a><span data-ttu-id="fed82-103">通过保存筛选器，然后处理和重新应用筛选器，跨表移动行</span><span class="sxs-lookup"><span data-stu-id="fed82-103">Move rows across tables by saving filters, then processing and reapplying the filters</span></span>

<span data-ttu-id="fed82-104">此脚本执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="fed82-104">This script does the following:</span></span>

* <span data-ttu-id="fed82-105">从源表中选择行，其中列中的值等于 _某个值_。</span><span class="sxs-lookup"><span data-stu-id="fed82-105">Selects rows from the source table where the value in a column is equal to _some value_.</span></span>
* <span data-ttu-id="fed82-106">将所有选定的行移动到另一 (工作表) 中的目标行。</span><span class="sxs-lookup"><span data-stu-id="fed82-106">Moves all selected rows into another (target) table on another worksheet.</span></span>
* <span data-ttu-id="fed82-107">重新应用源表上的相关筛选器。</span><span class="sxs-lookup"><span data-stu-id="fed82-107">Reapplies the relevant filters on the source table.</span></span>

:::image type="content" source="../../images/table-filter-before-after.png" alt-text="工作簿之前和之后屏幕截图":::

<span data-ttu-id="fed82-109">此解决方案有两个脚本。</span><span class="sxs-lookup"><span data-stu-id="fed82-109">There are two scripts in this solution.</span></span> <span data-ttu-id="fed82-110">主要区别是如何选择行。</span><span class="sxs-lookup"><span data-stu-id="fed82-110">The main difference is how the rows are selected.</span></span>

* <span data-ttu-id="fed82-111">第 [一个变量](#sample-code-move-rows-using-table-filter)是，通过应用表格筛选器并读取可见区域来选择行。</span><span class="sxs-lookup"><span data-stu-id="fed82-111">In the [first variant](#sample-code-move-rows-using-table-filter), the rows are selected by applying the table filter and reading the visible range.</span></span>
* <span data-ttu-id="fed82-112">第 [二](#sample-code-move-rows-using-range-values)种是，通过读取值并提取行值来选择行。</span><span class="sxs-lookup"><span data-stu-id="fed82-112">In the [second](#sample-code-move-rows-using-range-values), the rows are selected by reading the values and extracting the row values.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="fed82-113">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="fed82-113">Sample Excel file</span></span>

<span data-ttu-id="fed82-114">下载此 <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> 中使用的文件，以尝试一下！</span><span class="sxs-lookup"><span data-stu-id="fed82-114">Download the file <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> used in this solution to try it out yourself!</span></span>

## <a name="sample-code-move-rows-using-table-filter"></a><span data-ttu-id="fed82-115">示例代码：使用表筛选器移动行</span><span class="sxs-lookup"><span data-stu-id="fed82-115">Sample code: Move rows using table filter</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Update the table names, column name, and look up value as needed.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const NameOfColumnToFilterOn = 'Category';
  const ValueToFilterOn = 'Clothing';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);
  // If you don't know the table names, you can instead use the following code to fetch the first table on a given worksheet name.
  // let targetTable = workbook.getWorksheet('Sheet1').getTables()[0];
  // let sourceTable = workbook.getWorksheet('Sheet2').getTables()[0];

  
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Save all of the filter criteria.
  // Initialize with an empty object to hold the filter criteria.
  const tableFilters = {};
  // For each table column, collect the filter criteria.
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      // If we don't remove these two keys, the API fails for some reason. So, remove these.
      delete colFilterCriteria['@odata.type'];
      delete colFilterCriteria['subField'];

      tableFilters[column.getName()] = colFilterCriteria;
    }
  });

  // Clear all filters on the table so that new filter can be applied.
  sourceTable.getAutoFilter().clearCriteria();

  // Get the name of the column that needs to be filtered (to select rows).
  const columnToFilter = sourceTable.getColumnByName(NameOfColumnToFilterOn);

  // Apply checked items filter on table 'targetTable' column '1'.
  columnToFilter.getFilter().applyValuesFilter([ValueToFilterOn]);

  // Get the visible range reference.
  const rangeView = sourceTable.getRange().getVisibleView();

  // Get number of rows that matched the filter criteria.
  const rowsCount = rangeView.getRowCount();
  console.log("Rows selected including header: " + rowsCount);

  // If no data rows to process, exit script.
  if (rowsCount < 2) {
    console.log('No data rows selected from the source table that matched the filter criteria.');
    reApplyFilters(sourceTable, NameOfColumnToFilterOn, tableFilters);
    return;
  }

  // Get the data values to insert to target table.
  let dataRows: (number | string | boolean)[][] = [];
  rangeView.getRows().forEach( (r, i) => {
    dataRows.push(r.getValues()[0]);
    return;
  });
  // Remove header row as it shows up as part of visible range rows.
  dataRows.splice(0,1);
  console.log(`Adding ${dataRows.length} rows to target table.`);
  // Insert rows at the end of target table. Change the first argument to suit your target location (e.g., 0 for beginning, -1 for end).
  targetTable.addRows(-1, dataRows);

  // Collect all the range addresses to be removed.
  const rowAddressToRemove = rangeView.getRows().map((r) => r.getRange().getAddress());
  // Remove the header row address.
  rowAddressToRemove.splice(0,1);
  
  // Remove the filters from the table prior to deleting rows.
  columnToFilter.getFilter().clear();

  // Get worksheet reference where the table rows to be deleted resides.
  const sheet = sourceTable.getWorksheet();

  // Important: Reverse the address and remove from the bottom so that the right rows are removed.
  // If not reversed, the resulting row upwards shift will mean that incorrect rows will be removed. 
  console.log(`Removing ${rowAddressToRemove.length} from the source table. `)
  rowAddressToRemove.reverse().forEach((address) => {
    // console.log(`Deleting row: ${address}`);
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });

  // Format source table to wrap text (important to do this before reapplying filters).
  sourceTable.getRange().getFormat().setWrapText(true);  
  
  // Reapply filters.
  // Log the criteria for testing purpose (not required).
  console.log(tableFilters);
  reApplyFilters(sourceTable, NameOfColumnToFilterOn, tableFilters);
  console.log("Finished.")
  return;
}

function reApplyFilters(sourceTable: ExcelScript.Table, columnNameFilteredOn: string, tableFilters: {}): void {

  // Reapply all column filters.
  Object.keys(tableFilters).forEach((columnName) => {
    sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
  });
  // Clear the filter on the column that we used to select rows
  // as those rows may have already been removed (retain all other filters).
  sourceTable.getColumnByName(columnNameFilteredOn).getFilter().clear();
  return;
}
```

## <a name="sample-code-move-rows-using-range-values"></a><span data-ttu-id="fed82-116">示例代码：使用范围值移动行</span><span class="sxs-lookup"><span data-stu-id="fed82-116">Sample code: Move rows using range values</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Update the table names, column index to look up as needed.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const IndexOfColumnToFilterOn = 1; // 0-index
  const NameOfColumnToFilterOn = 'Category';
  const ValueToFilterOn = 'Clothing';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);
  
  // If you don't know the table names, you can fetch first table on a given worksheet name.
  // let targetTable = workbook.getWorksheet('Sheet1').getTables()[0];
  // let sourceTable = workbook.getWorksheet('Sheet2').getTables()[0];

  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }

  // Save all of the filter criteria.
  // Initialize with an empty object to hold the filter criteria.
  const tableFilters = {};
  // For each table column, collect the filter criteria.
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      // If we don't remove these two keys, the API fails for some reason. So, remove these.
      delete colFilterCriteria['@odata.type'];
      delete colFilterCriteria['subField'];
      tableFilters[column.getName()] = colFilterCriteria;
    }
  });

  // Range object of table data.
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
  // Get data values of the table rows.
  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  // Create variables to hold the rows to be moved and their addresses.
  let rowsToMoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values to insert to target table.
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][IndexOfColumnToFilterOn] === ValueToFilterOn) {
      rowsToMoveValues.push(dataRows[i]);
      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove.
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }

  // If no data rows to process, exit script.
  if (rowsToMoveValues.length < 1) {
    console.log('No rows selected from the source table that matched the filter criteria.');
    return;
  }
  console.log(`Adding ${rowsToMoveValues.length} rows to target table.`);
  // Insert rows at the end of target table. Change the first argument to suit your target location (e.g., 0 for beginning, -1 for end).
  targetTable.addRows(-1, rowsToMoveValues)
  // Get worksheet reference where the table rows to be deleted resides. 
  const sheet = sourceTable.getWorksheet();

  // Remove all filters before removing rows.
  sourceTable.getAutoFilter().clearCriteria();

  // Important: Reverse the address and remove from the bottom so that the right rows are removed.
  // If not reversed, the resulting row upwards shift will mean that incorrect rows will be removed. 
  console.log(`Removing ${rowAddressToRemove.length} from the source table. `)
  rowAddressToRemove.reverse().forEach((address) => {
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });
  // Reapply filters. 
  // Log the criteria for testing purpose (not required).
  console.log(tableFilters);
  // Format source table to wrap text (important to do this before reapplying filters).
  sourceTable.getRange().getFormat().setWrapText(true);  
  
  // Reapply all column filters.
  reApplyFilters(sourceTable, NameOfColumnToFilterOn, tableFilters);
  console.log("Finished.")
  return;
}

function reApplyFilters(sourceTable: ExcelScript.Table, columnNameFilteredOn: string, tableFilters: {}): void {

  // Reapply all column filters.
  Object.keys(tableFilters).forEach((columnName) => {
    sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
  });
  sourceTable.getColumnByName(columnNameFilteredOn).getFilter().clear();
  return;
}
```

## <a name="training-video-move-rows-across-tables"></a><span data-ttu-id="fed82-117">培训视频：跨表移动行</span><span class="sxs-lookup"><span data-stu-id="fed82-117">Training video: Move rows across tables</span></span>

<span data-ttu-id="fed82-118">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/_3t3Pk4i2L0)。</span><span class="sxs-lookup"><span data-stu-id="fed82-118">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/_3t3Pk4i2L0).</span></span>
