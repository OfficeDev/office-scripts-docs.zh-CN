---
title: Office脚本的基本Excel web 版
description: 要与 Excel web 版 中的脚本Office代码示例Excel web 版。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: ea0430910aa16ef8a0eed04cf9ebcab7d611ae62
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52233005"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="5178d-103">Office脚本的基本Excel web 版</span><span class="sxs-lookup"><span data-stu-id="5178d-103">Basic scripts for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="5178d-104">以下示例是简单脚本，您可以尝试自己的工作簿。</span><span class="sxs-lookup"><span data-stu-id="5178d-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="5178d-105">若要在Excel web 版：</span><span class="sxs-lookup"><span data-stu-id="5178d-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="5178d-106">打开“**自动**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="5178d-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="5178d-107">按 **代码编辑器**。</span><span class="sxs-lookup"><span data-stu-id="5178d-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="5178d-108">在 **代码编辑器** 的任务窗格中按"新建脚本"。</span><span class="sxs-lookup"><span data-stu-id="5178d-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="5178d-109">将整个脚本替换为你选择的示例。</span><span class="sxs-lookup"><span data-stu-id="5178d-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="5178d-110">在 **代码** 编辑器的任务窗格中按"运行"。</span><span class="sxs-lookup"><span data-stu-id="5178d-110">Press **Run** in the Code Editor's task pane.</span></span>

## <a name="scripting-basics"></a><span data-ttu-id="5178d-111">脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="5178d-111">Scripting basics</span></span>

<span data-ttu-id="5178d-112">这些示例演示了脚本的基本Office构建基块。</span><span class="sxs-lookup"><span data-stu-id="5178d-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="5178d-113">将其添加到脚本以扩展解决方案并解决常见问题。</span><span class="sxs-lookup"><span data-stu-id="5178d-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="5178d-114">读取和记录一个单元格</span><span class="sxs-lookup"><span data-stu-id="5178d-114">Read and log one cell</span></span>

<span data-ttu-id="5178d-115">此示例读取 **A1 的值，** 并打印到控制台。</span><span class="sxs-lookup"><span data-stu-id="5178d-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="5178d-116">读取活动单元格</span><span class="sxs-lookup"><span data-stu-id="5178d-116">Read the active cell</span></span>

<span data-ttu-id="5178d-117">此脚本记录当前活动单元格的值。</span><span class="sxs-lookup"><span data-stu-id="5178d-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="5178d-118">如果选择了多个单元格，将记录最左上方的单元格。</span><span class="sxs-lookup"><span data-stu-id="5178d-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="5178d-119">更改相邻单元格</span><span class="sxs-lookup"><span data-stu-id="5178d-119">Change an adjacent cell</span></span>

<span data-ttu-id="5178d-120">此脚本使用相对引用获取相邻单元格。</span><span class="sxs-lookup"><span data-stu-id="5178d-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="5178d-121">请注意，如果活动单元格位于最上面一行，脚本的一部分将失败，因为它引用当前选定单元格上方的单元格。</span><span class="sxs-lookup"><span data-stu-id="5178d-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="5178d-122">更改所有相邻单元格</span><span class="sxs-lookup"><span data-stu-id="5178d-122">Change all adjacent cells</span></span>

<span data-ttu-id="5178d-123">此脚本将活动单元格中的格式复制到相邻单元格。</span><span class="sxs-lookup"><span data-stu-id="5178d-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="5178d-124">请注意，此脚本仅在活动单元格不在工作表边缘时有效。</span><span class="sxs-lookup"><span data-stu-id="5178d-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="5178d-125">更改区域的每个单元格</span><span class="sxs-lookup"><span data-stu-id="5178d-125">Change each individual cell in a range</span></span>

<span data-ttu-id="5178d-126">此脚本将循环遍历当前选择的范围。</span><span class="sxs-lookup"><span data-stu-id="5178d-126">This script loops over the currently select range.</span></span> <span data-ttu-id="5178d-127">它清除当前格式，将每个单元格中的填充颜色设置为随机颜色。</span><span class="sxs-lookup"><span data-stu-id="5178d-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="5178d-128">根据特殊条件获取单元格组</span><span class="sxs-lookup"><span data-stu-id="5178d-128">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="5178d-129">此脚本获取当前工作表的已用区域的所有空白单元格。</span><span class="sxs-lookup"><span data-stu-id="5178d-129">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="5178d-130">然后，它用黄色背景突出显示所有这些单元格。</span><span class="sxs-lookup"><span data-stu-id="5178d-130">It then highlights all those cells with a yellow background.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a><span data-ttu-id="5178d-131">收藏</span><span class="sxs-lookup"><span data-stu-id="5178d-131">Collections</span></span>

<span data-ttu-id="5178d-132">这些示例处理工作簿中的对象集合。</span><span class="sxs-lookup"><span data-stu-id="5178d-132">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterating-over-collections"></a><span data-ttu-id="5178d-133">对集合进行 Iterating</span><span class="sxs-lookup"><span data-stu-id="5178d-133">Iterating over collections</span></span>

<span data-ttu-id="5178d-134">此脚本获取并记录工作簿中所有工作表的名称。</span><span class="sxs-lookup"><span data-stu-id="5178d-134">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="5178d-135">它还将选项卡颜色设置为随机颜色。</span><span class="sxs-lookup"><span data-stu-id="5178d-135">It also sets the their tab colors to a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### <a name="querying-and-deleting-from-a-collection"></a><span data-ttu-id="5178d-136">查询和删除集合</span><span class="sxs-lookup"><span data-stu-id="5178d-136">Querying and deleting from a collection</span></span>

<span data-ttu-id="5178d-137">此脚本创建新的工作表。</span><span class="sxs-lookup"><span data-stu-id="5178d-137">This script creates a new worksheet.</span></span> <span data-ttu-id="5178d-138">它在新建工作表之前检查工作表的现有副本并将其删除。</span><span class="sxs-lookup"><span data-stu-id="5178d-138">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## <a name="dates"></a><span data-ttu-id="5178d-139">日期</span><span class="sxs-lookup"><span data-stu-id="5178d-139">Dates</span></span>

<span data-ttu-id="5178d-140">本节中的示例显示如何使用 JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) 对象。</span><span class="sxs-lookup"><span data-stu-id="5178d-140">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="5178d-141">以下示例获取当前日期和时间，然后将这些值写入活动工作表中的两个单元格。</span><span class="sxs-lookup"><span data-stu-id="5178d-141">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

<span data-ttu-id="5178d-142">下一个示例将读取存储在 Excel 并将其转换为 JavaScript Date 对象。</span><span class="sxs-lookup"><span data-stu-id="5178d-142">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="5178d-143">它将日期 [的数字序列号用作](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) JavaScript Date 的输入。</span><span class="sxs-lookup"><span data-stu-id="5178d-143">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="5178d-144">显示数据</span><span class="sxs-lookup"><span data-stu-id="5178d-144">Display data</span></span>

<span data-ttu-id="5178d-145">这些示例演示如何使用工作表数据，并为用户提供更好的视图或组织。</span><span class="sxs-lookup"><span data-stu-id="5178d-145">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="5178d-146">应用条件格式</span><span class="sxs-lookup"><span data-stu-id="5178d-146">Apply conditional formatting</span></span>

<span data-ttu-id="5178d-147">本示例将条件格式应用于工作表中当前使用的范围。</span><span class="sxs-lookup"><span data-stu-id="5178d-147">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="5178d-148">条件格式是前 10% 的值的绿色填充。</span><span class="sxs-lookup"><span data-stu-id="5178d-148">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="5178d-149">创建排序表</span><span class="sxs-lookup"><span data-stu-id="5178d-149">Create a sorted table</span></span>

<span data-ttu-id="5178d-150">本示例从当前工作表的已用区域创建一个表格，然后根据第一列对表格进行排序。</span><span class="sxs-lookup"><span data-stu-id="5178d-150">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="5178d-151">记录数据透视表中的"总计"值</span><span class="sxs-lookup"><span data-stu-id="5178d-151">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="5178d-152">本示例查找工作簿中的第一个数据透视表，并记录"总计"单元格 (在下面的图像中以绿色突出显示) 。</span><span class="sxs-lookup"><span data-stu-id="5178d-152">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="一个数据透视表，其中&quot;总计&quot;行突出显示为绿色，":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

### <a name="use-data-validation-to-create-a-drop-down-list"></a><span data-ttu-id="5178d-154">使用数据验证创建下拉列表</span><span class="sxs-lookup"><span data-stu-id="5178d-154">Use data validation to create a drop-down list</span></span>

<span data-ttu-id="5178d-155">此脚本为单元格创建下拉选择列表。</span><span class="sxs-lookup"><span data-stu-id="5178d-155">This script creates a drop-down selection list for a cell.</span></span> <span data-ttu-id="5178d-156">它将所选区域的现有值用作列表的选项。</span><span class="sxs-lookup"><span data-stu-id="5178d-156">It uses the existing values of the selected range as the choices for the list.</span></span>

:::image type="content" source="../../images/sample-data-validation.png" alt-text="显示包含颜色选项&quot;红色、蓝色、绿色&quot;且旁边包含颜色选项的三个单元格的工作表，下拉列表中显示的选项相同":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  let selectedRange = workbook.getSelectedRange();
  let rangeValues = selectedRange.getValues();

  // Convert the values into a comma-delimited string.
  let dataValidationListString = "";
  rangeValues.forEach((rangeValueRow) => {
    rangeValueRow.forEach((value) => {
      dataValidationListString += value + ",";
    });
  });

  // Clear the old range.
  selectedRange.clear(ExcelScript.ClearApplyTo.contents);

  // Apply the data validation to the first cell in the selected range.
  let targetCell = selectedRange.getCell(0,0);
  let dataValidation = targetCell.getDataValidation();

  // Set the content of the drop-down list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```

## <a name="formulas"></a><span data-ttu-id="5178d-158">公式</span><span class="sxs-lookup"><span data-stu-id="5178d-158">Formulas</span></span>

<span data-ttu-id="5178d-159">这些示例Excel公式，并展示如何在脚本中使用它们。</span><span class="sxs-lookup"><span data-stu-id="5178d-159">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="5178d-160">单个公式</span><span class="sxs-lookup"><span data-stu-id="5178d-160">Single formula</span></span>

<span data-ttu-id="5178d-161">此脚本设置单元格的公式，然后Excel单独存储单元格的公式和值。</span><span class="sxs-lookup"><span data-stu-id="5178d-161">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### <a name="spilling-results-from-a-formula"></a><span data-ttu-id="5178d-162">从公式中溢出结果</span><span class="sxs-lookup"><span data-stu-id="5178d-162">Spilling results from a formula</span></span>

<span data-ttu-id="5178d-163">此脚本使用 TRANSPOSE 函数将区域"A1：D2"转置为"A4：B7"。</span><span class="sxs-lookup"><span data-stu-id="5178d-163">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="5178d-164">如果转置导致错误#SPILL，它将清除目标区域并再次应用公式。</span><span class="sxs-lookup"><span data-stu-id="5178d-164">If the transpose results in a #SPILL error, it clears the target range and applies the formula again.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

## <a name="suggest-new-samples"></a><span data-ttu-id="5178d-165">建议新示例</span><span class="sxs-lookup"><span data-stu-id="5178d-165">Suggest new samples</span></span>

<span data-ttu-id="5178d-166">我们欢迎您提出有关新示例的建议。</span><span class="sxs-lookup"><span data-stu-id="5178d-166">We welcome suggestions for new samples.</span></span> <span data-ttu-id="5178d-167">如果有一种有助于其他脚本开发人员的常见方案，请在页面底部的反馈部分告诉我们。</span><span class="sxs-lookup"><span data-stu-id="5178d-167">If there is a common scenario that would help other script developers, please tell us in the feedback section at the bottom of the page.</span></span>

## <a name="see-also"></a><span data-ttu-id="5178d-168">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5178d-168">See also</span></span>

* [<span data-ttu-id="5178d-169">YouTube 上的 Sudhi Ramamurthy 的"Range 基础知识"</span><span class="sxs-lookup"><span data-stu-id="5178d-169">Sudhi Ramamurthy's "Range basics" on YouTube</span></span>](https://youtu.be/4emjkOFdLBA)
* [<span data-ttu-id="5178d-170">Office脚本示例和方案</span><span class="sxs-lookup"><span data-stu-id="5178d-170">Office Scripts samples and scenarios</span></span>](samples-overview.md)
* [<span data-ttu-id="5178d-171">在 Excel 网页版中录制、编辑和创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="5178d-171">Record, edit, and create Office Scripts in Excel on the web</span></span>](../../tutorials/excel-tutorial.md)
