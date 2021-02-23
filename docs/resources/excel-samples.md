---
title: Excel 网页中的 Office 脚本示例脚本
description: 要与 Excel 网页中的 Office 脚本一起使用的代码示例集合。
ms.date: 12/21/2020
localization_priority: Normal
ms.openlocfilehash: 35a7fdb4dcfa4c349aa594e5b13d1b7e4d33a178
ms.sourcegitcommit: 9df67e007ddbfec79a7360df9f4ea5ac6c86fb08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49772962"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="2f55f-103">Excel 网页版中的 Office 脚本示例 (预览) </span><span class="sxs-lookup"><span data-stu-id="2f55f-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="2f55f-104">以下示例是一些简单的脚本，您可以尝试自己的工作簿。</span><span class="sxs-lookup"><span data-stu-id="2f55f-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="2f55f-105">若要在 Excel 网页中使用它们，请：</span><span class="sxs-lookup"><span data-stu-id="2f55f-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="2f55f-106">打开“**自动**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="2f55f-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="2f55f-107">按 **代码编辑器**。</span><span class="sxs-lookup"><span data-stu-id="2f55f-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="2f55f-108">在 **代码编辑器** 的任务窗格中按"新建脚本"。</span><span class="sxs-lookup"><span data-stu-id="2f55f-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="2f55f-109">将整个脚本替换为你选择的示例。</span><span class="sxs-lookup"><span data-stu-id="2f55f-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="2f55f-110">在 **代码** 编辑器的任务窗格中按"运行"。</span><span class="sxs-lookup"><span data-stu-id="2f55f-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="2f55f-111">脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="2f55f-111">Scripting basics</span></span>

<span data-ttu-id="2f55f-112">这些示例演示 Office 脚本的基本构建基块。</span><span class="sxs-lookup"><span data-stu-id="2f55f-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="2f55f-113">将其添加到脚本以扩展解决方案并解决常见问题。</span><span class="sxs-lookup"><span data-stu-id="2f55f-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="2f55f-114">读取和记录一个单元格</span><span class="sxs-lookup"><span data-stu-id="2f55f-114">Read and log one cell</span></span>

<span data-ttu-id="2f55f-115">此示例读取 **A1 的值** ，并打印到控制台。</span><span class="sxs-lookup"><span data-stu-id="2f55f-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="2f55f-116">读取活动单元格</span><span class="sxs-lookup"><span data-stu-id="2f55f-116">Read the active cell</span></span>

<span data-ttu-id="2f55f-117">此脚本记录当前活动单元格的值。</span><span class="sxs-lookup"><span data-stu-id="2f55f-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="2f55f-118">如果选择了多个单元格，将记录最左上方的单元格。</span><span class="sxs-lookup"><span data-stu-id="2f55f-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="2f55f-119">更改相邻单元格</span><span class="sxs-lookup"><span data-stu-id="2f55f-119">Change an adjacent cell</span></span>

<span data-ttu-id="2f55f-120">此脚本使用相对引用获取相邻单元格。</span><span class="sxs-lookup"><span data-stu-id="2f55f-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="2f55f-121">请注意，如果活动单元格位于最上面一行，脚本的一部分将失败，因为它引用当前选定单元格上方的单元格。</span><span class="sxs-lookup"><span data-stu-id="2f55f-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```typescript
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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="2f55f-122">更改所有相邻单元格</span><span class="sxs-lookup"><span data-stu-id="2f55f-122">Change all adjacent cells</span></span>

<span data-ttu-id="2f55f-123">此脚本将活动单元格中的格式复制到相邻单元格。</span><span class="sxs-lookup"><span data-stu-id="2f55f-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="2f55f-124">请注意，此脚本仅在活动单元格不在工作表边缘时有效。</span><span class="sxs-lookup"><span data-stu-id="2f55f-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```typescript
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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="2f55f-125">更改区域的每个单元格</span><span class="sxs-lookup"><span data-stu-id="2f55f-125">Change each individual cell in a range</span></span>

<span data-ttu-id="2f55f-126">此脚本将循环遍历当前选择的范围。</span><span class="sxs-lookup"><span data-stu-id="2f55f-126">This script loops over the currently select range.</span></span> <span data-ttu-id="2f55f-127">它清除当前格式，将每个单元格中的填充颜色设置为随机颜色。</span><span class="sxs-lookup"><span data-stu-id="2f55f-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```typescript
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

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="2f55f-128">根据特殊条件获取单元格组</span><span class="sxs-lookup"><span data-stu-id="2f55f-128">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="2f55f-129">此脚本获取当前工作表使用区域的所有空白单元格。</span><span class="sxs-lookup"><span data-stu-id="2f55f-129">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="2f55f-130">然后，它突出显示所有带黄色背景的单元格。</span><span class="sxs-lookup"><span data-stu-id="2f55f-130">It then highlights all those cells with a yellow background.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a><span data-ttu-id="2f55f-131">收藏</span><span class="sxs-lookup"><span data-stu-id="2f55f-131">Collections</span></span>

<span data-ttu-id="2f55f-132">这些示例使用工作簿中的对象集合。</span><span class="sxs-lookup"><span data-stu-id="2f55f-132">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterating-over-collections"></a><span data-ttu-id="2f55f-133">对集合进行 Itererating</span><span class="sxs-lookup"><span data-stu-id="2f55f-133">Iterating over collections</span></span>

<span data-ttu-id="2f55f-134">此脚本获取并记录工作簿中所有工作表的名称。</span><span class="sxs-lookup"><span data-stu-id="2f55f-134">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="2f55f-135">它还将选项卡颜色设置为随机颜色。</span><span class="sxs-lookup"><span data-stu-id="2f55f-135">It also sets the their tab colors to a random color.</span></span>

```typescript
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

### <a name="querying-and-deleting-from-a-collection"></a><span data-ttu-id="2f55f-136">从集合中查询和删除</span><span class="sxs-lookup"><span data-stu-id="2f55f-136">Querying and deleting from a collection</span></span>

<span data-ttu-id="2f55f-137">此脚本创建新的工作表。</span><span class="sxs-lookup"><span data-stu-id="2f55f-137">This script creates a new worksheet.</span></span> <span data-ttu-id="2f55f-138">它在新建工作表之前检查工作表的现有副本并将其删除。</span><span class="sxs-lookup"><span data-stu-id="2f55f-138">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```typescript
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

## <a name="dates"></a><span data-ttu-id="2f55f-139">日期</span><span class="sxs-lookup"><span data-stu-id="2f55f-139">Dates</span></span>

<span data-ttu-id="2f55f-140">本节中的示例显示如何使用 JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) 对象。</span><span class="sxs-lookup"><span data-stu-id="2f55f-140">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="2f55f-141">以下示例获取当前日期和时间，然后将这些值写入活动工作表中的两个单元格。</span><span class="sxs-lookup"><span data-stu-id="2f55f-141">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="2f55f-142">下一个示例将读取 Excel 中存储的日期，并将其转换为 JavaScript Date 对象。</span><span class="sxs-lookup"><span data-stu-id="2f55f-142">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="2f55f-143">它将 [日期的数字序列号用作](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) JavaScript 日期的输入。</span><span class="sxs-lookup"><span data-stu-id="2f55f-143">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="2f55f-144">显示数据</span><span class="sxs-lookup"><span data-stu-id="2f55f-144">Display data</span></span>

<span data-ttu-id="2f55f-145">这些示例演示如何使用工作表数据，并为用户提供更好的视图或组织。</span><span class="sxs-lookup"><span data-stu-id="2f55f-145">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="2f55f-146">应用条件格式</span><span class="sxs-lookup"><span data-stu-id="2f55f-146">Apply conditional formatting</span></span>

<span data-ttu-id="2f55f-147">本示例将条件格式应用于工作表中当前使用的范围。</span><span class="sxs-lookup"><span data-stu-id="2f55f-147">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="2f55f-148">条件格式是前 10% 值的绿色填充。</span><span class="sxs-lookup"><span data-stu-id="2f55f-148">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="2f55f-149">创建排序表</span><span class="sxs-lookup"><span data-stu-id="2f55f-149">Create a sorted table</span></span>

<span data-ttu-id="2f55f-150">本示例从当前工作表的已用区域创建一个表格，然后基于第一列对表格进行排序。</span><span class="sxs-lookup"><span data-stu-id="2f55f-150">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="2f55f-151">记录数据透视表中的"总计"值</span><span class="sxs-lookup"><span data-stu-id="2f55f-151">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="2f55f-152">本示例查找工作簿中的第一个数据透视表，并记录"总计"单元格 (在下面的图像中以绿色突出显示) 。</span><span class="sxs-lookup"><span data-stu-id="2f55f-152">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![一个结果销售数据透视表，"总计"行突出显示为绿色。](../images/sample-pivottable-grand-total-row.png)

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

## <a name="formulas"></a><span data-ttu-id="2f55f-154">公式</span><span class="sxs-lookup"><span data-stu-id="2f55f-154">Formulas</span></span>

<span data-ttu-id="2f55f-155">这些示例使用 Excel 公式，并展示如何在脚本中使用它们。</span><span class="sxs-lookup"><span data-stu-id="2f55f-155">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="2f55f-156">单个公式</span><span class="sxs-lookup"><span data-stu-id="2f55f-156">Single formula</span></span>

<span data-ttu-id="2f55f-157">此脚本设置单元格的公式，然后显示 Excel 如何单独存储单元格的公式和值。</span><span class="sxs-lookup"><span data-stu-id="2f55f-157">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```typescript
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

### <a name="spilling-results-from-a-formula"></a><span data-ttu-id="2f55f-158">从公式中溢出结果</span><span class="sxs-lookup"><span data-stu-id="2f55f-158">Spilling results from a formula</span></span>

<span data-ttu-id="2f55f-159">此脚本使用 TRANSPOSE 函数将区域"A1：D2"转置为"A4：B7"。</span><span class="sxs-lookup"><span data-stu-id="2f55f-159">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="2f55f-160">如果转置导致#SPILL错误，它将清除目标区域并再次应用公式。</span><span class="sxs-lookup"><span data-stu-id="2f55f-160">If the transpose results in a #SPILL error, it clears the target range and applies the formula again.</span></span>

```typescript
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

## <a name="scenario-samples"></a><span data-ttu-id="2f55f-161">方案示例</span><span class="sxs-lookup"><span data-stu-id="2f55f-161">Scenario samples</span></span>

<span data-ttu-id="2f55f-162">有关展示大型实际解决方案的示例，请访问 [Office 脚本的示例方案](scenarios/sample-scenario-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="2f55f-162">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="2f55f-163">建议新示例</span><span class="sxs-lookup"><span data-stu-id="2f55f-163">Suggest new samples</span></span>

<span data-ttu-id="2f55f-164">欢迎提供新示例建议。</span><span class="sxs-lookup"><span data-stu-id="2f55f-164">We welcome suggestions for new samples.</span></span> <span data-ttu-id="2f55f-165">如果存在有助于其他脚本开发人员的常见方案，请在下面的反馈部分中告诉我们。</span><span class="sxs-lookup"><span data-stu-id="2f55f-165">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
