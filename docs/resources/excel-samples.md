---
title: Web 上的 Excel 中 Office 脚本的示例脚本
description: 要用于 web 上 Excel 中的 Office 脚本的一组代码示例。
ms.date: 08/04/2020
localization_priority: Normal
ms.openlocfilehash: 4f8d6f2395a841a8dcba2ea0e712e645a84a6d91
ms.sourcegitcommit: 1c88abcf5df16a05913f12df89490ce843cfebe2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/13/2020
ms.locfileid: "46665227"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="3498a-103">Excel 网页版中的 Office 脚本示例脚本 (预览) </span><span class="sxs-lookup"><span data-stu-id="3498a-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="3498a-104">下面的示例是您在自己的工作簿中尝试的简单脚本。</span><span class="sxs-lookup"><span data-stu-id="3498a-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="3498a-105">若要在 Excel 网页上使用它们，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="3498a-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="3498a-106">打开“**自动**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="3498a-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="3498a-107">按 **代码编辑器**。</span><span class="sxs-lookup"><span data-stu-id="3498a-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="3498a-108">在代码编辑器的任务窗格中，按 " **新建脚本** "。</span><span class="sxs-lookup"><span data-stu-id="3498a-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="3498a-109">将整个脚本替换为您选择的示例。</span><span class="sxs-lookup"><span data-stu-id="3498a-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="3498a-110">在代码编辑器的任务窗格中按 " **运行** "。</span><span class="sxs-lookup"><span data-stu-id="3498a-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="3498a-111">脚本基础</span><span class="sxs-lookup"><span data-stu-id="3498a-111">Scripting basics</span></span>

<span data-ttu-id="3498a-112">这些示例演示 Office 脚本的基本构建基块。</span><span class="sxs-lookup"><span data-stu-id="3498a-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="3498a-113">将这些应用程序添加到脚本以扩展解决方案并解决常见问题。</span><span class="sxs-lookup"><span data-stu-id="3498a-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="3498a-114">读取和记录一个单元格</span><span class="sxs-lookup"><span data-stu-id="3498a-114">Read and log one cell</span></span>

<span data-ttu-id="3498a-115">此示例读取 **A1** 的值并将其打印到控制台。</span><span class="sxs-lookup"><span data-stu-id="3498a-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="read-the-active-cell"></a><span data-ttu-id="3498a-116">读取活动单元格</span><span class="sxs-lookup"><span data-stu-id="3498a-116">Read the active cell</span></span>

<span data-ttu-id="3498a-117">此脚本记录当前活动单元格的值。</span><span class="sxs-lookup"><span data-stu-id="3498a-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="3498a-118">如果选择了多个单元格，则将记录最左侧的单元格。</span><span class="sxs-lookup"><span data-stu-id="3498a-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="3498a-119">更改相邻单元格</span><span class="sxs-lookup"><span data-stu-id="3498a-119">Change an adjacent cell</span></span>

<span data-ttu-id="3498a-120">此脚本使用相对引用获取相邻的单元格。</span><span class="sxs-lookup"><span data-stu-id="3498a-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="3498a-121">请注意，如果活动单元格位于最上面一行，则脚本的一部分将失败，因为它引用当前选定的单元格上面的单元格。</span><span class="sxs-lookup"><span data-stu-id="3498a-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="3498a-122">更改所有相邻单元格</span><span class="sxs-lookup"><span data-stu-id="3498a-122">Change all adjacent cells</span></span>

<span data-ttu-id="3498a-123">此脚本将活动单元格中的格式复制到相邻单元格。</span><span class="sxs-lookup"><span data-stu-id="3498a-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="3498a-124">请注意，此脚本仅当活动单元格不在工作表的边缘时才有效。</span><span class="sxs-lookup"><span data-stu-id="3498a-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="3498a-125">更改区域中的每个单个单元格</span><span class="sxs-lookup"><span data-stu-id="3498a-125">Change each individual cell in a range</span></span>

<span data-ttu-id="3498a-126">此脚本循环访问当前选定的区域。</span><span class="sxs-lookup"><span data-stu-id="3498a-126">This script loops over the currently select range.</span></span> <span data-ttu-id="3498a-127">它清除当前的格式设置，并将每个单元格中的填充颜色设置为随机颜色。</span><span class="sxs-lookup"><span data-stu-id="3498a-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

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

## <a name="collections"></a><span data-ttu-id="3498a-128">收藏</span><span class="sxs-lookup"><span data-stu-id="3498a-128">Collections</span></span>

<span data-ttu-id="3498a-129">这些示例处理工作簿中的对象集合。</span><span class="sxs-lookup"><span data-stu-id="3498a-129">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterating-over-collections"></a><span data-ttu-id="3498a-130">循环访问集合</span><span class="sxs-lookup"><span data-stu-id="3498a-130">Iterating over collections</span></span>

<span data-ttu-id="3498a-131">此脚本获取并记录工作簿中所有工作表的名称。</span><span class="sxs-lookup"><span data-stu-id="3498a-131">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="3498a-132">它还将其制表符颜色设置为随机颜色。</span><span class="sxs-lookup"><span data-stu-id="3498a-132">It also sets the their tab colors to a random color.</span></span>

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

### <a name="querying-and-deleting-from-a-collection"></a><span data-ttu-id="3498a-133">查询和删除集合</span><span class="sxs-lookup"><span data-stu-id="3498a-133">Querying and deleting from a collection</span></span>

<span data-ttu-id="3498a-134">此脚本将创建一个新的工作表。</span><span class="sxs-lookup"><span data-stu-id="3498a-134">This script creates a new worksheet.</span></span> <span data-ttu-id="3498a-135">它将检查工作表的现有副本并在生成新工作表之前将其删除。</span><span class="sxs-lookup"><span data-stu-id="3498a-135">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

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

## <a name="dates"></a><span data-ttu-id="3498a-136">日期</span><span class="sxs-lookup"><span data-stu-id="3498a-136">Dates</span></span>

<span data-ttu-id="3498a-137">本节中的示例演示如何使用 JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) 对象。</span><span class="sxs-lookup"><span data-stu-id="3498a-137">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="3498a-138">下面的示例获取当前日期和时间，然后将这些值写入活动工作表中的两个单元格。</span><span class="sxs-lookup"><span data-stu-id="3498a-138">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="3498a-139">下一个示例读取存储在 Excel 中的日期，并将其转换为 JavaScript Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3498a-139">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="3498a-140">它使用 [日期的数字序列号](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) 作为 JavaScript 日期的输入。</span><span class="sxs-lookup"><span data-stu-id="3498a-140">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue();
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="3498a-141">显示数据</span><span class="sxs-lookup"><span data-stu-id="3498a-141">Display data</span></span>

<span data-ttu-id="3498a-142">这些示例演示如何使用工作表数据，并为用户提供更好的视图或组织。</span><span class="sxs-lookup"><span data-stu-id="3498a-142">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="3498a-143">应用条件格式</span><span class="sxs-lookup"><span data-stu-id="3498a-143">Apply conditional formatting</span></span>

<span data-ttu-id="3498a-144">此示例向工作表中当前使用的区域应用条件格式。</span><span class="sxs-lookup"><span data-stu-id="3498a-144">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="3498a-145">条件格式是前10% 的数值的绿色填充。</span><span class="sxs-lookup"><span data-stu-id="3498a-145">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="3498a-146">创建已排序的表</span><span class="sxs-lookup"><span data-stu-id="3498a-146">Create a sorted table</span></span>

<span data-ttu-id="3498a-147">本示例从当前工作表的已用区域创建一个表格，然后基于第一列对其进行排序。</span><span class="sxs-lookup"><span data-stu-id="3498a-147">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="3498a-148">记录数据透视表中的 "总计" 值</span><span class="sxs-lookup"><span data-stu-id="3498a-148">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="3498a-149">本示例在工作簿中查找第一个数据透视表，并将 "总计" 单元格中的值记录 () 下图中的绿色突出显示）。</span><span class="sxs-lookup"><span data-stu-id="3498a-149">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![一个水果销售数据透视表，总计行突出显示为绿色。](../images/sample-pivottable-grand-total-row.png)

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getRangeBetweenHeaderAndTotal();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

## <a name="formulas"></a><span data-ttu-id="3498a-151">公式</span><span class="sxs-lookup"><span data-stu-id="3498a-151">Formulas</span></span>

<span data-ttu-id="3498a-152">这些示例使用 Excel 公式，并演示如何在脚本中使用它们。</span><span class="sxs-lookup"><span data-stu-id="3498a-152">These samples use Excel formulas and show how to work with them in scripts.</span></span>

## <a name="single-formula"></a><span data-ttu-id="3498a-153">单个公式</span><span class="sxs-lookup"><span data-stu-id="3498a-153">Single formula</span></span>

<span data-ttu-id="3498a-154">此脚本设置单元格的公式，然后显示 Excel 如何单独存储单元格的公式和值。</span><span class="sxs-lookup"><span data-stu-id="3498a-154">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

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

### <a name="spilling-results-from-a-formula"></a><span data-ttu-id="3498a-155">Spilling 公式中的结果</span><span class="sxs-lookup"><span data-stu-id="3498a-155">Spilling results from a formula</span></span>

<span data-ttu-id="3498a-156">此脚本使用换位函数将区域 "A1： D2" 转置为 "A4： B7"。</span><span class="sxs-lookup"><span data-stu-id="3498a-156">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="3498a-157">如果换位的转置结果为 #SPILL 错误，它将清除目标区域并再次应用公式。</span><span class="sxs-lookup"><span data-stu-id="3498a-157">If the transpose results in a #SPILL error, it clears the target range and applies the formula again.</span></span>

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

## <a name="scenario-samples"></a><span data-ttu-id="3498a-158">方案示例</span><span class="sxs-lookup"><span data-stu-id="3498a-158">Scenario samples</span></span>

<span data-ttu-id="3498a-159">有关 showcasing 大型的真实解决方案的示例，请访问 [Office 脚本的示例方案](scenarios/sample-scenario-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="3498a-159">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="3498a-160">建议新示例</span><span class="sxs-lookup"><span data-stu-id="3498a-160">Suggest new samples</span></span>

<span data-ttu-id="3498a-161">我们欢迎您提出新示例建议。</span><span class="sxs-lookup"><span data-stu-id="3498a-161">We welcome suggestions for new samples.</span></span> <span data-ttu-id="3498a-162">如果有一个可帮助其他脚本开发人员的常见方案，请在下面的 "反馈" 部分告诉我们。</span><span class="sxs-lookup"><span data-stu-id="3498a-162">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
