---
title: Web 上的 Excel 中 Office 脚本的示例脚本
description: 要用于 web 上 Excel 中的 Office 脚本的一组代码示例。
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: fa330bfa284799e26ee2cf49800102072d66612b
ms.sourcegitcommit: 8d549884e68170f808d3d417104a4451a37da83c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2020
ms.locfileid: "45229601"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a>Excel 网页版中 Office 脚本的示例脚本（预览）

下面的示例是您在自己的工作簿中尝试的简单脚本。 若要在 Excel 网页上使用它们，请执行以下操作：

1. 打开“**自动**”选项卡。
2. 按**代码编辑器**。
3. 在代码编辑器的任务窗格中，按 "**新建脚本**"。
4. 将整个脚本替换为您选择的示例。
5. 在代码编辑器的任务窗格中按 "**运行**"。

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a>脚本基础

这些示例演示 Office 脚本的基本构建基块。 将这些应用程序添加到脚本以扩展解决方案并解决常见问题。

### <a name="read-and-log-one-cell"></a>读取和记录一个单元格

此示例读取**A1**的值并将其打印到控制台。

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

### <a name="read-the-active-cell"></a>读取活动单元格

此脚本记录当前活动单元格的值。 如果选择了多个单元格，则将记录最左侧的单元格。

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>更改相邻单元格

此脚本使用相对引用获取相邻的单元格。 请注意，如果活动单元格位于最上面一行，则脚本的一部分将失败，因为它引用当前选定的单元格上面的单元格。

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

### <a name="change-all-adjacent-cells"></a>更改所有相邻单元格

此脚本将活动单元格中的格式复制到相邻单元格。 请注意，此脚本仅当活动单元格不在工作表的边缘时才有效。

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

### <a name="change-each-individual-cell-in-a-range"></a>更改区域中的每个单个单元格

此脚本循环访问当前选定的区域。 它清除当前的格式设置，并将每个单元格中的填充颜色设置为随机颜色。

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

## <a name="collections"></a>收藏

这些示例处理工作簿中的对象集合。

### <a name="iterating-over-collections"></a>循环访问集合

此脚本获取并记录工作簿中所有工作表的名称。 它还将其制表符颜色设置为随机颜色。

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

## <a name="dates"></a>日期

本节中的示例演示如何使用 JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date)对象。

下面的示例获取当前日期和时间，然后将这些值写入活动工作表中的两个单元格。

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

下一个示例读取存储在 Excel 中的日期，并将其转换为 JavaScript Date 对象。 它使用[日期的数字序列号](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)作为 JavaScript 日期的输入。

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

## <a name="display-data"></a>显示数据

这些示例演示如何使用工作表数据，并为用户提供更好的视图或组织。

### <a name="apply-conditional-formatting"></a>应用条件格式

此示例向工作表中当前使用的区域应用条件格式。 条件格式是前10% 的数值的绿色填充。

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

### <a name="create-a-sorted-table"></a>创建已排序的表

本示例从当前工作表的已用区域创建一个表格，然后基于第一列对其进行排序。

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a>记录数据透视表中的 "总计" 值

本示例在工作簿中查找第一个数据透视表，并将值记录在 "总计" 单元格中（在下图中突出显示为绿色）。

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

## <a name="scenario-samples"></a>方案示例

有关 showcasing 大型的真实解决方案的示例，请访问[Office 脚本的示例方案](scenarios/sample-scenario-overview.md)。

## <a name="suggest-new-samples"></a>建议新示例

我们欢迎您提出新示例建议。 如果有一个可帮助其他脚本开发人员的常见方案，请在下面的 "反馈" 部分告诉我们。
