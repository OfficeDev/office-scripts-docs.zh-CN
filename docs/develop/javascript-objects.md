---
title: 在 Office 脚本中使用内置的 JavaScript 对象
description: 如何从 Excel web 版 中的 Office 脚本调用内置 JavaScript EXCEL WEB 版。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 6c15daf0429009d289a17e604caf51b807510442bf6e6fa6e42c85d7457f6164
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846608"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>在脚本中使用内置的 JavaScript Office对象

JavaScript 提供了多个可用于 Office 脚本的内置对象，无论你是使用 JavaScript 还是[TypeScript](../overview/code-editor-environment.md)编写脚本， (JavaScript 脚本的超集) 。 本文介绍如何使用 Office Scripts for Excel web 版 中的一些内置 JavaScript 对象。

> [!NOTE]
> 有关所有内置 JavaScript 对象的完整列表，请参阅 Mozilla 的 [Standard 内置对象](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) 文章。

## <a name="array"></a>数组

[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)对象提供了一种在脚本中处理数组的标准化方法。 虽然数组是标准 JavaScript 构造，但是它们Office与脚本相关：范围和集合。

### <a name="work-with-ranges"></a>使用区域

区域包含多个二维数组，这些数组直接映射到该范围中的单元格。 这些数组包含有关该范围中每个单元格的特定信息。 例如，返回这些单元格的所有值 (二维数组映射到该工作表子节中的行和列的行和列的行 `Range.getValues`) 。 `Range.getFormulas``Range.getNumberFormats`和 是返回数组的其他常用方法，如 `Range.getValues` 。

以下脚本在 **A1：D4** 范围内搜索包含"$"的任何数字格式。 该脚本将这些单元格中的填充颜色设置为"黄色"。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="work-with-collections"></a>使用集合

集合Excel许多对象。 该集合由 Office 脚本 API 管理，并作为数组公开。 例如，工作表中所有 [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) 都包含在 `Shape[]` 方法返回的 `Worksheet.getShapes` 中。 可以使用此数组读取集合中的值，也可以从父对象的方法访问特定 `get*` 对象。

> [!NOTE]
> 不要手动添加或删除这些集合数组中的对象。 对 `add` 父对象使用 方法，对 `delete` 集合类型对象使用方法。 例如，使用 方法将 [Table](/javascript/api/office-scripts/excelscript/excelscript.table) 添加到 [Worksheet，](/javascript/api/office-scripts/excelscript/excelscript.worksheet) `Worksheet.addTable` 并删除 using `Table` `Table.delete` 。

以下脚本记录当前工作表中每个形状的类型。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

以下脚本删除当前工作表中最早的形状。

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Date

[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)对象提供了一种使用脚本中的日期的标准化方法。 `Date.now()` 生成一个包含当前日期和时间的对象，在向脚本的数据输入中添加时间戳时，这非常有用。

以下脚本将当前日期添加到工作表。 请注意，通过使用 `toLocaleDateString` 方法，Excel值识别为日期并自动更改单元格的编号格式。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

示例 [的"使用日期](../resources/samples/excel-samples.md#dates) "部分具有更多与日期相关的脚本。

## <a name="math"></a>数学

[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)对象提供用于常见数学运算的方法和常量。 这些函数在工作簿中也Excel，无需使用工作簿的计算引擎。 这样一来，脚本就无需查询工作簿，从而提高了性能。

以下脚本用于 `Math.min` 查找和记录 **A1：D4 范围中最小的** 数字。 请注意，此示例假定整个区域仅包含数字，而不包含字符串。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>不支持使用外部 JavaScript 库

Office脚本不支持使用外部第三方库。 脚本只能使用内置的 JavaScript 对象和 Office 脚本 API。

## <a name="see-also"></a>另请参阅

- [标准内置对象](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office脚本代码编辑器环境](../overview/code-editor-environment.md)
