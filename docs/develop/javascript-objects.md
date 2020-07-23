---
title: 在 Office 脚本中使用内置的 JavaScript 对象
description: 如何：从 web 上的 Excel 中的 Office 脚本中调用内置 JavaScript Api。
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: 4bb5fb5444887005ececbbfdf0130cba3784e0c4
ms.sourcegitcommit: 8d549884e68170f808d3d417104a4451a37da83c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2020
ms.locfileid: "45229594"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>在 Office 脚本中使用内置的 JavaScript 对象

JavaScript 提供了几个内置对象，您可以在 Office 脚本中使用，而不管您是在 JavaScript 还是使用[TypeScript](../overview/code-editor-environment.md) （javascript 的超集）编写脚本。 本文介绍如何使用 Office 脚本中的某些内置 JavaScript 对象在 web 上运行 Excel。

> [!NOTE]
> 有关所有内置 JavaScript 对象的完整列表，请参阅 Mozilla 的[标准内置对象](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)一文。

## <a name="array"></a>数组

[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)对象提供了在脚本中使用数组的标准化方法。 虽然阵列是标准的 JavaScript 构造，但它们与 Office 脚本有以下两种主要的关系：范围和集合。

### <a name="working-with-ranges"></a>处理区域

区域包含多个直接映射到该范围中的单元格的二维数组。 这些数组包含有关该范围中每个单元格的特定信息。 例如， `Range.getValues` 返回这些单元格中的所有值（二维数组的行和列映射到该工作表子部分的行和列）。 `Range.getFormulas`以及 `Range.getNumberFormats` 返回像这样的数组的其他频繁使用的方法 `Range.getValues` 。

下面的脚本在**A1： D4**范围中搜索任何包含 "$" 的数字格式。 该脚本将这些单元格中的填充颜色设置为 "黄色"。

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

### <a name="working-with-collections"></a>使用集合

集合中包含许多 Excel 对象。 该集合由 Office 脚本 API 管理并作为一个数组公开。 例如，工作表中的所有[形状](/javascript/api/office-scripts/excelscript/excelscript.shape)都包含在 `Shape[]` 方法返回的中 `Worksheet.getShapes` 。 您可以使用此数组读取集合中的值，也可以从父对象的方法访问特定的对象 `get*` 。

> [!NOTE]
> 请勿手动添加或删除这些集合数组中的对象。 `add`对父对象和 `delete` 集合类型对象上的方法使用方法。 例如，使用方法向[工作表](/javascript/api/office-scripts/excelscript/excelscript.worksheet)中添加[表](/javascript/api/office-scripts/excelscript/excelscript.table) `Worksheet.addTable` 并删除 `Table` using `Table.delete` 。

下面的脚本记录当前工作表中的每个形状的类型。

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

下面的脚本删除当前工作表中最旧的形状。

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

## <a name="date"></a>日期

[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)对象提供处理脚本中的日期的标准化方法。 `Date.now()`生成具有当前日期和时间的对象，这在向脚本的数据输入中添加时间戳时非常有用。

下面的脚本将当前日期添加到工作表中。 请注意，通过使用 `toLocaleDateString` 方法，Excel 会将值识别为日期，并自动更改单元格的数字格式。

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

示例中的 "[处理日期](../resources/excel-samples.md#dates)" 部分具有与日期相关的脚本。

## <a name="math"></a>数学

[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)对象为常见的数学运算提供了方法和常量。 这些功能在 Excel 中也可以提供许多功能，而无需使用工作簿的计算引擎。 这将使您的脚本不必查询工作簿，从而提高性能。

下面的脚本使用 `Math.min` 来查找并记录**A1： D4**范围中的最小数字。 请注意，此示例假定整个区域仅包含数字，而不包含字符串。

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

Office 脚本不支持使用外部第三方库。 您的脚本只能使用内置 JavaScript 对象和 Office 脚本 Api。

## <a name="see-also"></a>另请参阅

- [标准内置对象](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office 脚本代码编辑器环境](../overview/code-editor-environment.md)
