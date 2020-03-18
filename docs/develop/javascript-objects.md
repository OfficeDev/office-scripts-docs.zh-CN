---
title: 在 Office 脚本中使用内置 JavaScript 对象
description: 如何：从 web 上的 Excel 中的 Office 脚本中调用内置 JavaScript Api。
ms.date: 01/21/2020
localization_priority: Normal
ms.openlocfilehash: e0fcd98117125ead18e55675e195415ff59c0c5d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700113"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>在 Office 脚本中使用内置 JavaScript 对象

JavaScript 提供了几个内置对象，您可以在 Office 脚本中使用，而不管您是在 JavaScript 还是使用[TypeScript](../overview/code-editor-environment.md) （javascript 的超集）编写脚本。 本文介绍如何使用 Office 脚本中的某些内置 JavaScript 对象在 web 上运行 Excel。

> [!NOTE]
> 有关所有内置 JavaScript 对象的完整列表，请参阅 Mozilla 的[标准内置对象](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)一文。

## <a name="array"></a>数组

[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)对象提供了在脚本中使用数组的标准化方法。 虽然阵列是标准的 JavaScript 构造，但它们与 Office 脚本有以下两种主要的关系：范围和集合。

### <a name="working-with-ranges"></a>处理区域

区域包含多个直接映射到该范围中的单元格的二维数组。 其中包括`values`、 `formulas`和`numberFormat`等属性。 数组类型属性的[加载](scripting-fundamentals.md#sync-and-load)方式必须与任何其他属性一样。

下面的脚本在**A1： D4**范围中搜索任何包含 "$" 的数字格式。 该脚本将这些单元格中的填充颜色设置为 "黄色"。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range From A1 to D4.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");

  // Load the numberFormat property on the range.
  range.load("numberFormat");
  await context.sync();

  // Iterate through the arrays of rows and columns corresponding to those in the range.
  range.numberFormat.forEach((rowItem, rowIndex) => {
    range.numberFormat[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).format.fill.color = "yellow";
      }
    });
  });
}
```

### <a name="working-with-collections"></a>使用集合

集合中包含许多 Excel 对象。 例如，工作表中的所有[形状](/javascript/api/office-scripts/excel/excel.shape)都包含在[ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection)中（作为`Worksheet.shapes`属性）。 每`*Collection`个对象都`items`包含一个属性，该属性是一个存储该集合中的对象的数组。 这可以像常规 JavaScript 数组一样进行处理，但必须首先加载集合中的项目。 如果需要在集合中的每个对象上使用属性，请使用分层加载语句（`items/propertyName`）。

下面的脚本记录当前工作表中的每个形状的类型。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.shapes;
  shapes.load("items/type");
  await context.sync();

  // Log the type of every shape in the collection.
  shapes.items.forEach((shape) => {
    console.log(shape.type);
  });
}
```

您可以使用`getItem`或`getItemAt`方法从集合中加载单个对象。 `getItem`通过使用唯一标识符（如名称）获取对象（这些名称通常由脚本指定）。 `getItemAt`通过使用其在集合中的索引获取对象。 在可以使用该对象之前， `await context.sync();`必须先调用一个命令。

下面的脚本删除当前工作表中最旧的形状。

```Typescript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.shapes.getItemAt(0);

  // Sync to load `shape` from the collection.
  await context.sync();

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Date

[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)对象提供处理脚本中的日期的标准化方法。 `Date.now()`生成具有当前日期和时间的对象，这在向脚本的数据输入中添加时间戳时非常有用。

下面的脚本将当前日期添加到工作表中。 请注意，通过使用`toLocaleDateString`方法，Excel 会将值识别为日期，并自动更改单元格的数字格式。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range for cell A1.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.values = [[date.toLocaleDateString()]];
}
```

## <a name="math"></a>数学

[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)对象为常见的数学运算提供了方法和常量。 这些功能在 Excel 中也可以提供许多功能，而无需使用工作簿的计算引擎。 这将使您的脚本不必查询工作簿，从而提高性能。

下面的脚本使用`Math.min`来查找并记录**A1： D4**范围中的最小数字。 请注意，此示例假定整个区域仅包含数字，而不包含字符串。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range from A1 to D4.
  let comparisonRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");
  
  // Load the range's values.
  comparisonRange.load("values");
  await context.sync();

  // Set the minimum values as the first value.
  let minimum = comparisonRange.values[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRange.values.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRange.values[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });
  
  console.log(minimum);
}

```

## <a name="see-also"></a>另请参阅

- [标准内置对象](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office 脚本代码编辑器环境](../overview/code-editor-environment.md)
