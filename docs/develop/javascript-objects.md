---
title: 在 Office 脚本中使用内置的 JavaScript 对象
description: 如何从 Excel web 版 中的 Office 脚本调用内置 JavaScript EXCEL WEB 版。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545045"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="f82e2-103">在脚本中使用内置的 JavaScript Office对象</span><span class="sxs-lookup"><span data-stu-id="f82e2-103">Use built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="f82e2-104">JavaScript 提供了多个可用于 Office 脚本的内置对象，无论你是使用 JavaScript 还是[TypeScript](../overview/code-editor-environment.md)编写脚本， (JavaScript 脚本的超集) 。</span><span class="sxs-lookup"><span data-stu-id="f82e2-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="f82e2-105">本文介绍如何使用 Office Scripts for Excel web 版 中的一些内置 JavaScript 对象。</span><span class="sxs-lookup"><span data-stu-id="f82e2-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="f82e2-106">有关所有内置 JavaScript 对象的完整列表，请参阅 Mozilla 的 [Standard 内置对象](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) 文章。</span><span class="sxs-lookup"><span data-stu-id="f82e2-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="f82e2-107">数组</span><span class="sxs-lookup"><span data-stu-id="f82e2-107">Array</span></span>

<span data-ttu-id="f82e2-108">[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)对象提供了一种在脚本中处理数组的标准化方法。</span><span class="sxs-lookup"><span data-stu-id="f82e2-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="f82e2-109">虽然数组是标准 JavaScript 构造，但是它们Office与脚本相关：范围和集合。</span><span class="sxs-lookup"><span data-stu-id="f82e2-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="work-with-ranges"></a><span data-ttu-id="f82e2-110">使用区域</span><span class="sxs-lookup"><span data-stu-id="f82e2-110">Work with ranges</span></span>

<span data-ttu-id="f82e2-111">区域包含多个二维数组，这些数组直接映射到该范围中的单元格。</span><span class="sxs-lookup"><span data-stu-id="f82e2-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="f82e2-112">这些数组包含有关该范围中每个单元格的特定信息。</span><span class="sxs-lookup"><span data-stu-id="f82e2-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="f82e2-113">例如，返回这些单元格的所有值 (二维数组映射到该工作表子节中的行和列的行和列 `Range.getValues`) 。</span><span class="sxs-lookup"><span data-stu-id="f82e2-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="f82e2-114">`Range.getFormulas``Range.getNumberFormats`和 是返回数组的其他常用方法，如 `Range.getValues` 。</span><span class="sxs-lookup"><span data-stu-id="f82e2-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="f82e2-115">以下脚本在 **A1：D4** 范围内搜索包含"$"的任何数字格式。</span><span class="sxs-lookup"><span data-stu-id="f82e2-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="f82e2-116">该脚本将这些单元格中的填充颜色设置为"黄色"。</span><span class="sxs-lookup"><span data-stu-id="f82e2-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="work-with-collections"></a><span data-ttu-id="f82e2-117">使用集合</span><span class="sxs-lookup"><span data-stu-id="f82e2-117">Work with collections</span></span>

<span data-ttu-id="f82e2-118">许多Excel对象都包含在集合中。</span><span class="sxs-lookup"><span data-stu-id="f82e2-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="f82e2-119">该集合由 Office 脚本 API 管理，并作为数组公开。</span><span class="sxs-lookup"><span data-stu-id="f82e2-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="f82e2-120">例如，工作表中所有 [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) 都包含在 `Shape[]` 方法返回的 `Worksheet.getShapes` 中。</span><span class="sxs-lookup"><span data-stu-id="f82e2-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="f82e2-121">可以使用此数组读取集合中的值，也可以从父对象的方法访问特定 `get*` 对象。</span><span class="sxs-lookup"><span data-stu-id="f82e2-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="f82e2-122">不要手动添加或删除这些集合数组中的对象。</span><span class="sxs-lookup"><span data-stu-id="f82e2-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="f82e2-123">对 `add` 父对象使用 方法，对 `delete` 集合类型对象使用方法。</span><span class="sxs-lookup"><span data-stu-id="f82e2-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="f82e2-124">例如，使用 方法将 [Table](/javascript/api/office-scripts/excelscript/excelscript.table) 添加到 [Worksheet，](/javascript/api/office-scripts/excelscript/excelscript.worksheet) `Worksheet.addTable` 并删除 using `Table` `Table.delete` 。</span><span class="sxs-lookup"><span data-stu-id="f82e2-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="f82e2-125">以下脚本记录当前工作表中每个形状的类型。</span><span class="sxs-lookup"><span data-stu-id="f82e2-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="f82e2-126">以下脚本删除当前工作表中最早的形状。</span><span class="sxs-lookup"><span data-stu-id="f82e2-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="f82e2-127">日期</span><span class="sxs-lookup"><span data-stu-id="f82e2-127">Date</span></span>

<span data-ttu-id="f82e2-128">[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)对象提供了一种使用脚本中的日期的标准化方法。</span><span class="sxs-lookup"><span data-stu-id="f82e2-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="f82e2-129">`Date.now()` 生成一个包含当前日期和时间的对象，在向脚本的数据输入中添加时间戳时，这非常有用。</span><span class="sxs-lookup"><span data-stu-id="f82e2-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="f82e2-130">以下脚本将当前日期添加到工作表。</span><span class="sxs-lookup"><span data-stu-id="f82e2-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="f82e2-131">请注意，通过使用 `toLocaleDateString` 方法，Excel值识别为日期并自动更改单元格的编号格式。</span><span class="sxs-lookup"><span data-stu-id="f82e2-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="f82e2-132">示例 [的"使用日期](../resources/samples/excel-samples.md#dates) "部分具有更多与日期相关的脚本。</span><span class="sxs-lookup"><span data-stu-id="f82e2-132">The [Work with dates](../resources/samples/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="f82e2-133">数学</span><span class="sxs-lookup"><span data-stu-id="f82e2-133">Math</span></span>

<span data-ttu-id="f82e2-134">[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)对象提供用于常见数学运算的方法和常量。</span><span class="sxs-lookup"><span data-stu-id="f82e2-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="f82e2-135">这些函数也提供许多Excel，而无需使用工作簿的计算引擎。</span><span class="sxs-lookup"><span data-stu-id="f82e2-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="f82e2-136">这样一来，脚本就无需查询工作簿，从而提高了性能。</span><span class="sxs-lookup"><span data-stu-id="f82e2-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="f82e2-137">以下脚本用于 `Math.min` 查找和记录 **A1：D4 范围中最小的** 数字。</span><span class="sxs-lookup"><span data-stu-id="f82e2-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="f82e2-138">请注意，此示例假定整个区域仅包含数字，而不包含字符串。</span><span class="sxs-lookup"><span data-stu-id="f82e2-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="f82e2-139">不支持使用外部 JavaScript 库</span><span class="sxs-lookup"><span data-stu-id="f82e2-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="f82e2-140">Office脚本不支持使用外部第三方库。</span><span class="sxs-lookup"><span data-stu-id="f82e2-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="f82e2-141">脚本只能使用内置的 JavaScript 对象和 Office 脚本 API。</span><span class="sxs-lookup"><span data-stu-id="f82e2-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="f82e2-142">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f82e2-142">See also</span></span>

- [<span data-ttu-id="f82e2-143">标准内置对象</span><span class="sxs-lookup"><span data-stu-id="f82e2-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="f82e2-144">Office脚本代码编辑器环境</span><span class="sxs-lookup"><span data-stu-id="f82e2-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
