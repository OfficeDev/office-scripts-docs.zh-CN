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
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="14eff-103">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="14eff-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="14eff-104">JavaScript 提供了几个内置对象，您可以在 Office 脚本中使用，而不管您是在 JavaScript 还是使用[TypeScript](../overview/code-editor-environment.md) （javascript 的超集）编写脚本。</span><span class="sxs-lookup"><span data-stu-id="14eff-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="14eff-105">本文介绍如何使用 Office 脚本中的某些内置 JavaScript 对象在 web 上运行 Excel。</span><span class="sxs-lookup"><span data-stu-id="14eff-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="14eff-106">有关所有内置 JavaScript 对象的完整列表，请参阅 Mozilla 的[标准内置对象](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)一文。</span><span class="sxs-lookup"><span data-stu-id="14eff-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="14eff-107">数组</span><span class="sxs-lookup"><span data-stu-id="14eff-107">Array</span></span>

<span data-ttu-id="14eff-108">[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)对象提供了在脚本中使用数组的标准化方法。</span><span class="sxs-lookup"><span data-stu-id="14eff-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="14eff-109">虽然阵列是标准的 JavaScript 构造，但它们与 Office 脚本有以下两种主要的关系：范围和集合。</span><span class="sxs-lookup"><span data-stu-id="14eff-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="14eff-110">处理区域</span><span class="sxs-lookup"><span data-stu-id="14eff-110">Working with ranges</span></span>

<span data-ttu-id="14eff-111">区域包含多个直接映射到该范围中的单元格的二维数组。</span><span class="sxs-lookup"><span data-stu-id="14eff-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="14eff-112">这些数组包含有关该范围中每个单元格的特定信息。</span><span class="sxs-lookup"><span data-stu-id="14eff-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="14eff-113">例如， `Range.getValues` 返回这些单元格中的所有值（二维数组的行和列映射到该工作表子部分的行和列）。</span><span class="sxs-lookup"><span data-stu-id="14eff-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="14eff-114">`Range.getFormulas`以及 `Range.getNumberFormats` 返回像这样的数组的其他频繁使用的方法 `Range.getValues` 。</span><span class="sxs-lookup"><span data-stu-id="14eff-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="14eff-115">下面的脚本在**A1： D4**范围中搜索任何包含 "$" 的数字格式。</span><span class="sxs-lookup"><span data-stu-id="14eff-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="14eff-116">该脚本将这些单元格中的填充颜色设置为 "黄色"。</span><span class="sxs-lookup"><span data-stu-id="14eff-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="14eff-117">使用集合</span><span class="sxs-lookup"><span data-stu-id="14eff-117">Working with collections</span></span>

<span data-ttu-id="14eff-118">集合中包含许多 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="14eff-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="14eff-119">该集合由 Office 脚本 API 管理并作为一个数组公开。</span><span class="sxs-lookup"><span data-stu-id="14eff-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="14eff-120">例如，工作表中的所有[形状](/javascript/api/office-scripts/excelscript/excelscript.shape)都包含在 `Shape[]` 方法返回的中 `Worksheet.getShapes` 。</span><span class="sxs-lookup"><span data-stu-id="14eff-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="14eff-121">您可以使用此数组读取集合中的值，也可以从父对象的方法访问特定的对象 `get*` 。</span><span class="sxs-lookup"><span data-stu-id="14eff-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="14eff-122">请勿手动添加或删除这些集合数组中的对象。</span><span class="sxs-lookup"><span data-stu-id="14eff-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="14eff-123">`add`对父对象和 `delete` 集合类型对象上的方法使用方法。</span><span class="sxs-lookup"><span data-stu-id="14eff-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="14eff-124">例如，使用方法向[工作表](/javascript/api/office-scripts/excelscript/excelscript.worksheet)中添加[表](/javascript/api/office-scripts/excelscript/excelscript.table) `Worksheet.addTable` 并删除 `Table` using `Table.delete` 。</span><span class="sxs-lookup"><span data-stu-id="14eff-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="14eff-125">下面的脚本记录当前工作表中的每个形状的类型。</span><span class="sxs-lookup"><span data-stu-id="14eff-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="14eff-126">下面的脚本删除当前工作表中最旧的形状。</span><span class="sxs-lookup"><span data-stu-id="14eff-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="14eff-127">日期</span><span class="sxs-lookup"><span data-stu-id="14eff-127">Date</span></span>

<span data-ttu-id="14eff-128">[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)对象提供处理脚本中的日期的标准化方法。</span><span class="sxs-lookup"><span data-stu-id="14eff-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="14eff-129">`Date.now()`生成具有当前日期和时间的对象，这在向脚本的数据输入中添加时间戳时非常有用。</span><span class="sxs-lookup"><span data-stu-id="14eff-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="14eff-130">下面的脚本将当前日期添加到工作表中。</span><span class="sxs-lookup"><span data-stu-id="14eff-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="14eff-131">请注意，通过使用 `toLocaleDateString` 方法，Excel 会将值识别为日期，并自动更改单元格的数字格式。</span><span class="sxs-lookup"><span data-stu-id="14eff-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="14eff-132">示例中的 "[处理日期](../resources/excel-samples.md#dates)" 部分具有与日期相关的脚本。</span><span class="sxs-lookup"><span data-stu-id="14eff-132">The [Work with dates](../resources/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="14eff-133">数学</span><span class="sxs-lookup"><span data-stu-id="14eff-133">Math</span></span>

<span data-ttu-id="14eff-134">[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)对象为常见的数学运算提供了方法和常量。</span><span class="sxs-lookup"><span data-stu-id="14eff-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="14eff-135">这些功能在 Excel 中也可以提供许多功能，而无需使用工作簿的计算引擎。</span><span class="sxs-lookup"><span data-stu-id="14eff-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="14eff-136">这将使您的脚本不必查询工作簿，从而提高性能。</span><span class="sxs-lookup"><span data-stu-id="14eff-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="14eff-137">下面的脚本使用 `Math.min` 来查找并记录**A1： D4**范围中的最小数字。</span><span class="sxs-lookup"><span data-stu-id="14eff-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="14eff-138">请注意，此示例假定整个区域仅包含数字，而不包含字符串。</span><span class="sxs-lookup"><span data-stu-id="14eff-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="14eff-139">不支持使用外部 JavaScript 库</span><span class="sxs-lookup"><span data-stu-id="14eff-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="14eff-140">Office 脚本不支持使用外部第三方库。</span><span class="sxs-lookup"><span data-stu-id="14eff-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="14eff-141">您的脚本只能使用内置 JavaScript 对象和 Office 脚本 Api。</span><span class="sxs-lookup"><span data-stu-id="14eff-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="14eff-142">另请参阅</span><span class="sxs-lookup"><span data-stu-id="14eff-142">See also</span></span>

- [<span data-ttu-id="14eff-143">标准内置对象</span><span class="sxs-lookup"><span data-stu-id="14eff-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="14eff-144">Office 脚本代码编辑器环境</span><span class="sxs-lookup"><span data-stu-id="14eff-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
