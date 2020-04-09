---
title: 在 Office 脚本中使用内置的 JavaScript 对象
description: 如何：从 web 上的 Excel 中的 Office 脚本中调用内置 JavaScript Api。
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: a4b698215edea5f266e159fee0e08690904dd379
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191016"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="383a0-103">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="383a0-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="383a0-104">JavaScript 提供了几个内置对象，您可以在 Office 脚本中使用，而不管您是在 JavaScript 还是使用[TypeScript](../overview/code-editor-environment.md) （javascript 的超集）编写脚本。</span><span class="sxs-lookup"><span data-stu-id="383a0-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="383a0-105">本文介绍如何使用 Office 脚本中的某些内置 JavaScript 对象在 web 上运行 Excel。</span><span class="sxs-lookup"><span data-stu-id="383a0-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="383a0-106">有关所有内置 JavaScript 对象的完整列表，请参阅 Mozilla 的[标准内置对象](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)一文。</span><span class="sxs-lookup"><span data-stu-id="383a0-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="383a0-107">数组</span><span class="sxs-lookup"><span data-stu-id="383a0-107">Array</span></span>

<span data-ttu-id="383a0-108">[Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)对象提供了在脚本中使用数组的标准化方法。</span><span class="sxs-lookup"><span data-stu-id="383a0-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="383a0-109">虽然阵列是标准的 JavaScript 构造，但它们与 Office 脚本有以下两种主要的关系：范围和集合。</span><span class="sxs-lookup"><span data-stu-id="383a0-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="383a0-110">处理区域</span><span class="sxs-lookup"><span data-stu-id="383a0-110">Working with ranges</span></span>

<span data-ttu-id="383a0-111">区域包含多个直接映射到该范围中的单元格的二维数组。</span><span class="sxs-lookup"><span data-stu-id="383a0-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="383a0-112">其中包括`values`、 `formulas`和`numberFormat`等属性。</span><span class="sxs-lookup"><span data-stu-id="383a0-112">These include properties such as `values`, `formulas`, and `numberFormat`.</span></span> <span data-ttu-id="383a0-113">数组类型属性的[加载](scripting-fundamentals.md#sync-and-load)方式必须与任何其他属性一样。</span><span class="sxs-lookup"><span data-stu-id="383a0-113">Array-type properties must be [loaded](scripting-fundamentals.md#sync-and-load) like any other properties.</span></span>

<span data-ttu-id="383a0-114">下面的脚本在**A1： D4**范围中搜索任何包含 "$" 的数字格式。</span><span class="sxs-lookup"><span data-stu-id="383a0-114">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="383a0-115">该脚本将这些单元格中的填充颜色设置为 "黄色"。</span><span class="sxs-lookup"><span data-stu-id="383a0-115">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="383a0-116">使用集合</span><span class="sxs-lookup"><span data-stu-id="383a0-116">Working with collections</span></span>

<span data-ttu-id="383a0-117">集合中包含许多 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="383a0-117">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="383a0-118">例如，工作表中的所有[形状](/javascript/api/office-scripts/excel/excel.shape)都包含在[ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection)中（作为`Worksheet.shapes`属性）。</span><span class="sxs-lookup"><span data-stu-id="383a0-118">For example, all [Shapes](/javascript/api/office-scripts/excel/excel.shape) in a worksheet are contained in a [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (as the `Worksheet.shapes` property).</span></span> <span data-ttu-id="383a0-119">每`*Collection`个对象都`items`包含一个属性，该属性是一个存储该集合中的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="383a0-119">Each `*Collection` object contains an `items` property, which is an array that stores the objects inside that collection.</span></span> <span data-ttu-id="383a0-120">这可以像常规 JavaScript 数组一样进行处理，但必须首先加载集合中的项目。</span><span class="sxs-lookup"><span data-stu-id="383a0-120">This can be treated like a normal JavaScript array, but the items in the collection have to first be loaded.</span></span> <span data-ttu-id="383a0-121">如果需要在集合中的每个对象上使用属性，请使用分层加载语句（`items/propertyName`）。</span><span class="sxs-lookup"><span data-stu-id="383a0-121">If you need to work with a property on every object in the collection, use a hierarchal load statement (`items/propertyName`).</span></span>

<span data-ttu-id="383a0-122">下面的脚本记录当前工作表中的每个形状的类型。</span><span class="sxs-lookup"><span data-stu-id="383a0-122">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="383a0-123">您可以使用`getItem`或`getItemAt`方法从集合中加载单个对象。</span><span class="sxs-lookup"><span data-stu-id="383a0-123">You can load individual objects from a collection using the `getItem` or `getItemAt` methods.</span></span> <span data-ttu-id="383a0-124">`getItem`通过使用唯一标识符（如名称）获取对象（这些名称通常由脚本指定）。</span><span class="sxs-lookup"><span data-stu-id="383a0-124">`getItem` gets an object by using a unique identifier like a name (such names are often specified by your script).</span></span> <span data-ttu-id="383a0-125">`getItemAt`通过使用其在集合中的索引获取对象。</span><span class="sxs-lookup"><span data-stu-id="383a0-125">`getItemAt` gets an object by using its index in the collection.</span></span> <span data-ttu-id="383a0-126">在可以使用该对象之前， `await context.sync();`必须先调用一个命令。</span><span class="sxs-lookup"><span data-stu-id="383a0-126">Either call must be followed by a `await context.sync();` command before the object can be used.</span></span>

<span data-ttu-id="383a0-127">下面的脚本删除当前工作表中最旧的形状。</span><span class="sxs-lookup"><span data-stu-id="383a0-127">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="383a0-128">日期</span><span class="sxs-lookup"><span data-stu-id="383a0-128">Date</span></span>

<span data-ttu-id="383a0-129">[Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)对象提供处理脚本中的日期的标准化方法。</span><span class="sxs-lookup"><span data-stu-id="383a0-129">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="383a0-130">`Date.now()`生成具有当前日期和时间的对象，这在向脚本的数据输入中添加时间戳时非常有用。</span><span class="sxs-lookup"><span data-stu-id="383a0-130">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="383a0-131">下面的脚本将当前日期添加到工作表中。</span><span class="sxs-lookup"><span data-stu-id="383a0-131">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="383a0-132">请注意，通过使用`toLocaleDateString`方法，Excel 会将值识别为日期，并自动更改单元格的数字格式。</span><span class="sxs-lookup"><span data-stu-id="383a0-132">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="383a0-133">示例中的 "[处理日期](../resources/excel-samples.md#work-with-dates)" 部分具有与日期相关的脚本。</span><span class="sxs-lookup"><span data-stu-id="383a0-133">The [Work with dates](../resources/excel-samples.md#work-with-dates) section of the samples has more Date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="383a0-134">数学</span><span class="sxs-lookup"><span data-stu-id="383a0-134">Math</span></span>

<span data-ttu-id="383a0-135">[Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)对象为常见的数学运算提供了方法和常量。</span><span class="sxs-lookup"><span data-stu-id="383a0-135">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="383a0-136">这些功能在 Excel 中也可以提供许多功能，而无需使用工作簿的计算引擎。</span><span class="sxs-lookup"><span data-stu-id="383a0-136">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="383a0-137">这将使您的脚本不必查询工作簿，从而提高性能。</span><span class="sxs-lookup"><span data-stu-id="383a0-137">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="383a0-138">下面的脚本使用`Math.min`来查找并记录**A1： D4**范围中的最小数字。</span><span class="sxs-lookup"><span data-stu-id="383a0-138">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="383a0-139">请注意，此示例假定整个区域仅包含数字，而不包含字符串。</span><span class="sxs-lookup"><span data-stu-id="383a0-139">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="383a0-140">另请参阅</span><span class="sxs-lookup"><span data-stu-id="383a0-140">See also</span></span>

- [<span data-ttu-id="383a0-141">标准内置对象</span><span class="sxs-lookup"><span data-stu-id="383a0-141">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="383a0-142">Office 脚本代码编辑器环境</span><span class="sxs-lookup"><span data-stu-id="383a0-142">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
