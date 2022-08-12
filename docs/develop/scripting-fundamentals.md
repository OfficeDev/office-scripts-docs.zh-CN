---
title: Excel 网页版中 Office 脚本的脚本基础
description: 在编写 Office 脚本之前需要了解的对象模型信息和其他基础知识。
ms.date: 05/24/2021
ms.localizationpriority: high
ms.openlocfilehash: 97aa840809010f3640b045ce2fd28a39a47243b4
ms.sourcegitcommit: 33fe0f6807daefb16b148fd73c863de101f47cea
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/08/2022
ms.locfileid: "67281922"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a>Excel 网页版中 Office 脚本的脚本基础

本文将介绍 Office 脚本技术方面的知识。 你将了解基于 TypeScript 的脚本代码的关键部分，以及 Excel 对象和 API 如何协同工作。

## <a name="typescript-the-language-of-office-scripts"></a>TypeScript：Office 脚本的语言

Office 脚本以 [TypeScript](https://www.typescriptlang.org/docs/home.html) 编写，它是 [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) 的一个超集。 如果熟悉 JavaScript，你的知识将会延续下去，因为两种语言的大部分代码是相同的。 在开始 Office 脚本编码之旅之前，我们建议你先掌握一些初级编程知识。 以下资源可以帮助理解 Office 脚本的编码方面。

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a>`main` 函数：脚本的起点

每个脚本都必须包含一个 `main` 函数，并以 `ExcelScript.Workbook` 类型作为第一个参数。 函数运行时，Excel 应用程序通过提供工作簿作为第一个参数来调用 `main` 函数。 `ExcelScript.Workbook` 应始终是第一个参数。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

运行脚本时，`main` 函数中的代码将运行。 `main` 可以调用脚本中的其他函数，但是该函数中未包含的代码将不会运行。 脚本无法调用其他 Office 脚本。

通过 [Power Automate](https://flow.microsoft.com)，可以在流中连接脚本。 数据通过 `main` 函数的参数和返回在脚本和流之间传递。 [使用 Power Automate 运行 Office 脚本](power-automate-integration.md) 中详细介绍了如何集成 Office 脚本和 Power Automate。

## <a name="object-model-overview"></a>对象模型概述

要编写脚本，需要了解 Office 脚本 API 的组合方式。 工作簿的组件之间彼此有着特定的关系。 这些关系在许多方面与 Excel UI 的关系匹配。

- 一个 **Workbook** 包含一个或多个 **Worksheet**。
- **Worksheet** 可通过 **Range** 对象访问单元格。
- **Range** 代表一组连续的单元格。
- **Range** 用于创建和放置 **Table**、**Chart** 和 **Shape** 以及其他数据可视化或组织对象。
- **Worksheet** 包含单个工作表中存在的那些数据对象的集合。
- **Workbook** 包含整个 **Workbook** 的某些数据对象（例如，**Table**）的集合。

## <a name="workbook"></a>工作簿

每个脚本都会由 `main` 函数提供一个 `Workbook` 类型的 `workbook` 对象。 这表示顶层对象，你的脚本将通过该对象与 Excel 工作簿进行交互。

以下脚本将获取工作簿中的活动工作表并记录其名称。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a>Ranges

范围是工作簿中的一组连续单元格。脚本通常使用 A1 样式表示法 (例如，对于列 **B** 和行 **3** 中的单个单元格 **B3**，或者从列 **C** 到 **F** 和行 **2** 到 **4** 的单元格 **C2:F4**) 定义范围。

Range 有三个核心属性：值、公式和格式。 这些属性将获取或设置单元格值、要计算的公式以及单元格的视觉对象格式。 它们可通过 `getValues`、`getFormulas` 和 `getFormat` 进行访问。 值和公式可通过 `setValues` 和 `setFormulas` 进行更改，而格式则是由单独设置的多个较小对象组成的 `RangeFormat` 对象。

Range 使用二维数组管理信息。 有关在 Office 脚本框架中处理数组的详细信息，请参阅 [使用范围工作](javascript-objects.md#work-with-ranges)。

### <a name="range-sample"></a>Range 示例

以下示例显示了如何创建销售记录。 该脚本使用 `Range` 对象来设置值、公式和部分格式。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [["Product", "Quantity", "Unit Price", "Totals"]];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.setValues(headers);
    headerRange.getFormat().getFill().setColor("#4472C4");
    headerRange.getFormat().getFont().setColor("white");

    // Create the product data rows.
    let productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.54],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.setValues(productData);

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"],
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.setFormulas(totalFormulas);
    totalRange.getFormat().getFont().setBold(true);

    // Display the totals as US dollar amounts.
    totalRange.setNumberFormat("$0.00");
}
```

运行此脚本将在当前工作表中创建以下数据：

:::image type="content" source="../images/range-sample.png" alt-text="包含由值行、公式列和带格式的标头组成的销售记录的工作表。":::

### <a name="the-types-of-range-values"></a>范围值的类型

每个单元格都有值。 该值是输入到单元格中的基础值，可能不同于 Excel 中显示的文本。 例如，你可能会看到单元格中的日期显示为“5/2/2021”，但实际值为 44318。 可以使用数字格式更改此显示，但是单元格的实际值和键入内容仅在设置新值时才会发生变化。

使用单元格值时，请告诉 TypeScript 期望从单元格或范围中获得什么值，这一点很重要。 包含以下其中一个类型的单元格：`string`、`number` 或 `boolean`。 为了让脚本将返回的值作为其中一种类型的值，必须声明类型。

以下脚本从上一个示例中的表格中获取平均价格。 为代码添加备注`priceRange.getValues() as number[][]`。 这段代码将范围值的类型[声明](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions)为`number[][]`。 然后，该数组中的所有值都可以视为脚本中的数字。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet.
  let sheet = workbook.getActiveWorksheet();

  // Get the "Unit Price" column. 
  // The result of calling getValues is declared to be a number[][] so that we can perform arithmetic operations.
  let priceRange = sheet.getRange("D3:D5");
  let prices = priceRange.getValues() as number[][];

  // Get the average price.
  let totalPrices = 0;
  prices.forEach((price) => totalPrices += price[0]);
  let averagePrice = totalPrices / prices.length;
  console.log(averagePrice);
}
```

## <a name="charts-tables-and-other-data-objects"></a>Chart、Table 和其他数据对象

脚本可以在 Excel 中创建和设置数据结构和可视化效果。 Table 和 Chart 是最常用的两个对象，但是 API 支持数据透视表、形状和图像等。 这些都存储在集合中，本文后面将对该内容进行讨论。

### <a name="create-a-table"></a>创建表格

通过使用数据填充区域创建表。自动将格式设置和表格控件（如筛选器）应用到区域。

以下脚本使用上一个示例中的范围创建一个表。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

在工作表上使用之前的数据运行此脚本将创建下表：

:::image type="content" source="../images/table-sample.png" alt-text="包含根据以前销售记录所创建表的工作表。":::

### <a name="create-a-chart"></a>创建图表

创建图表以直观显示某个范围内的数据。 脚本支持数十种图表类型，每种都可以根据需要进行自定义。

下面的脚本为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方，并将其设置为 100 像素。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create a column chart using the data from B3:C5.
    let chart = sheet.addChart(
        ExcelScript.ChartType.columnStacked,
        sheet.getRange("B3:C5")
    );

    // Set the margin of the chart to be 100 pixels from the top of the screen.
    chart.setTop(100);
}
```

在工作表上使用上一个表运行此脚本将创建以下图表：

:::image type="content" source="../images/chart-sample.png" alt-text="一个柱形图，显示上一个销售记录中三个项目的数量。":::

## <a name="collections"></a>集合

当 Excel 对象具有一个或多个相同类型对象的集合时，则将它们存储在数组中。 例如，`Workbook` 对象包含一个 `Worksheet[]`。 此数组由 `Workbook.getWorksheets()` 方法访问。 复数形式的 `get` 方法（如 `Worksheet.getCharts()`）将整个对象集合作为数组返回。 你将在整个 Office 脚本 API 中查看此模式：`Worksheet` 对象采用 `getTables()` 方法返回 `Table[]`，`Table` 对象采用 `getColumns()` 方法返回 `TableColumn[]`，以此类推。

返回的数组是一个普通数组，因此所有常规数组操作均可用于脚本。 你还可以使用数组索引值访问集合中的单个对象。 例如，`workbook.getTables()[0]` 将返回集合中的第一个表格。 有关通过 Office 脚本框架使用内置数组功能的详细信息，请参阅 [使用集合工作](javascript-objects.md#work-with-collections)。

此外，还可通过 `get` 方法从集合中访问单个对象。 单数形式的 `get` 方法（如 `Worksheet.getTable(name)`）返回单个对象，并且需要特定对象的 ID 或名称。 此 ID 或名称通常由脚本或通过 Excel UI 设置。

以下脚本获取工作簿中所有表。然后可确保显示标题、筛选按钮可见，并且表格样式设置为“TableStyleLight1”。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table collection.
  let tables = workbook.getTables();

  // Set the table formatting properties for every table.
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

## <a name="add-excel-objects-with-a-script"></a>使用脚本添加 Excel 对象

通过调用可在父对象上使用的相应 `add` 方法，可以以编程方式添加文档对象，如表格或图表。

> [!IMPORTANT]
> 不要手动将对象添加到集合数组。 请在父对象上使用 `add` 方法。例如，使用 `Worksheet.addTable` 方法向 `Worksheet` 添加 `Table`。

以下脚本将在 Excel 工作簿中的第一个工作表上创建一个表格。 请注意，所创建的表格是通过 `addTable` 方法返回的。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in A1:G10.
    let table = sheet.addTable(
      "A1:G10",
       true /* True because the table has headers. */
    );
    
    // Give the table a name for easy reference in other scripts.
    table.setName("MyTable");
}
```

> [!TIP]
> 大多数 Excel 对象都具有 `setName` 方法。通过这一方法，可稍后在同一工作簿的脚本或其他脚本中轻松访问 Excel 对象。

### <a name="verify-an-object-exists-in-the-collection"></a>验证集合中是否存在某个对象

在继续之前，脚本通常需要检查表或类似对象是否存在。 使用脚本或 Excel UI 提供的名称确定必要的对象，并执行相应操作。 请求的对象不在集合中时，`get` 方法返回 `undefined`。

以下脚本请求名为“MyTable”的表，并使用 `if...else` 语句检查是否已找到该表。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable".
  let myTable = workbook.getTable("MyTable");

  // If the table is in the workbook, myTable will have a value.
  // Otherwise, the variable will be undefined and go to the else clause.
  if (myTable) {
    let worksheetName = myTable.getWorksheet().getName();
    console.log(`MyTable is on the ${worksheetName} worksheet`);
  } else {
    console.log(`MyTable is not in the workbook.`);
  }
}
```

Office 脚本中的一种常见模式是在每次运行脚本时重新创建表、图表或其他对象。 如果不需要旧数据，最好先删除旧对象，然后再创建新对象。 此操作可避免出现名称冲突或已由其他用户引入的其他差异。

以下脚本删除名为“MyTable”的表，如果存在该表，则添加名称相同的新表。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable" from the first worksheet.
  let sheet = workbook.getWorksheets()[0];
  let tableName = "MyTable";
  let oldTable = sheet.getTable(tableName);

  // If the table exists, remove it.
  if (oldTable) {
    oldTable.delete();
  }

  // Add a new table with the same name.
  let newTable = sheet.addTable("A1:G10", true);
  newTable.setName(tableName);
}
```

## <a name="remove-excel-objects-with-a-script"></a>使用脚本删除 Excel 对象

若要删除对象，请调用对象的 `delete` 方法。

> [!NOTE]
> 与添加对象一样，不要手动从集合数组中删除对象。 请在集合类型的对象上使用 `delete` 方法。 例如，使用 `Table.delete`从 `Worksheet` 中删除 `Table`。

以下脚本将删除工作簿中的第一个工作表。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a>进一步了解对象模型

[Office 脚本 API 参考文档](/javascript/api/office-scripts/overview)是 Office 脚本中使用的对象的完整列表。 在这里，可以使用目录导航到想进一步了解的任何课程。 以下是几个经常查看的页面。

- [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [Comment](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [区域](/javascript/api/office-scripts/excelscript/excelscript.range)
- [RangeFormat](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [Shape](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [Table](/javascript/api/office-scripts/excelscript/excelscript.table)
- [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

有关特定于数据透视表对象模型的信息，请参阅[在 Office 脚本中使用数据透视表](pivottables.md)。

## <a name="see-also"></a>另请参阅

- [在 Excel 网页版中录制、编辑和创建 Office 脚本](../tutorials/excel-tutorial.md)
- [在 Excel 网页版中使用 Office 脚本读取工作簿数据](../tutorials/excel-read-tutorial.md)
- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
- [在 Office 脚本中使用数据透视表](pivottables.md)
- [在 Office 脚本中使用内置的 JavaScript 对象](javascript-objects.md)
- [Office 脚本中的最佳实践](best-practices.md)
- [Office 脚本开发中心](https://developer.microsoft.com/office-scripts)
