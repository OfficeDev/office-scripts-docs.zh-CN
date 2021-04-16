---
title: Excel 网页版中 Office 脚本的脚本基础
description: 在编写 Office 脚本之前需要了解的对象模型信息和其他基础知识。
ms.date: 07/08/2020
localization_priority: Priority
ms.openlocfilehash: 2c2fd683e77a0dfbfd3e9df8c79db31e78ceee8b
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755061"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Excel 网页版中 Office 脚本的脚本基础（预览）

本文将介绍 Office 脚本技术方面的知识。 你将了解 Excel 对象如何协同工作以及代码编辑器如何与工作簿同步。

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a>`main` 函数

每个 Office 脚本都必须包含以 `ExcelScript.Workbook` 类型作为第一参数的 `main` 函数。 执行函数时，Excel 应用程序通过提供相应工作簿作为第一个参数来调用此 `main` 函数。 因此，在记录脚本或从代码编辑器创建新脚本后，请务必不要再修改 `main` 函数的基本签名。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

运行脚本时，`main` 函数中的代码将运行。 `main` 可以调用脚本中的其他函数，但是该函数中未包含的代码将不会运行。

> [!CAUTION]
> 如果你的 `main` 函数看起来像 `async function main(context: Excel.RequestContext)`，那你的脚本使用的是旧版异步 API 模型。 有关详细信息（包括如何将你的脚本转换为当前 API 模型），请参阅[支持使用异步 API 的旧 Office 脚本](excel-async-model.md)。

## <a name="object-model"></a>对象模型

若要编写脚本，你需要了解 Office 脚本 API 如何组合在一起。 工作簿的组件之间彼此有着特定的关系。 这些关系在许多方面与 Excel UI 的关系匹配。

- 一个 **Workbook** 包含一个或多个 **Worksheet**。
- **Worksheet** 可通过 **Range** 对象访问单元格。
- **Range** 代表一组连续的单元格。
- **Range** 用于创建和放置 **Table**、**Chart** 和 **Shape** 以及其他数据可视化或组织对象。
- **Worksheet** 包含单个工作表中存在的那些数据对象的集合。
- **Workbook** 包含整个 **Workbook** 的某些数据对象（例如，**Table**）的集合。

### <a name="workbook"></a>工作簿

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

### <a name="ranges"></a>Ranges

Range 是工作簿中的一组连续单元格。 脚本通常使用 A1 样式表示法（例如，对于列 **B** 和行 **3** 中单个单元格，即 **B3** 或从列 **C** 至列 **F** 和行 **2** 至行 **4** 的单元格，即 **C2:F4**）来定义范围。

Range 有三个核心属性：值、公式和格式。 这些属性将获取或设置单元格值、要计算的公式以及单元格的视觉对象格式。 它们可通过 `getValues`、`getFormulas` 和 `getFormat` 进行访问。 值和公式可通过 `setValues` 和 `setFormulas` 进行更改，而格式则是由单独设置的多个较小对象组成的 `RangeFormat` 对象。

Range 使用二维数组管理信息。 有关如何在 Office 脚本框架中处理这些数组的详细信息，请参阅[《在 Office 脚本中使用内置的 JavaScript 对象》的“使用区域”部分](javascript-objects.md#working-with-ranges)。

#### <a name="range-sample"></a>Range 示例

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
        ["Chocolate", 10, 9.56],
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

### <a name="charts-tables-and-other-data-objects"></a>Chart、Table 和其他数据对象

脚本可以在 Excel 中创建和设置数据结构和可视化效果。 Table 和 Chart 是最常用的两个对象，但是 API 支持数据透视表、形状和图像等。 这些都存储在集合中，本文后面将对该内容进行讨论。

#### <a name="creating-a-table"></a>创建表

通过使用数据填充范围创建表。 会将格式设置和表控件（如筛选器）自动应用到该范围。

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

#### <a name="creating-a-chart"></a>创建图表

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

### <a name="collections-and-other-object-relations"></a>集合和其他对象关系

任何子对象都可通过其父对象访问。 例如，可从 `Workbook` 对象中读取 `Worksheets`。 父类上将会有一个相关的 `get` 方法（例如 `Workbook.getWorksheets()` 或 `Workbook.getWorksheet(name)` ）。 单数形式的 `get` 方法将返回单个对象，并且需要特定对象的 ID 或名称（如工作表名称）。 复数形式的 `get` 方法会将整个对象集合作为数组返回。 如果集合为空，将得到一个空数组 (`[]`)。

检索到相应集合后，可在其上面使用常规数组操作（如获取其 `length` 或使用 `for`、`for..of` 或 `while` 循环进行迭代）或使用 TypeScript 数组方法（如 `map` 或 `forEach`）。 你还可以使用数组索引值访问集合中的单个对象。 例如，`workbook.getTables()[0]` 将返回集合中的第一个表格。 请阅读[《在 Office 脚本中使用内置的 JavaScript 对象》的“使用集合”部分](javascript-objects.md#working-with-collections)，深入了解如何在 Office 脚本框架中使用内置数组功能。

以下脚本将获取工作簿中的所有表格。 然后，它将确保显示标题、筛选按钮可见并且将表格样式设置为“TableStyleLight1”。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  /* Get table collection */
  const tables = workbook.getTables();
  /* Set table formatting properties */
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

#### <a name="adding-excel-objects-with-a-script"></a>使用脚本添加 Excel 对象

通过调用可在父对象上使用的相应 `add` 方法，可以以编程方式添加文档对象，如表格或图表。

> [!NOTE]
> 不要手动将对象添加到集合数组。 请在父对象上使用 `add` 方法。例如，使用 `Worksheet.addTable` 方法向 `Worksheet` 添加 `Table`。

以下脚本将在 Excel 工作簿中的第一个工作表上创建一个表格。 请注意，所创建的表格是通过 `addTable` 方法返回的。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in C3:G10.
    let table = sheet.addTable(
      "C3:G10",
       true /* True because the table has headers. */
    );
}
```

## <a name="removing-excel-objects-with-a-script"></a>使用脚本删除 Excel 对象

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

### <a name="further-reading-on-the-object-model"></a>进一步了解对象模型

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

## <a name="see-also"></a>另请参阅

- [在 Excel 网页版中录制、编辑和创建 Office 脚本](../tutorials/excel-tutorial.md)
- [在 Excel 网页版中使用 Office 脚本读取工作簿数据](../tutorials/excel-read-tutorial.md)
- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
- [在 Office 脚本中使用内置的 JavaScript 对象](javascript-objects.md)
