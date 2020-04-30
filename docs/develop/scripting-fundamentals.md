---
title: Excel 网页版中 Office 脚本的脚本基础
description: 在编写 Office 脚本之前需要了解的对象模型信息和其他基础知识。
ms.date: 04/24/2020
localization_priority: Priority
ms.openlocfilehash: 8449654e359f665677f3d416a8e28fa4d6930f26
ms.sourcegitcommit: 350bd2447f616fa87bb23ac826c7731fb813986b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/28/2020
ms.locfileid: "43919796"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Excel 网页版中 Office 脚本的脚本基础（预览）

本文将介绍 Office 脚本技术方面的知识。 你将了解 Excel 对象如何协同工作以及代码编辑器如何与工作簿同步。

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>对象模型

若要了解 Excel API，则必须了解工作簿的各个组件之间如何相互关联。

- 一个 **Workbook** 包含一个或多个 **Worksheet**。
- **Worksheet** 可通过 **Range** 对象访问单元格。
- **Range** 代表一组连续的单元格。
- **Range** 用于创建和放置 **Table**、**Chart** 和 **Shape** 以及其他数据可视化或组织对象。
- **Worksheet** 包含单个工作表中存在的那些数据对象的集合。
- **Workbook** 包含整个 **Workbook** 的某些数据对象（例如，**Table**）的集合。

### <a name="ranges"></a>Range

Range 是工作簿中的一组连续单元格。 脚本通常使用 A1 样式表示法（例如，对于列 **B** 和行 **3** 中单个的单元格 **B3** 或从列 **C** 至 **列F**和行 **2** 至 **行4** 的单元格 **C2:F4**）来定义范围。

Range 具有三个核心属性：`values`、`formulas` 和 `format`。 这些属性获取或设置单元格值、要计算的公式以及单元格的视觉对象格式设置。

#### <a name="range-sample"></a>Range 示例

以下示例显示了如何创建销售记录。 该脚本使用 `Range` 对象来设置值、公式和格式。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the active worksheet.
  let sheet = context.workbook.worksheets.getActiveWorksheet();

  // Create the headers and format them to stand out.
  let headers = [
    ["Product", "Quantity", "Unit Price", "Totals"]
  ];
  let headerRange = sheet.getRange("B2:E2");
  headerRange.values = headers;
  headerRange.format.fill.color = "#4472C4";
  headerRange.format.font.color = "white";

  // Create the product data rows.
  let productData = [
    ["Almonds", 6, 7.5],
    ["Coffee", 20, 34.5],
    ["Chocolate", 10, 9.56],
  ];
  let dataRange = sheet.getRange("B3:D5");
  dataRange.values = productData;

  // Create the formulas to total the amounts sold.
  let totalFormulas = [
    ["=C3 * D3"],
    ["=C4 * D4"],
    ["=C5 * D5"],
    ["=SUM(E3:E5)"]
  ];
  let totalRange = sheet.getRange("E3:E6");
  totalRange.formulas = totalFormulas;
  totalRange.format.font.bold = true;

  // Display the totals as US dollar amounts.
  totalRange.numberFormat = [["$0.00"]];
}
```

运行此脚本将在当前工作表中创建以下数据：

![显示值行、公式列和格式化标题的销售记录。](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Chart、Table 和其他数据对象

脚本可以在 Excel 中创建和设置数据结构和可视化效果。 Table 和 Chart 是最常用的两个对象，但是 API 支持数据透视表、形状和图像等。

#### <a name="creating-a-table"></a>创建表

通过使用数据填充范围创建表。 会将格式设置和表控件（如筛选器）自动应用到该范围。

以下脚本使用上一个示例中的范围创建一个表。

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

在工作表上使用之前的数据运行此脚本将创建下表：

![使用之前的销售记录制成的表。](../images/table-sample.png)

#### <a name="creating-a-chart"></a>创建图表

创建图表以直观显示某个范围内的数据。 脚本支持数十种图表类型，每种都可以根据需要进行自定义。

下面的脚本为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方，并将其设置为 100 像素。

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

在工作表上使用上一个表运行此脚本将创建以下图表：

![一个柱形图，显示上一个销售记录中三个项目的数量。](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>进一步了解对象模型

[Office 脚本 API 参考文档](/javascript/api/office-scripts/overview)是 Office 脚本中使用的对象的完整列表。 在这里，可以使用目录导航到想进一步了解的任何课程。 以下是几个经常查看的页面。

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Comment](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [区域](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Shape](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>`main` 函数

每个 Office 脚本都必须包含带有以下签名的 `main` 函数，其中包括 `Excel.RequestContext` 类型定义：

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

运行脚本时，`main` 函数中的代码将运行。 `main` 可以调用脚本中的其他函数，但是该函数中未包含的代码将不会运行。

## <a name="context"></a>上下文

`main` 函数接受名为 `context` 的 `Excel.RequestContext` 参数。 将 `context` 视作脚本和工作簿之间的桥梁。 脚本使用 `context` 对象访问工作簿，并使用该 `context` 来回发送数据。

`context` 对象是必需的，因为脚本和 Excel 在不同的进程和位置中运行。 该脚本将需要对云中的工作簿进行更改或从中查询数据。 `context` 对象管理以下事务。

## <a name="sync-and-load"></a>同步和加载

因为脚本和工作簿在不同的位置运行，所以两者之间的任何数据传输都需要时间。 为了提高脚本性能，对命令进行排队，直到脚本显式调用 `sync` 操作来同步脚本和工作簿。 脚本可以独立运行，直到需要执行以下任一操作：

- 从工作簿中读取数据（遵循返回 [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult) 的 `load` 操作或方法）。
- 将数据写入工作簿（通常是因为脚本已完成）。

下图显示了脚本和工作簿之间的示例控制流：

![该图显示了从脚本转到工作簿的读取和写入操作。](../images/load-sync.png)

### <a name="sync"></a>同步

每当脚本需要从工作簿读取数据或将数据写入工作簿时，请调用 `RequestContext.sync` 方法，如下所示：

```TypeScript
await context.sync();
```

> [!NOTE]
> 脚本结束时将隐式调用 `context.sync()`。

`sync` 操作完成后，工作簿将更新以反映脚本已指定的任何写入操作。 写入操作在 Excel 对象上设置任何属性（例如 `range.format.fill.color = "red"`），或调用更改属性的方法（例如 `range.format.autoFitColumns()`）。 `sync` 操作还从脚本请求的工作簿中读取任何值，方式是通过使用能返回 `ClientResult` 的 `load` 操作或方法（如下一节所述）。

将脚本与工作簿同步可能需要一些时间，具体取决于网络。 应尽量减少 `sync` 调用的次数，以帮助脚本快速运行。  

### <a name="load"></a>加载

脚本必须先从工作簿加载数据，然后才能读取数据。 但是，从整个工作簿中频繁加载数据将大大降低脚本的速度。 相反，通过 `load` 方法，脚本能够明确说明应从工作簿中检索哪些数据。

`load` 方法可用于每个 Excel 对象。 脚本必须先加载对象的属性，然后才能读取它们。 否则，将导致错误。

下面的示例使用 `Range` 对象显示 `load` 方法可用于加载数据的三种方式。

|意图 |示例命令 | 效果 |
|:--|:--|:--|
|加载一个属性 |`myRange.load("values");` | 加载单个属性，此例中为此范围内的二维值数组。 |
|加载多个属性 |`myRange.load("values, rowCount, columnCount");`| 从逗号分隔的列表中加载所有属性，此例中为值、行数和列数。 |
|加载所有内容 | `myRange.load();`|加载范围内的所有属性。 不建议采用此解决方案，因为获取不必要的数据会减慢脚本速度。 仅在测试脚本或需要对象的每个属性时，才应使用此方法。 |

脚本必须先调用 `context.sync()`，然后才能读取任何加载的值。

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

还可以在整个集合中加载属性。 每个集合对象都有一个 `items` 属性，该属性是一个包含该集合中的对象的数组。 使用 `items` 作为对 `load` 的层次调用 (`items\myProperty`) 的开始，将在其中的每个项目上加载指定的属性。 下面的示例在工作表的 `CommentCollection` 对象中的每个 `Comment` 对象上加载 `resolved` 属性。

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> 要了解有关在 Office 脚本中使用集合的更多信息，请参阅[在 Office 脚本中使用内置 JavaScript 对象的数组部分](javascript-objects.md#array)一文。

### <a name="clientresult"></a>ClientResult

从工作簿中返回信息的方法与`load`/`sync`范例的模式相同。 举个例子，`TableCollection.getCount`获取集合中的表的数量。 `getCount` 返回 `ClientResult<number>`，这意味着返回 `ClientResult` 中的 `value` 属性为 "数字"。 在调用 `context.sync()` 之前，脚本无法访问此值。 与加载属性很相似，直到 `sync` 调用，`value` 是本地 "空" 值。

以下脚本获取工作簿中的表的总数，并将该数目记录到控制台。

```TypeScript
async function main(context: Excel.RequestContext) {
  let tableCount = context.workbook.tables.getCount();

  // This sync call implicitly loads tableCount.value.
  // Any other ClientResult values are loaded too.
  await context.sync();

  // Trying to log the value before calling sync would throw an error.
  console.log(tableCount.value);
}
```

## <a name="see-also"></a>另请参阅

- [在 Excel 网页版中录制、编辑和创建 Office 脚本](../tutorials/excel-tutorial.md)
- [在 Excel 网页版中使用 Office 脚本读取工作簿数据](../tutorials/excel-read-tutorial.md)
- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
- [在 Office 脚本中使用内置的 JavaScript 对象](javascript-objects.md)
