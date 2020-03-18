---
title: Web 上的 Excel 中 Office 脚本的脚本基础
description: 编写 Office 脚本前要了解的对象模型信息和其他基础知识。
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700112"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Web 上的 Excel 中的 Office 脚本的脚本基础（预览）

本文将向您介绍 Office 脚本的技术方面。 您将了解 Excel 对象如何协同工作，以及代码编辑器如何与工作簿同步。

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>对象模型

若要了解 Excel Api，您必须了解工作簿的组件如何彼此关联。

- 一个**工作簿**包含一个或多个**工作表**。
- **工作表**通过**Range**对象提供对单元格的访问。
- **区域**代表一组连续的单元格。
- **区域**用于创建和放置**表**、**图表**、**形状**以及其他数据可视化或组织对象。
- **工作表**包含各个工作表中存在的这些数据对象的集合。
- **工作簿**包含用于整个**工作簿**的一些数据对象（如**表**）的集合。

### <a name="ranges"></a>Ranges

区域是工作簿中的一组连续单元格。 脚本通常使用 A1 样式表示法（**例如，在**第**B**行中的单个单元格和第**3**列或 C2 中的单元格 **： F4**行**C**至**F**中的单元格以及**2**到**4**列的单元格）来定义区域。

范围具有三个核心属性`values`： `formulas`、和`format`。 这些属性获取或设置单元格的值、要计算的公式以及单元格的可视格式。

#### <a name="range-sample"></a>范围示例

下面的示例展示了如何创建销售记录。 此脚本使用`Range`对象来设置值、公式和格式。

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

运行此脚本会在当前工作表中创建以下数据：

![显示值行、公式列和格式化标头的销售记录。](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>图表、表和其他数据对象

脚本可以创建和操作 Excel 中的数据结构和可视化效果。 表和图表是两个更常用的对象，但 Api 支持数据透视表、形状、图像等。

#### <a name="creating-a-table"></a>创建表

使用数据填充区域创建表格。 格式和表控件（如筛选器）将自动应用于区域。

下面的脚本使用上一示例中的区域创建一个表。

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

使用以前的数据在工作表上运行此脚本，将创建下表：

![从上一个销售记录中创建的表。](../images/table-sample.png)

#### <a name="creating-a-chart"></a>创建图表

创建图表以可视化区域中的数据。 脚本允许几十个图表种类，可以根据自己的需要对每个图表进行自定义。

下面的脚本为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方100像素处。

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

在具有上表的工作表上运行此脚本将创建以下图表：

![显示来自前一条销售记录的三个项目数量的柱形图。](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>进一步阅读对象模型

[Office 脚本 API 参考文档](/javascript/api/office-scripts/overview)是 office 脚本中使用的对象的完整列表。 在这里，您可以使用目录导航到您想要了解详细信息的任何类。 以下是几个经常查看的页面。

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Comment](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Shape](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>`main`function

每个 Office 脚本必须包含`main`具有以下签名的函数，其中包括`Excel.RequestContext`类型定义：

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

`main`函数中的代码在脚本运行时运行。 `main`可以在脚本中调用其他函数，但不会运行不包含在函数中的代码。

## <a name="context"></a>Context

`main`函数接受名为`Excel.RequestContext` `context`的参数。 将脚本`context`与工作簿之间的桥梁视为桥梁。 您的脚本使用`context`对象访问工作簿，并使用`context`它来来回发送数据。

该`context`对象是必需的，因为脚本和 Excel 运行在不同的进程和位置。 脚本将需要对云中的工作簿进行更改或查询数据。 `context`对象管理这些事务。

## <a name="sync-and-load"></a>同步和加载

由于您的脚本和工作簿运行在不同的位置，因此两者之间的任何数据传输都会占用时间。 为了提高脚本性能，命令在脚本显式调用`sync`操作以同步脚本和工作簿之前将排队。 您的脚本可以独立运行，直到需要执行以下操作之一：

- 从工作簿中读取数据（遵循`load`操作）。
- 将数据写入工作簿（通常是因为脚本已完成）。

下图显示了脚本和工作簿之间的控制流示例：

![显示从脚本转到工作簿的读取和写入操作的图表。](../images/load-sync.png)

### <a name="sync"></a>同步

只要脚本需要从工作簿中读取数据或将数据写入工作簿，请`RequestContext.sync`调用如下所示的方法：

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()`在脚本结束时隐式调用。

`sync`操作完成后，工作簿将进行更新以反映该脚本指定的任何写操作。 写操作是设置 Excel 对象的任何属性（例如`range.format.fill.color = "red"`）或调用更改属性（如`range.format.autoFitColumns()`）的方法。 该`sync`操作还会从工作簿中读取使用`load`操作请求的脚本的任何值（下一节将对此进行了讨论）。

将脚本与工作簿同步可能需要一些时间，具体取决于您的网络。 应尽量减少`sync`调用次数，以帮助脚本运行速度更快。  

### <a name="load"></a>负载

脚本必须先从工作簿加载数据，然后才能阅读。 但是，从整个工作簿中频繁加载数据将极大地降低脚本速度。 相反，此`load`方法允许您的脚本状态专门从工作簿中检索哪些数据。

该`load`方法可用于每个 Excel 对象。 您的脚本必须先加载对象的属性，然后它才能阅读。 如果不这样做，则会导致错误。

下面的示例使用`Range`对象显示`load`方法可用于加载数据的三种方法。

|Intent |示例命令 | 效果 |
|:--|:--|:--|
|加载一个属性 |`myRange.load("values");` | 加载单个属性，在此示例中，此范围中值的二维数组。 |
|加载多个属性 |`myRange.load("values, rowCount, columnCount");`| 从逗号分隔的列表中加载所有属性，本示例中的值、行数和列数。 |
|加载所有内容 | `myRange.load();`|加载区域中的所有属性。 这不是建议的解决方案，因为它将通过获取不必要的数据来降低脚本的速度。 只应在测试脚本时使用此属性，或者如果需要从对象中的每个属性。 |

在读取任何加载`context.sync()`的值之前，必须先调用您的脚本。

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

您还可以在整个集合中加载属性。 每个集合对象都`items`具有一个属性，该属性是包含该集合中的对象的数组。 使用`items`作为分层调用的开始（`items\myProperty`）， `load`在每个项目上加载指定的属性。 下面的示例将在`resolved`工作表的`Comment` `CommentCollection`对象中的每个对象上加载属性。

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> 若要了解有关在 Office 脚本中使用集合的详细信息，请参阅在[Office 脚本文章中使用内置 JavaScript 对象的数组部分](javascript-objects.md#array)。

## <a name="see-also"></a>另请参阅

- [在 Excel 网页上记录、编辑和创建 Office 脚本](../tutorials/excel-tutorial.md)
- [在 Excel 网页上使用 Office 脚本读取工作簿数据](../tutorials/excel-read-tutorial.md)
- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
- [在 Office 脚本中使用内置 JavaScript 对象](javascript-objects.md)
