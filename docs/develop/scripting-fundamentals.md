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
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="0297c-103">Web 上的 Excel 中的 Office 脚本的脚本基础（预览）</span><span class="sxs-lookup"><span data-stu-id="0297c-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="0297c-104">本文将向您介绍 Office 脚本的技术方面。</span><span class="sxs-lookup"><span data-stu-id="0297c-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="0297c-105">您将了解 Excel 对象如何协同工作，以及代码编辑器如何与工作簿同步。</span><span class="sxs-lookup"><span data-stu-id="0297c-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="0297c-106">对象模型</span><span class="sxs-lookup"><span data-stu-id="0297c-106">Object model</span></span>

<span data-ttu-id="0297c-107">若要了解 Excel Api，您必须了解工作簿的组件如何彼此关联。</span><span class="sxs-lookup"><span data-stu-id="0297c-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="0297c-108">一个**工作簿**包含一个或多个**工作表**。</span><span class="sxs-lookup"><span data-stu-id="0297c-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="0297c-109">**工作表**通过**Range**对象提供对单元格的访问。</span><span class="sxs-lookup"><span data-stu-id="0297c-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="0297c-110">**区域**代表一组连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="0297c-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="0297c-111">**区域**用于创建和放置**表**、**图表**、**形状**以及其他数据可视化或组织对象。</span><span class="sxs-lookup"><span data-stu-id="0297c-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="0297c-112">**工作表**包含各个工作表中存在的这些数据对象的集合。</span><span class="sxs-lookup"><span data-stu-id="0297c-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="0297c-113">**工作簿**包含用于整个**工作簿**的一些数据对象（如**表**）的集合。</span><span class="sxs-lookup"><span data-stu-id="0297c-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="0297c-114">Ranges</span><span class="sxs-lookup"><span data-stu-id="0297c-114">Ranges</span></span>

<span data-ttu-id="0297c-115">区域是工作簿中的一组连续单元格。</span><span class="sxs-lookup"><span data-stu-id="0297c-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="0297c-116">脚本通常使用 A1 样式表示法（**例如，在**第**B**行中的单个单元格和第**3**列或 C2 中的单元格 **： F4**行**C**至**F**中的单元格以及**2**到**4**列的单元格）来定义区域。</span><span class="sxs-lookup"><span data-stu-id="0297c-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in row **B** and column **3** or **C2:F4** for the cells from rows **C** through **F** and columns **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="0297c-117">范围具有三个核心属性`values`： `formulas`、和`format`。</span><span class="sxs-lookup"><span data-stu-id="0297c-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="0297c-118">这些属性获取或设置单元格的值、要计算的公式以及单元格的可视格式。</span><span class="sxs-lookup"><span data-stu-id="0297c-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="0297c-119">范围示例</span><span class="sxs-lookup"><span data-stu-id="0297c-119">Range sample</span></span>

<span data-ttu-id="0297c-120">下面的示例展示了如何创建销售记录。</span><span class="sxs-lookup"><span data-stu-id="0297c-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="0297c-121">此脚本使用`Range`对象来设置值、公式和格式。</span><span class="sxs-lookup"><span data-stu-id="0297c-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="0297c-122">运行此脚本会在当前工作表中创建以下数据：</span><span class="sxs-lookup"><span data-stu-id="0297c-122">Running this script creates the following data in the current worksheet:</span></span>

![显示值行、公式列和格式化标头的销售记录。](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="0297c-124">图表、表和其他数据对象</span><span class="sxs-lookup"><span data-stu-id="0297c-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="0297c-125">脚本可以创建和操作 Excel 中的数据结构和可视化效果。</span><span class="sxs-lookup"><span data-stu-id="0297c-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="0297c-126">表和图表是两个更常用的对象，但 Api 支持数据透视表、形状、图像等。</span><span class="sxs-lookup"><span data-stu-id="0297c-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="0297c-127">创建表</span><span class="sxs-lookup"><span data-stu-id="0297c-127">Creating a table</span></span>

<span data-ttu-id="0297c-128">使用数据填充区域创建表格。</span><span class="sxs-lookup"><span data-stu-id="0297c-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="0297c-129">格式和表控件（如筛选器）将自动应用于区域。</span><span class="sxs-lookup"><span data-stu-id="0297c-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="0297c-130">下面的脚本使用上一示例中的区域创建一个表。</span><span class="sxs-lookup"><span data-stu-id="0297c-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="0297c-131">使用以前的数据在工作表上运行此脚本，将创建下表：</span><span class="sxs-lookup"><span data-stu-id="0297c-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![从上一个销售记录中创建的表。](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="0297c-133">创建图表</span><span class="sxs-lookup"><span data-stu-id="0297c-133">Creating a chart</span></span>

<span data-ttu-id="0297c-134">创建图表以可视化区域中的数据。</span><span class="sxs-lookup"><span data-stu-id="0297c-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="0297c-135">脚本允许几十个图表种类，可以根据自己的需要对每个图表进行自定义。</span><span class="sxs-lookup"><span data-stu-id="0297c-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="0297c-136">下面的脚本为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方100像素处。</span><span class="sxs-lookup"><span data-stu-id="0297c-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="0297c-137">在具有上表的工作表上运行此脚本将创建以下图表：</span><span class="sxs-lookup"><span data-stu-id="0297c-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![显示来自前一条销售记录的三个项目数量的柱形图。](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="0297c-139">进一步阅读对象模型</span><span class="sxs-lookup"><span data-stu-id="0297c-139">Further reading on the object model</span></span>

<span data-ttu-id="0297c-140">[Office 脚本 API 参考文档](/javascript/api/office-scripts/overview)是 office 脚本中使用的对象的完整列表。</span><span class="sxs-lookup"><span data-stu-id="0297c-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="0297c-141">在这里，您可以使用目录导航到您想要了解详细信息的任何类。</span><span class="sxs-lookup"><span data-stu-id="0297c-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="0297c-142">以下是几个经常查看的页面。</span><span class="sxs-lookup"><span data-stu-id="0297c-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="0297c-143">Chart</span><span class="sxs-lookup"><span data-stu-id="0297c-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="0297c-144">Comment</span><span class="sxs-lookup"><span data-stu-id="0297c-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="0297c-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="0297c-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="0297c-146">Range</span><span class="sxs-lookup"><span data-stu-id="0297c-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="0297c-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="0297c-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="0297c-148">Shape</span><span class="sxs-lookup"><span data-stu-id="0297c-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="0297c-149">Table</span><span class="sxs-lookup"><span data-stu-id="0297c-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="0297c-150">Workbook</span><span class="sxs-lookup"><span data-stu-id="0297c-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="0297c-151">Worksheet</span><span class="sxs-lookup"><span data-stu-id="0297c-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="0297c-152">`main`function</span><span class="sxs-lookup"><span data-stu-id="0297c-152">`main` function</span></span>

<span data-ttu-id="0297c-153">每个 Office 脚本必须包含`main`具有以下签名的函数，其中包括`Excel.RequestContext`类型定义：</span><span class="sxs-lookup"><span data-stu-id="0297c-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="0297c-154">`main`函数中的代码在脚本运行时运行。</span><span class="sxs-lookup"><span data-stu-id="0297c-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="0297c-155">`main`可以在脚本中调用其他函数，但不会运行不包含在函数中的代码。</span><span class="sxs-lookup"><span data-stu-id="0297c-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="0297c-156">Context</span><span class="sxs-lookup"><span data-stu-id="0297c-156">Context</span></span>

<span data-ttu-id="0297c-157">`main`函数接受名为`Excel.RequestContext` `context`的参数。</span><span class="sxs-lookup"><span data-stu-id="0297c-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="0297c-158">将脚本`context`与工作簿之间的桥梁视为桥梁。</span><span class="sxs-lookup"><span data-stu-id="0297c-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="0297c-159">您的脚本使用`context`对象访问工作簿，并使用`context`它来来回发送数据。</span><span class="sxs-lookup"><span data-stu-id="0297c-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="0297c-160">该`context`对象是必需的，因为脚本和 Excel 运行在不同的进程和位置。</span><span class="sxs-lookup"><span data-stu-id="0297c-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="0297c-161">脚本将需要对云中的工作簿进行更改或查询数据。</span><span class="sxs-lookup"><span data-stu-id="0297c-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="0297c-162">`context`对象管理这些事务。</span><span class="sxs-lookup"><span data-stu-id="0297c-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="0297c-163">同步和加载</span><span class="sxs-lookup"><span data-stu-id="0297c-163">Sync and Load</span></span>

<span data-ttu-id="0297c-164">由于您的脚本和工作簿运行在不同的位置，因此两者之间的任何数据传输都会占用时间。</span><span class="sxs-lookup"><span data-stu-id="0297c-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="0297c-165">为了提高脚本性能，命令在脚本显式调用`sync`操作以同步脚本和工作簿之前将排队。</span><span class="sxs-lookup"><span data-stu-id="0297c-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="0297c-166">您的脚本可以独立运行，直到需要执行以下操作之一：</span><span class="sxs-lookup"><span data-stu-id="0297c-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="0297c-167">从工作簿中读取数据（遵循`load`操作）。</span><span class="sxs-lookup"><span data-stu-id="0297c-167">Read data from the workbook (following a `load` operation).</span></span>
- <span data-ttu-id="0297c-168">将数据写入工作簿（通常是因为脚本已完成）。</span><span class="sxs-lookup"><span data-stu-id="0297c-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="0297c-169">下图显示了脚本和工作簿之间的控制流示例：</span><span class="sxs-lookup"><span data-stu-id="0297c-169">The following image shows an example control flow between the script and workbook:</span></span>

![显示从脚本转到工作簿的读取和写入操作的图表。](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="0297c-171">同步</span><span class="sxs-lookup"><span data-stu-id="0297c-171">Sync</span></span>

<span data-ttu-id="0297c-172">只要脚本需要从工作簿中读取数据或将数据写入工作簿，请`RequestContext.sync`调用如下所示的方法：</span><span class="sxs-lookup"><span data-stu-id="0297c-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="0297c-173">`context.sync()`在脚本结束时隐式调用。</span><span class="sxs-lookup"><span data-stu-id="0297c-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="0297c-174">`sync`操作完成后，工作簿将进行更新以反映该脚本指定的任何写操作。</span><span class="sxs-lookup"><span data-stu-id="0297c-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="0297c-175">写操作是设置 Excel 对象的任何属性（例如`range.format.fill.color = "red"`）或调用更改属性（如`range.format.autoFitColumns()`）的方法。</span><span class="sxs-lookup"><span data-stu-id="0297c-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="0297c-176">该`sync`操作还会从工作簿中读取使用`load`操作请求的脚本的任何值（下一节将对此进行了讨论）。</span><span class="sxs-lookup"><span data-stu-id="0297c-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation (as discussed in the next section).</span></span>

<span data-ttu-id="0297c-177">将脚本与工作簿同步可能需要一些时间，具体取决于您的网络。</span><span class="sxs-lookup"><span data-stu-id="0297c-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="0297c-178">应尽量减少`sync`调用次数，以帮助脚本运行速度更快。</span><span class="sxs-lookup"><span data-stu-id="0297c-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="0297c-179">负载</span><span class="sxs-lookup"><span data-stu-id="0297c-179">Load</span></span>

<span data-ttu-id="0297c-180">脚本必须先从工作簿加载数据，然后才能阅读。</span><span class="sxs-lookup"><span data-stu-id="0297c-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="0297c-181">但是，从整个工作簿中频繁加载数据将极大地降低脚本速度。</span><span class="sxs-lookup"><span data-stu-id="0297c-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="0297c-182">相反，此`load`方法允许您的脚本状态专门从工作簿中检索哪些数据。</span><span class="sxs-lookup"><span data-stu-id="0297c-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="0297c-183">该`load`方法可用于每个 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="0297c-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="0297c-184">您的脚本必须先加载对象的属性，然后它才能阅读。</span><span class="sxs-lookup"><span data-stu-id="0297c-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="0297c-185">如果不这样做，则会导致错误。</span><span class="sxs-lookup"><span data-stu-id="0297c-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="0297c-186">下面的示例使用`Range`对象显示`load`方法可用于加载数据的三种方法。</span><span class="sxs-lookup"><span data-stu-id="0297c-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="0297c-187">Intent</span><span class="sxs-lookup"><span data-stu-id="0297c-187">Intent</span></span> |<span data-ttu-id="0297c-188">示例命令</span><span class="sxs-lookup"><span data-stu-id="0297c-188">Example Command</span></span> | <span data-ttu-id="0297c-189">效果</span><span class="sxs-lookup"><span data-stu-id="0297c-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="0297c-190">加载一个属性</span><span class="sxs-lookup"><span data-stu-id="0297c-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="0297c-191">加载单个属性，在此示例中，此范围中值的二维数组。</span><span class="sxs-lookup"><span data-stu-id="0297c-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="0297c-192">加载多个属性</span><span class="sxs-lookup"><span data-stu-id="0297c-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="0297c-193">从逗号分隔的列表中加载所有属性，本示例中的值、行数和列数。</span><span class="sxs-lookup"><span data-stu-id="0297c-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="0297c-194">加载所有内容</span><span class="sxs-lookup"><span data-stu-id="0297c-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="0297c-195">加载区域中的所有属性。</span><span class="sxs-lookup"><span data-stu-id="0297c-195">Loads all the properties on the range.</span></span> <span data-ttu-id="0297c-196">这不是建议的解决方案，因为它将通过获取不必要的数据来降低脚本的速度。</span><span class="sxs-lookup"><span data-stu-id="0297c-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="0297c-197">只应在测试脚本时使用此属性，或者如果需要从对象中的每个属性。</span><span class="sxs-lookup"><span data-stu-id="0297c-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="0297c-198">在读取任何加载`context.sync()`的值之前，必须先调用您的脚本。</span><span class="sxs-lookup"><span data-stu-id="0297c-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="0297c-199">您还可以在整个集合中加载属性。</span><span class="sxs-lookup"><span data-stu-id="0297c-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="0297c-200">每个集合对象都`items`具有一个属性，该属性是包含该集合中的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="0297c-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="0297c-201">使用`items`作为分层调用的开始（`items\myProperty`）， `load`在每个项目上加载指定的属性。</span><span class="sxs-lookup"><span data-stu-id="0297c-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="0297c-202">下面的示例将在`resolved`工作表的`Comment` `CommentCollection`对象中的每个对象上加载属性。</span><span class="sxs-lookup"><span data-stu-id="0297c-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="0297c-203">若要了解有关在 Office 脚本中使用集合的详细信息，请参阅在[Office 脚本文章中使用内置 JavaScript 对象的数组部分](javascript-objects.md#array)。</span><span class="sxs-lookup"><span data-stu-id="0297c-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="0297c-204">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0297c-204">See also</span></span>

- [<span data-ttu-id="0297c-205">在 Excel 网页上记录、编辑和创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="0297c-205">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="0297c-206">在 Excel 网页上使用 Office 脚本读取工作簿数据</span><span class="sxs-lookup"><span data-stu-id="0297c-206">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="0297c-207">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="0297c-207">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="0297c-208">在 Office 脚本中使用内置 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="0297c-208">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
