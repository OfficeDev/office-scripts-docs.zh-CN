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
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="b0448-103">Excel 网页版中 Office 脚本的脚本基础（预览）</span><span class="sxs-lookup"><span data-stu-id="b0448-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="b0448-104">本文将介绍 Office 脚本技术方面的知识。</span><span class="sxs-lookup"><span data-stu-id="b0448-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="b0448-105">你将了解 Excel 对象如何协同工作以及代码编辑器如何与工作簿同步。</span><span class="sxs-lookup"><span data-stu-id="b0448-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="b0448-106">对象模型</span><span class="sxs-lookup"><span data-stu-id="b0448-106">Object model</span></span>

<span data-ttu-id="b0448-107">若要了解 Excel API，则必须了解工作簿的各个组件之间如何相互关联。</span><span class="sxs-lookup"><span data-stu-id="b0448-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="b0448-108">一个 **Workbook** 包含一个或多个 **Worksheet**。</span><span class="sxs-lookup"><span data-stu-id="b0448-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="b0448-109">**Worksheet** 可通过 **Range** 对象访问单元格。</span><span class="sxs-lookup"><span data-stu-id="b0448-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="b0448-110">**Range** 代表一组连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="b0448-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="b0448-111">**Range** 用于创建和放置 **Table**、**Chart** 和 **Shape** 以及其他数据可视化或组织对象。</span><span class="sxs-lookup"><span data-stu-id="b0448-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="b0448-112">**Worksheet** 包含单个工作表中存在的那些数据对象的集合。</span><span class="sxs-lookup"><span data-stu-id="b0448-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="b0448-113">**Workbook** 包含整个 **Workbook** 的某些数据对象（例如，**Table**）的集合。</span><span class="sxs-lookup"><span data-stu-id="b0448-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="b0448-114">Range</span><span class="sxs-lookup"><span data-stu-id="b0448-114">Ranges</span></span>

<span data-ttu-id="b0448-115">Range 是工作簿中的一组连续单元格。</span><span class="sxs-lookup"><span data-stu-id="b0448-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="b0448-116">脚本通常使用 A1 样式表示法（例如，对于列 **B** 和行 **3** 中单个的单元格 **B3** 或从列 **C** 至 **列F**和行 **2** 至 **行4** 的单元格 **C2:F4**）来定义范围。</span><span class="sxs-lookup"><span data-stu-id="b0448-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="b0448-117">Range 具有三个核心属性：`values`、`formulas` 和 `format`。</span><span class="sxs-lookup"><span data-stu-id="b0448-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="b0448-118">这些属性获取或设置单元格值、要计算的公式以及单元格的视觉对象格式设置。</span><span class="sxs-lookup"><span data-stu-id="b0448-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="b0448-119">Range 示例</span><span class="sxs-lookup"><span data-stu-id="b0448-119">Range sample</span></span>

<span data-ttu-id="b0448-120">以下示例显示了如何创建销售记录。</span><span class="sxs-lookup"><span data-stu-id="b0448-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="b0448-121">该脚本使用 `Range` 对象来设置值、公式和格式。</span><span class="sxs-lookup"><span data-stu-id="b0448-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="b0448-122">运行此脚本将在当前工作表中创建以下数据：</span><span class="sxs-lookup"><span data-stu-id="b0448-122">Running this script creates the following data in the current worksheet:</span></span>

![显示值行、公式列和格式化标题的销售记录。](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="b0448-124">Chart、Table 和其他数据对象</span><span class="sxs-lookup"><span data-stu-id="b0448-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="b0448-125">脚本可以在 Excel 中创建和设置数据结构和可视化效果。</span><span class="sxs-lookup"><span data-stu-id="b0448-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="b0448-126">Table 和 Chart 是最常用的两个对象，但是 API 支持数据透视表、形状和图像等。</span><span class="sxs-lookup"><span data-stu-id="b0448-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="b0448-127">创建表</span><span class="sxs-lookup"><span data-stu-id="b0448-127">Creating a table</span></span>

<span data-ttu-id="b0448-128">通过使用数据填充范围创建表。</span><span class="sxs-lookup"><span data-stu-id="b0448-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="b0448-129">会将格式设置和表控件（如筛选器）自动应用到该范围。</span><span class="sxs-lookup"><span data-stu-id="b0448-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="b0448-130">以下脚本使用上一个示例中的范围创建一个表。</span><span class="sxs-lookup"><span data-stu-id="b0448-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="b0448-131">在工作表上使用之前的数据运行此脚本将创建下表：</span><span class="sxs-lookup"><span data-stu-id="b0448-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![使用之前的销售记录制成的表。](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="b0448-133">创建图表</span><span class="sxs-lookup"><span data-stu-id="b0448-133">Creating a chart</span></span>

<span data-ttu-id="b0448-134">创建图表以直观显示某个范围内的数据。</span><span class="sxs-lookup"><span data-stu-id="b0448-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="b0448-135">脚本支持数十种图表类型，每种都可以根据需要进行自定义。</span><span class="sxs-lookup"><span data-stu-id="b0448-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="b0448-136">下面的脚本为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方，并将其设置为 100 像素。</span><span class="sxs-lookup"><span data-stu-id="b0448-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="b0448-137">在工作表上使用上一个表运行此脚本将创建以下图表：</span><span class="sxs-lookup"><span data-stu-id="b0448-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![一个柱形图，显示上一个销售记录中三个项目的数量。](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="b0448-139">进一步了解对象模型</span><span class="sxs-lookup"><span data-stu-id="b0448-139">Further reading on the object model</span></span>

<span data-ttu-id="b0448-140">[Office 脚本 API 参考文档](/javascript/api/office-scripts/overview)是 Office 脚本中使用的对象的完整列表。</span><span class="sxs-lookup"><span data-stu-id="b0448-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="b0448-141">在这里，可以使用目录导航到想进一步了解的任何课程。</span><span class="sxs-lookup"><span data-stu-id="b0448-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="b0448-142">以下是几个经常查看的页面。</span><span class="sxs-lookup"><span data-stu-id="b0448-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="b0448-143">Chart</span><span class="sxs-lookup"><span data-stu-id="b0448-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="b0448-144">Comment</span><span class="sxs-lookup"><span data-stu-id="b0448-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="b0448-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="b0448-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="b0448-146">区域</span><span class="sxs-lookup"><span data-stu-id="b0448-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="b0448-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="b0448-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="b0448-148">Shape</span><span class="sxs-lookup"><span data-stu-id="b0448-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="b0448-149">Table</span><span class="sxs-lookup"><span data-stu-id="b0448-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="b0448-150">Workbook</span><span class="sxs-lookup"><span data-stu-id="b0448-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="b0448-151">Worksheet</span><span class="sxs-lookup"><span data-stu-id="b0448-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="b0448-152">`main` 函数</span><span class="sxs-lookup"><span data-stu-id="b0448-152">`main` function</span></span>

<span data-ttu-id="b0448-153">每个 Office 脚本都必须包含带有以下签名的 `main` 函数，其中包括 `Excel.RequestContext` 类型定义：</span><span class="sxs-lookup"><span data-stu-id="b0448-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="b0448-154">运行脚本时，`main` 函数中的代码将运行。</span><span class="sxs-lookup"><span data-stu-id="b0448-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="b0448-155">`main` 可以调用脚本中的其他函数，但是该函数中未包含的代码将不会运行。</span><span class="sxs-lookup"><span data-stu-id="b0448-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="b0448-156">上下文</span><span class="sxs-lookup"><span data-stu-id="b0448-156">Context</span></span>

<span data-ttu-id="b0448-157">`main` 函数接受名为 `context` 的 `Excel.RequestContext` 参数。</span><span class="sxs-lookup"><span data-stu-id="b0448-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="b0448-158">将 `context` 视作脚本和工作簿之间的桥梁。</span><span class="sxs-lookup"><span data-stu-id="b0448-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="b0448-159">脚本使用 `context` 对象访问工作簿，并使用该 `context` 来回发送数据。</span><span class="sxs-lookup"><span data-stu-id="b0448-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="b0448-160">`context` 对象是必需的，因为脚本和 Excel 在不同的进程和位置中运行。</span><span class="sxs-lookup"><span data-stu-id="b0448-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="b0448-161">该脚本将需要对云中的工作簿进行更改或从中查询数据。</span><span class="sxs-lookup"><span data-stu-id="b0448-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="b0448-162">`context` 对象管理以下事务。</span><span class="sxs-lookup"><span data-stu-id="b0448-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="b0448-163">同步和加载</span><span class="sxs-lookup"><span data-stu-id="b0448-163">Sync and Load</span></span>

<span data-ttu-id="b0448-164">因为脚本和工作簿在不同的位置运行，所以两者之间的任何数据传输都需要时间。</span><span class="sxs-lookup"><span data-stu-id="b0448-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="b0448-165">为了提高脚本性能，对命令进行排队，直到脚本显式调用 `sync` 操作来同步脚本和工作簿。</span><span class="sxs-lookup"><span data-stu-id="b0448-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="b0448-166">脚本可以独立运行，直到需要执行以下任一操作：</span><span class="sxs-lookup"><span data-stu-id="b0448-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="b0448-167">从工作簿中读取数据（遵循返回 [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult) 的 `load` 操作或方法）。</span><span class="sxs-lookup"><span data-stu-id="b0448-167">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult)).</span></span>
- <span data-ttu-id="b0448-168">将数据写入工作簿（通常是因为脚本已完成）。</span><span class="sxs-lookup"><span data-stu-id="b0448-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="b0448-169">下图显示了脚本和工作簿之间的示例控制流：</span><span class="sxs-lookup"><span data-stu-id="b0448-169">The following image shows an example control flow between the script and workbook:</span></span>

![该图显示了从脚本转到工作簿的读取和写入操作。](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="b0448-171">同步</span><span class="sxs-lookup"><span data-stu-id="b0448-171">Sync</span></span>

<span data-ttu-id="b0448-172">每当脚本需要从工作簿读取数据或将数据写入工作簿时，请调用 `RequestContext.sync` 方法，如下所示：</span><span class="sxs-lookup"><span data-stu-id="b0448-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="b0448-173">脚本结束时将隐式调用 `context.sync()`。</span><span class="sxs-lookup"><span data-stu-id="b0448-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="b0448-174">`sync` 操作完成后，工作簿将更新以反映脚本已指定的任何写入操作。</span><span class="sxs-lookup"><span data-stu-id="b0448-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="b0448-175">写入操作在 Excel 对象上设置任何属性（例如 `range.format.fill.color = "red"`），或调用更改属性的方法（例如 `range.format.autoFitColumns()`）。</span><span class="sxs-lookup"><span data-stu-id="b0448-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="b0448-176">`sync` 操作还从脚本请求的工作簿中读取任何值，方式是通过使用能返回 `ClientResult` 的 `load` 操作或方法（如下一节所述）。</span><span class="sxs-lookup"><span data-stu-id="b0448-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="b0448-177">将脚本与工作簿同步可能需要一些时间，具体取决于网络。</span><span class="sxs-lookup"><span data-stu-id="b0448-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="b0448-178">应尽量减少 `sync` 调用的次数，以帮助脚本快速运行。</span><span class="sxs-lookup"><span data-stu-id="b0448-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="b0448-179">加载</span><span class="sxs-lookup"><span data-stu-id="b0448-179">Load</span></span>

<span data-ttu-id="b0448-180">脚本必须先从工作簿加载数据，然后才能读取数据。</span><span class="sxs-lookup"><span data-stu-id="b0448-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="b0448-181">但是，从整个工作簿中频繁加载数据将大大降低脚本的速度。</span><span class="sxs-lookup"><span data-stu-id="b0448-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="b0448-182">相反，通过 `load` 方法，脚本能够明确说明应从工作簿中检索哪些数据。</span><span class="sxs-lookup"><span data-stu-id="b0448-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="b0448-183">`load` 方法可用于每个 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="b0448-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="b0448-184">脚本必须先加载对象的属性，然后才能读取它们。</span><span class="sxs-lookup"><span data-stu-id="b0448-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="b0448-185">否则，将导致错误。</span><span class="sxs-lookup"><span data-stu-id="b0448-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="b0448-186">下面的示例使用 `Range` 对象显示 `load` 方法可用于加载数据的三种方式。</span><span class="sxs-lookup"><span data-stu-id="b0448-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="b0448-187">意图</span><span class="sxs-lookup"><span data-stu-id="b0448-187">Intent</span></span> |<span data-ttu-id="b0448-188">示例命令</span><span class="sxs-lookup"><span data-stu-id="b0448-188">Example Command</span></span> | <span data-ttu-id="b0448-189">效果</span><span class="sxs-lookup"><span data-stu-id="b0448-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="b0448-190">加载一个属性</span><span class="sxs-lookup"><span data-stu-id="b0448-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="b0448-191">加载单个属性，此例中为此范围内的二维值数组。</span><span class="sxs-lookup"><span data-stu-id="b0448-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="b0448-192">加载多个属性</span><span class="sxs-lookup"><span data-stu-id="b0448-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="b0448-193">从逗号分隔的列表中加载所有属性，此例中为值、行数和列数。</span><span class="sxs-lookup"><span data-stu-id="b0448-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="b0448-194">加载所有内容</span><span class="sxs-lookup"><span data-stu-id="b0448-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="b0448-195">加载范围内的所有属性。</span><span class="sxs-lookup"><span data-stu-id="b0448-195">Loads all the properties on the range.</span></span> <span data-ttu-id="b0448-196">不建议采用此解决方案，因为获取不必要的数据会减慢脚本速度。</span><span class="sxs-lookup"><span data-stu-id="b0448-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="b0448-197">仅在测试脚本或需要对象的每个属性时，才应使用此方法。</span><span class="sxs-lookup"><span data-stu-id="b0448-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="b0448-198">脚本必须先调用 `context.sync()`，然后才能读取任何加载的值。</span><span class="sxs-lookup"><span data-stu-id="b0448-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="b0448-199">还可以在整个集合中加载属性。</span><span class="sxs-lookup"><span data-stu-id="b0448-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="b0448-200">每个集合对象都有一个 `items` 属性，该属性是一个包含该集合中的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="b0448-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="b0448-201">使用 `items` 作为对 `load` 的层次调用 (`items\myProperty`) 的开始，将在其中的每个项目上加载指定的属性。</span><span class="sxs-lookup"><span data-stu-id="b0448-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="b0448-202">下面的示例在工作表的 `CommentCollection` 对象中的每个 `Comment` 对象上加载 `resolved` 属性。</span><span class="sxs-lookup"><span data-stu-id="b0448-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="b0448-203">要了解有关在 Office 脚本中使用集合的更多信息，请参阅[在 Office 脚本中使用内置 JavaScript 对象的数组部分](javascript-objects.md#array)一文。</span><span class="sxs-lookup"><span data-stu-id="b0448-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

### <a name="clientresult"></a><span data-ttu-id="b0448-204">ClientResult</span><span class="sxs-lookup"><span data-stu-id="b0448-204">ClientResult</span></span>

<span data-ttu-id="b0448-205">从工作簿中返回信息的方法与`load`/`sync`范例的模式相同。</span><span class="sxs-lookup"><span data-stu-id="b0448-205">Methods that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="b0448-206">举个例子，`TableCollection.getCount`获取集合中的表的数量。</span><span class="sxs-lookup"><span data-stu-id="b0448-206">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="b0448-207">`getCount` 返回 `ClientResult<number>`，这意味着返回 `ClientResult` 中的 `value` 属性为 "数字"。</span><span class="sxs-lookup"><span data-stu-id="b0448-207">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the return `ClientResult` is a number.</span></span> <span data-ttu-id="b0448-208">在调用 `context.sync()` 之前，脚本无法访问此值。</span><span class="sxs-lookup"><span data-stu-id="b0448-208">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="b0448-209">与加载属性很相似，直到 `sync` 调用，`value` 是本地 "空" 值。</span><span class="sxs-lookup"><span data-stu-id="b0448-209">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="b0448-210">以下脚本获取工作簿中的表的总数，并将该数目记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="b0448-210">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="b0448-211">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b0448-211">See also</span></span>

- [<span data-ttu-id="b0448-212">在 Excel 网页版中录制、编辑和创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="b0448-212">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="b0448-213">在 Excel 网页版中使用 Office 脚本读取工作簿数据</span><span class="sxs-lookup"><span data-stu-id="b0448-213">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="b0448-214">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="b0448-214">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="b0448-215">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="b0448-215">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
