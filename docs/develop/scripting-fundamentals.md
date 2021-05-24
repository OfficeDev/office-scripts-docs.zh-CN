---
title: Excel 网页版中 Office 脚本的脚本基础
description: 在编写 Office 脚本之前需要了解的对象模型信息和其他基础知识。
ms.date: 05/10/2021
localization_priority: Priority
ms.openlocfilehash: d930c9ee36933cb0458de8cce4f1d1adc7b6a001
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545096"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="cefc5-103">Excel 网页版中 Office 脚本的脚本基础（预览）</span><span class="sxs-lookup"><span data-stu-id="cefc5-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="cefc5-104">本文将介绍 Office 脚本技术方面的知识。</span><span class="sxs-lookup"><span data-stu-id="cefc5-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="cefc5-105">你将了解 Excel 对象如何协同工作以及代码编辑器如何与工作簿同步。</span><span class="sxs-lookup"><span data-stu-id="cefc5-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="typescript-the-language-of-office-scripts"></a><span data-ttu-id="cefc5-106">TypeScript：Office 脚本的语言</span><span class="sxs-lookup"><span data-stu-id="cefc5-106">TypeScript: The language of Office Scripts</span></span>

<span data-ttu-id="cefc5-107">Office 脚本以 [TypeScript](https://www.typescriptlang.org/docs/home.html) 编写，它是 [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) 的一个超集。</span><span class="sxs-lookup"><span data-stu-id="cefc5-107">Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span></span> <span data-ttu-id="cefc5-108">如果熟悉 JavaScript，你的知识将会延续下去，因为两种语言的大部分代码是相同的。</span><span class="sxs-lookup"><span data-stu-id="cefc5-108">If you're familiar with JavaScript, your knowledge will carry over because much of the code is the same in both languages.</span></span> <span data-ttu-id="cefc5-109">在开始 Office 脚本编码之旅之前，我们建议你先掌握一些初级编程知识。</span><span class="sxs-lookup"><span data-stu-id="cefc5-109">We recommend you have some beginner-level programming knowledge before starting your Office Scripts coding journey.</span></span> <span data-ttu-id="cefc5-110">以下资源可以帮助理解 Office 脚本的编码方面。</span><span class="sxs-lookup"><span data-stu-id="cefc5-110">The following resources can help you understand the coding side of Office Scripts.</span></span>

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="cefc5-111">`main` 函数：脚本的起点</span><span class="sxs-lookup"><span data-stu-id="cefc5-111">`main` function: The script's starting point</span></span>

<span data-ttu-id="cefc5-112">每个脚本都必须包含一个 `main` 函数，并以 `ExcelScript.Workbook` 类型作为第一个参数。</span><span class="sxs-lookup"><span data-stu-id="cefc5-112">Each script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="cefc5-113">函数运行时，Excel 应用程序通过提供工作簿作为第一个参数来调用 `main` 函数。</span><span class="sxs-lookup"><span data-stu-id="cefc5-113">When the function runs, the Excel application invokes the `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="cefc5-114">`ExcelScript.Workbook` 应始终是第一个参数。</span><span class="sxs-lookup"><span data-stu-id="cefc5-114">An `ExcelScript.Workbook` should always be the first parameter.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="cefc5-115">运行脚本时，`main` 函数中的代码将运行。</span><span class="sxs-lookup"><span data-stu-id="cefc5-115">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="cefc5-116">`main` 可以调用脚本中的其他函数，但是该函数中未包含的代码将不会运行。</span><span class="sxs-lookup"><span data-stu-id="cefc5-116">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span> <span data-ttu-id="cefc5-117">脚本无法调用其他 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="cefc5-117">Scripts cannot invoke or call other Office Scripts.</span></span>

<span data-ttu-id="cefc5-118">通过 [Power Automate](https://flow.microsoft.com)，可以在流中连接脚本。</span><span class="sxs-lookup"><span data-stu-id="cefc5-118">[Power Automate](https://flow.microsoft.com) allows you to connect scripts in flows.</span></span> <span data-ttu-id="cefc5-119">数据通过 `main` 方法的参数和返回在脚本和流之间传递。</span><span class="sxs-lookup"><span data-stu-id="cefc5-119">Data is passed between the scripts and the flow through the parameters and returns of the`main` method.</span></span> <span data-ttu-id="cefc5-120">[使用 Power Automate 运行 Office 脚本](power-automate-integration.md) 中详细介绍了如何集成 Office 脚本和 Power Automate。</span><span class="sxs-lookup"><span data-stu-id="cefc5-120">How to integrate Office Scripts with Power Automate is covered in detail in [Run Office Scripts with Power Automate](power-automate-integration.md).</span></span>

## <a name="object-model-overview"></a><span data-ttu-id="cefc5-121">对象模型概述</span><span class="sxs-lookup"><span data-stu-id="cefc5-121">Object model overview</span></span>

<span data-ttu-id="cefc5-122">要编写脚本，需要了解 Office 脚本 API 的组合方式。</span><span class="sxs-lookup"><span data-stu-id="cefc5-122">To write a script, you need to understand how the Office Scripts APIs fit together.</span></span> <span data-ttu-id="cefc5-123">工作簿的组件之间彼此有着特定的关系。</span><span class="sxs-lookup"><span data-stu-id="cefc5-123">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="cefc5-124">这些关系在许多方面与 Excel UI 的关系匹配。</span><span class="sxs-lookup"><span data-stu-id="cefc5-124">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="cefc5-125">一个 **Workbook** 包含一个或多个 **Worksheet**。</span><span class="sxs-lookup"><span data-stu-id="cefc5-125">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="cefc5-126">**Worksheet** 可通过 **Range** 对象访问单元格。</span><span class="sxs-lookup"><span data-stu-id="cefc5-126">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="cefc5-127">**Range** 代表一组连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="cefc5-127">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="cefc5-128">**Range** 用于创建和放置 **Table**、**Chart** 和 **Shape** 以及其他数据可视化或组织对象。</span><span class="sxs-lookup"><span data-stu-id="cefc5-128">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="cefc5-129">**Worksheet** 包含单个工作表中存在的那些数据对象的集合。</span><span class="sxs-lookup"><span data-stu-id="cefc5-129">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="cefc5-130">**Workbook** 包含整个 **Workbook** 的某些数据对象（例如，**Table**）的集合。</span><span class="sxs-lookup"><span data-stu-id="cefc5-130">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

## <a name="workbook"></a><span data-ttu-id="cefc5-131">工作簿</span><span class="sxs-lookup"><span data-stu-id="cefc5-131">Workbook</span></span>

<span data-ttu-id="cefc5-132">每个脚本都会由 `main` 函数提供一个 `Workbook` 类型的 `workbook` 对象。</span><span class="sxs-lookup"><span data-stu-id="cefc5-132">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="cefc5-133">这表示顶层对象，你的脚本将通过该对象与 Excel 工作簿进行交互。</span><span class="sxs-lookup"><span data-stu-id="cefc5-133">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="cefc5-134">以下脚本将获取工作簿中的活动工作表并记录其名称。</span><span class="sxs-lookup"><span data-stu-id="cefc5-134">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a><span data-ttu-id="cefc5-135">Ranges</span><span class="sxs-lookup"><span data-stu-id="cefc5-135">Ranges</span></span>

<span data-ttu-id="cefc5-136">Range 是工作簿中的一组连续单元格。</span><span class="sxs-lookup"><span data-stu-id="cefc5-136">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="cefc5-137">脚本通常使用 A1 样式表示法（例如，对于列 **B** 和行 **3** 中单个单元格，即 **B3** 或从列 **C** 至列 **F** 和行 **2** 至行 **4** 的单元格，即 **C2:F4**）来定义范围。</span><span class="sxs-lookup"><span data-stu-id="cefc5-137">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="cefc5-138">Range 有三个核心属性：值、公式和格式。</span><span class="sxs-lookup"><span data-stu-id="cefc5-138">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="cefc5-139">这些属性将获取或设置单元格值、要计算的公式以及单元格的视觉对象格式。</span><span class="sxs-lookup"><span data-stu-id="cefc5-139">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="cefc5-140">它们可通过 `getValues`、`getFormulas` 和 `getFormat` 进行访问。</span><span class="sxs-lookup"><span data-stu-id="cefc5-140">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="cefc5-141">值和公式可通过 `setValues` 和 `setFormulas` 进行更改，而格式则是由单独设置的多个较小对象组成的 `RangeFormat` 对象。</span><span class="sxs-lookup"><span data-stu-id="cefc5-141">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="cefc5-142">Range 使用二维数组管理信息。</span><span class="sxs-lookup"><span data-stu-id="cefc5-142">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="cefc5-143">有关在 Office 脚本框架中处理数组的详细信息，请参阅 [使用范围工作](javascript-objects.md#work-with-ranges)。</span><span class="sxs-lookup"><span data-stu-id="cefc5-143">For more information on handling arrays in the Office Scripts framework, see [Work with ranges](javascript-objects.md#work-with-ranges).</span></span>

### <a name="range-sample"></a><span data-ttu-id="cefc5-144">Range 示例</span><span class="sxs-lookup"><span data-stu-id="cefc5-144">Range sample</span></span>

<span data-ttu-id="cefc5-145">以下示例显示了如何创建销售记录。</span><span class="sxs-lookup"><span data-stu-id="cefc5-145">The following sample shows how to create sales records.</span></span> <span data-ttu-id="cefc5-146">该脚本使用 `Range` 对象来设置值、公式和部分格式。</span><span class="sxs-lookup"><span data-stu-id="cefc5-146">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

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

<span data-ttu-id="cefc5-147">运行此脚本将在当前工作表中创建以下数据：</span><span class="sxs-lookup"><span data-stu-id="cefc5-147">Running this script creates the following data in the current worksheet:</span></span>

:::image type="content" source="../images/range-sample.png" alt-text="包含由值行、公式列和带格式的标头组成的销售记录的工作表":::

## <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="cefc5-149">Chart、Table 和其他数据对象</span><span class="sxs-lookup"><span data-stu-id="cefc5-149">Charts, tables, and other data objects</span></span>

<span data-ttu-id="cefc5-150">脚本可以在 Excel 中创建和设置数据结构和可视化效果。</span><span class="sxs-lookup"><span data-stu-id="cefc5-150">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="cefc5-151">Table 和 Chart 是最常用的两个对象，但是 API 支持数据透视表、形状和图像等。</span><span class="sxs-lookup"><span data-stu-id="cefc5-151">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="cefc5-152">这些都存储在集合中，本文后面将对该内容进行讨论。</span><span class="sxs-lookup"><span data-stu-id="cefc5-152">These are stored in collections, which will be discussed later in this article.</span></span>

### <a name="create-a-table"></a><span data-ttu-id="cefc5-153">创建表格</span><span class="sxs-lookup"><span data-stu-id="cefc5-153">Create a table</span></span>

<span data-ttu-id="cefc5-p113">通过使用数据填充区域创建表。自动将格式设置和表格控件（如筛选器）应用到区域。</span><span class="sxs-lookup"><span data-stu-id="cefc5-p113">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="cefc5-156">以下脚本使用上一个示例中的范围创建一个表。</span><span class="sxs-lookup"><span data-stu-id="cefc5-156">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="cefc5-157">在工作表上使用之前的数据运行此脚本将创建下表：</span><span class="sxs-lookup"><span data-stu-id="cefc5-157">Running this script on the worksheet with the previous data creates the following table:</span></span>

:::image type="content" source="../images/table-sample.png" alt-text="包含根据以前销售记录所创建的表的工作表":::

### <a name="create-a-chart"></a><span data-ttu-id="cefc5-159">创建图表</span><span class="sxs-lookup"><span data-stu-id="cefc5-159">Create a chart</span></span>

<span data-ttu-id="cefc5-160">创建图表以直观显示某个范围内的数据。</span><span class="sxs-lookup"><span data-stu-id="cefc5-160">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="cefc5-161">脚本支持数十种图表类型，每种都可以根据需要进行自定义。</span><span class="sxs-lookup"><span data-stu-id="cefc5-161">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="cefc5-162">下面的脚本为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方，并将其设置为 100 像素。</span><span class="sxs-lookup"><span data-stu-id="cefc5-162">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

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

<span data-ttu-id="cefc5-163">在工作表上使用上一个表运行此脚本将创建以下图表：</span><span class="sxs-lookup"><span data-stu-id="cefc5-163">Running this script on the worksheet with the previous table creates the following chart:</span></span>

:::image type="content" source="../images/chart-sample.png" alt-text="显示上一个销售记录中三个项目的数量的柱形图":::

## <a name="collections"></a><span data-ttu-id="cefc5-165">集合</span><span class="sxs-lookup"><span data-stu-id="cefc5-165">Collections</span></span>

<span data-ttu-id="cefc5-166">当 Excel 对象具有一个或多个相同类型对象的集合时，则将它们存储在数组中。</span><span class="sxs-lookup"><span data-stu-id="cefc5-166">When an Excel object has a collection of one or more objects of the same type, it stores them in an array.</span></span> <span data-ttu-id="cefc5-167">例如，`Workbook` 对象包含一个 `Worksheet[]`。</span><span class="sxs-lookup"><span data-stu-id="cefc5-167">For example, a `Workbook` object contains a `Worksheet[]`.</span></span> <span data-ttu-id="cefc5-168">此数组由 `Workbook.getWorksheets()` 方法访问。</span><span class="sxs-lookup"><span data-stu-id="cefc5-168">This array is accessed by the `Workbook.getWorksheets()` method.</span></span> <span data-ttu-id="cefc5-169">复数形式的 `get` 方法（如 `Worksheet.getCharts()`）将整个对象集合作为数组返回。</span><span class="sxs-lookup"><span data-stu-id="cefc5-169">`get` methods that are plural, such as `Worksheet.getCharts()`, return the entire object collection as an array.</span></span> <span data-ttu-id="cefc5-170">你将在整个 Office 脚本 API 中查看此模式：`Worksheet` 对象采用 `getTables()` 方法返回 `Table[]`，`Table` 对象采用 `getColumns()` 方法返回 `TableColumn[]`，以此类推。</span><span class="sxs-lookup"><span data-stu-id="cefc5-170">You'll see this pattern throughout the Office Scripts APIs: the `Worksheet` object has a `getTables()` method that returns a `Table[]`, the `Table` object has a `getColumns()` method that returns a `TableColumn[]`, as so on.</span></span>

<span data-ttu-id="cefc5-171">返回的数组是一个普通数组，因此所有常规数组操作均可用于脚本。</span><span class="sxs-lookup"><span data-stu-id="cefc5-171">The returned array is a normal array, so all the regular array operations are available for your script.</span></span> <span data-ttu-id="cefc5-172">你还可以使用数组索引值访问集合中的单个对象。</span><span class="sxs-lookup"><span data-stu-id="cefc5-172">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="cefc5-173">例如，`workbook.getTables()[0]` 将返回集合中的第一个表格。</span><span class="sxs-lookup"><span data-stu-id="cefc5-173">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="cefc5-174">有关通过 Office 脚本框架使用内置数组功能的详细信息，请参阅 [使用集合工作](javascript-objects.md#work-with-collections)。</span><span class="sxs-lookup"><span data-stu-id="cefc5-174">For more information on using the built-in array functionality with the Office Scripts framework, see [Work with collections](javascript-objects.md#work-with-collections).</span></span> 

<span data-ttu-id="cefc5-175">此外，还可通过 `get` 方法从集合中访问单个对象。</span><span class="sxs-lookup"><span data-stu-id="cefc5-175">Individual objects are also accessed from the collection through a `get` method.</span></span> <span data-ttu-id="cefc5-176">单数形式的 `get` 方法（如 `Worksheet.getTable(name)`）返回单个对象，并且需要特定对象的 ID 或名称。</span><span class="sxs-lookup"><span data-stu-id="cefc5-176">`get` methods that are singular, such as `Worksheet.getTable(name)`, return a single object and require an ID or name for the specific object.</span></span> <span data-ttu-id="cefc5-177">此 ID 或名称通常由脚本或通过 Excel UI 设置。</span><span class="sxs-lookup"><span data-stu-id="cefc5-177">This ID or name is usually set by the script or through the Excel UI.</span></span>

<span data-ttu-id="cefc5-p118">以下脚本获取工作簿中所有表。然后可确保显示标题、筛选按钮可见，并且表格样式设置为“TableStyleLight1”。</span><span class="sxs-lookup"><span data-stu-id="cefc5-p118">The following script gets all tables in the workbook. It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

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

## <a name="add-excel-objects-with-a-script"></a><span data-ttu-id="cefc5-180">使用脚本添加 Excel 对象</span><span class="sxs-lookup"><span data-stu-id="cefc5-180">Add Excel objects with a script</span></span>

<span data-ttu-id="cefc5-181">通过调用可在父对象上使用的相应 `add` 方法，可以以编程方式添加文档对象，如表格或图表。</span><span class="sxs-lookup"><span data-stu-id="cefc5-181">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cefc5-182">不要手动将对象添加到集合数组。</span><span class="sxs-lookup"><span data-stu-id="cefc5-182">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="cefc5-183">请在父对象上使用 `add` 方法。例如，使用 `Worksheet.addTable` 方法向 `Worksheet` 添加 `Table`。</span><span class="sxs-lookup"><span data-stu-id="cefc5-183">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="cefc5-184">以下脚本将在 Excel 工作簿中的第一个工作表上创建一个表格。</span><span class="sxs-lookup"><span data-stu-id="cefc5-184">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="cefc5-185">请注意，所创建的表格是通过 `addTable` 方法返回的。</span><span class="sxs-lookup"><span data-stu-id="cefc5-185">Note that the created table is returned by the `addTable` method.</span></span>

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
> <span data-ttu-id="cefc5-186">大多数 Excel 对象都具有 `setName` 方法。</span><span class="sxs-lookup"><span data-stu-id="cefc5-186">Most Excel objects have a `setName` method.</span></span> <span data-ttu-id="cefc5-187">通过这一方法，可稍后在同一工作簿的脚本或其他脚本中轻松访问 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="cefc5-187">This gives you an easy way to access Excel objects later in the script or in other scripts for the same workbook.</span></span>

### <a name="verify-an-object-exists-in-the-collection"></a><span data-ttu-id="cefc5-188">验证集合中是否存在某个对象</span><span class="sxs-lookup"><span data-stu-id="cefc5-188">Verify an object exists in the collection</span></span>

<span data-ttu-id="cefc5-189">在继续之前，脚本通常需要检查表或类似对象是否存在。</span><span class="sxs-lookup"><span data-stu-id="cefc5-189">Scripts often need to check if a table or similar object exists before continuing.</span></span> <span data-ttu-id="cefc5-190">使用脚本或 Excel UI 提供的名称确定必要的对象，并执行相应操作。</span><span class="sxs-lookup"><span data-stu-id="cefc5-190">Use the names given by scripts or through the Excel UI to identify necessary objects and act accordingly.</span></span> <span data-ttu-id="cefc5-191">请求的对象不在集合中时，`get` 方法返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="cefc5-191">`get` methods return `undefined` when the requested object is not in the collection.</span></span>

<span data-ttu-id="cefc5-192">以下脚本请求名为“MyTable”的表，并使用 `if...else` 语句检查是否已找到该表。</span><span class="sxs-lookup"><span data-stu-id="cefc5-192">The following script requests a table named "MyTable" and uses an `if...else` statement to check if the table was found.</span></span>

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

<span data-ttu-id="cefc5-193">Office 脚本中的一种常见模式是在每次运行脚本时重新创建表、图表或其他对象。</span><span class="sxs-lookup"><span data-stu-id="cefc5-193">A common pattern in Office Scripts is to recreate a table, chart, or other object every time the script is run.</span></span> <span data-ttu-id="cefc5-194">如果不需要旧数据，最好先删除旧对象，然后再创建新对象。</span><span class="sxs-lookup"><span data-stu-id="cefc5-194">If you don't need the old data, it's best to delete the old object before creating the new one.</span></span> <span data-ttu-id="cefc5-195">此操作可避免出现名称冲突或已由其他用户引入的其他差异。</span><span class="sxs-lookup"><span data-stu-id="cefc5-195">This avoids name conflicts or other differences that may have been introduced by other users.</span></span>

<span data-ttu-id="cefc5-196">以下脚本删除名为“MyTable”的表，如果存在该表，则添加名称相同的新表。</span><span class="sxs-lookup"><span data-stu-id="cefc5-196">The following script removes the table named "MyTable", if it is present, then adds a new table with the same name.</span></span>

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

## <a name="remove-excel-objects-with-a-script"></a><span data-ttu-id="cefc5-197">使用脚本删除 Excel 对象</span><span class="sxs-lookup"><span data-stu-id="cefc5-197">Remove Excel objects with a script</span></span>

<span data-ttu-id="cefc5-198">若要删除对象，请调用对象的 `delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="cefc5-198">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="cefc5-199">与添加对象一样，不要手动从集合数组中删除对象。</span><span class="sxs-lookup"><span data-stu-id="cefc5-199">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="cefc5-200">请在集合类型的对象上使用 `delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="cefc5-200">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="cefc5-201">例如，使用 `Table.delete`从 `Worksheet` 中删除 `Table`。</span><span class="sxs-lookup"><span data-stu-id="cefc5-201">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="cefc5-202">以下脚本将删除工作簿中的第一个工作表。</span><span class="sxs-lookup"><span data-stu-id="cefc5-202">The following script removes the first worksheet in the workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a><span data-ttu-id="cefc5-203">进一步了解对象模型</span><span class="sxs-lookup"><span data-stu-id="cefc5-203">Further reading on the object model</span></span>

<span data-ttu-id="cefc5-204">[Office 脚本 API 参考文档](/javascript/api/office-scripts/overview)是 Office 脚本中使用的对象的完整列表。</span><span class="sxs-lookup"><span data-stu-id="cefc5-204">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="cefc5-205">在这里，可以使用目录导航到想进一步了解的任何课程。</span><span class="sxs-lookup"><span data-stu-id="cefc5-205">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="cefc5-206">以下是几个经常查看的页面。</span><span class="sxs-lookup"><span data-stu-id="cefc5-206">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="cefc5-207">Chart</span><span class="sxs-lookup"><span data-stu-id="cefc5-207">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="cefc5-208">Comment</span><span class="sxs-lookup"><span data-stu-id="cefc5-208">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="cefc5-209">PivotTable</span><span class="sxs-lookup"><span data-stu-id="cefc5-209">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="cefc5-210">区域</span><span class="sxs-lookup"><span data-stu-id="cefc5-210">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="cefc5-211">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="cefc5-211">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="cefc5-212">Shape</span><span class="sxs-lookup"><span data-stu-id="cefc5-212">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="cefc5-213">Table</span><span class="sxs-lookup"><span data-stu-id="cefc5-213">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="cefc5-214">Workbook</span><span class="sxs-lookup"><span data-stu-id="cefc5-214">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="cefc5-215">Worksheet</span><span class="sxs-lookup"><span data-stu-id="cefc5-215">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="cefc5-216">另请参阅</span><span class="sxs-lookup"><span data-stu-id="cefc5-216">See also</span></span>

- [<span data-ttu-id="cefc5-217">在 Excel 网页版中录制、编辑和创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="cefc5-217">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="cefc5-218">在 Excel 网页版中使用 Office 脚本读取工作簿数据</span><span class="sxs-lookup"><span data-stu-id="cefc5-218">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="cefc5-219">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="cefc5-219">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="cefc5-220">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="cefc5-220">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="cefc5-221">Office 脚本中的最佳实践</span><span class="sxs-lookup"><span data-stu-id="cefc5-221">Best practices in Office Scripts</span></span>](best-practices.md)
