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
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="9abc9-103">Excel 网页版中 Office 脚本的脚本基础（预览）</span><span class="sxs-lookup"><span data-stu-id="9abc9-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="9abc9-104">本文将介绍 Office 脚本技术方面的知识。</span><span class="sxs-lookup"><span data-stu-id="9abc9-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="9abc9-105">你将了解 Excel 对象如何协同工作以及代码编辑器如何与工作簿同步。</span><span class="sxs-lookup"><span data-stu-id="9abc9-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a><span data-ttu-id="9abc9-106">`main` 函数</span><span class="sxs-lookup"><span data-stu-id="9abc9-106">`main` function</span></span>

<span data-ttu-id="9abc9-107">每个 Office 脚本都必须包含以 `ExcelScript.Workbook` 类型作为第一参数的 `main` 函数。</span><span class="sxs-lookup"><span data-stu-id="9abc9-107">Each Office Script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="9abc9-108">执行函数时，Excel 应用程序通过提供相应工作簿作为第一个参数来调用此 `main` 函数。</span><span class="sxs-lookup"><span data-stu-id="9abc9-108">When the function is executed, the Excel application invokes this `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="9abc9-109">因此，在记录脚本或从代码编辑器创建新脚本后，请务必不要再修改 `main` 函数的基本签名。</span><span class="sxs-lookup"><span data-stu-id="9abc9-109">Hence, it is important to not modify the basic signature of the `main` function once you have either recorded the script or created a new script from the code editor.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="9abc9-110">运行脚本时，`main` 函数中的代码将运行。</span><span class="sxs-lookup"><span data-stu-id="9abc9-110">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="9abc9-111">`main` 可以调用脚本中的其他函数，但是该函数中未包含的代码将不会运行。</span><span class="sxs-lookup"><span data-stu-id="9abc9-111">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

> [!CAUTION]
> <span data-ttu-id="9abc9-112">如果你的 `main` 函数看起来像 `async function main(context: Excel.RequestContext)`，那你的脚本使用的是旧版异步 API 模型。</span><span class="sxs-lookup"><span data-stu-id="9abc9-112">If your `main` function looks like `async function main(context: Excel.RequestContext)`, your script is using the older async API model.</span></span> <span data-ttu-id="9abc9-113">有关详细信息（包括如何将你的脚本转换为当前 API 模型），请参阅[支持使用异步 API 的旧 Office 脚本](excel-async-model.md)。</span><span class="sxs-lookup"><span data-stu-id="9abc9-113">For more information (including how to convert your script to the current API model), refer to [Support older Office Scripts that use the Async APIs](excel-async-model.md).</span></span>

## <a name="object-model"></a><span data-ttu-id="9abc9-114">对象模型</span><span class="sxs-lookup"><span data-stu-id="9abc9-114">Object model</span></span>

<span data-ttu-id="9abc9-115">若要编写脚本，你需要了解 Office 脚本 API 如何组合在一起。</span><span class="sxs-lookup"><span data-stu-id="9abc9-115">To write a script, you need to understand how the Office Script APIs fit together.</span></span> <span data-ttu-id="9abc9-116">工作簿的组件之间彼此有着特定的关系。</span><span class="sxs-lookup"><span data-stu-id="9abc9-116">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="9abc9-117">这些关系在许多方面与 Excel UI 的关系匹配。</span><span class="sxs-lookup"><span data-stu-id="9abc9-117">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="9abc9-118">一个 **Workbook** 包含一个或多个 **Worksheet**。</span><span class="sxs-lookup"><span data-stu-id="9abc9-118">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="9abc9-119">**Worksheet** 可通过 **Range** 对象访问单元格。</span><span class="sxs-lookup"><span data-stu-id="9abc9-119">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="9abc9-120">**Range** 代表一组连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="9abc9-120">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="9abc9-121">**Range** 用于创建和放置 **Table**、**Chart** 和 **Shape** 以及其他数据可视化或组织对象。</span><span class="sxs-lookup"><span data-stu-id="9abc9-121">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="9abc9-122">**Worksheet** 包含单个工作表中存在的那些数据对象的集合。</span><span class="sxs-lookup"><span data-stu-id="9abc9-122">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="9abc9-123">**Workbook** 包含整个 **Workbook** 的某些数据对象（例如，**Table**）的集合。</span><span class="sxs-lookup"><span data-stu-id="9abc9-123">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="workbook"></a><span data-ttu-id="9abc9-124">工作簿</span><span class="sxs-lookup"><span data-stu-id="9abc9-124">Workbook</span></span>

<span data-ttu-id="9abc9-125">每个脚本都会由 `main` 函数提供一个 `Workbook` 类型的 `workbook` 对象。</span><span class="sxs-lookup"><span data-stu-id="9abc9-125">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="9abc9-126">这表示顶层对象，你的脚本将通过该对象与 Excel 工作簿进行交互。</span><span class="sxs-lookup"><span data-stu-id="9abc9-126">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="9abc9-127">以下脚本将获取工作簿中的活动工作表并记录其名称。</span><span class="sxs-lookup"><span data-stu-id="9abc9-127">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

### <a name="ranges"></a><span data-ttu-id="9abc9-128">Ranges</span><span class="sxs-lookup"><span data-stu-id="9abc9-128">Ranges</span></span>

<span data-ttu-id="9abc9-129">Range 是工作簿中的一组连续单元格。</span><span class="sxs-lookup"><span data-stu-id="9abc9-129">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="9abc9-130">脚本通常使用 A1 样式表示法（例如，对于列 **B** 和行 **3** 中单个单元格，即 **B3** 或从列 **C** 至列 **F** 和行 **2** 至行 **4** 的单元格，即 **C2:F4**）来定义范围。</span><span class="sxs-lookup"><span data-stu-id="9abc9-130">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="9abc9-131">Range 有三个核心属性：值、公式和格式。</span><span class="sxs-lookup"><span data-stu-id="9abc9-131">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="9abc9-132">这些属性将获取或设置单元格值、要计算的公式以及单元格的视觉对象格式。</span><span class="sxs-lookup"><span data-stu-id="9abc9-132">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="9abc9-133">它们可通过 `getValues`、`getFormulas` 和 `getFormat` 进行访问。</span><span class="sxs-lookup"><span data-stu-id="9abc9-133">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="9abc9-134">值和公式可通过 `setValues` 和 `setFormulas` 进行更改，而格式则是由单独设置的多个较小对象组成的 `RangeFormat` 对象。</span><span class="sxs-lookup"><span data-stu-id="9abc9-134">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="9abc9-135">Range 使用二维数组管理信息。</span><span class="sxs-lookup"><span data-stu-id="9abc9-135">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="9abc9-136">有关如何在 Office 脚本框架中处理这些数组的详细信息，请参阅[《在 Office 脚本中使用内置的 JavaScript 对象》的“使用区域”部分](javascript-objects.md#working-with-ranges)。</span><span class="sxs-lookup"><span data-stu-id="9abc9-136">Read the [Working with ranges section of Using built-in JavaScript objects in Office Scripts](javascript-objects.md#working-with-ranges) for more information on handling those arrays in the Office Scripts framework.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="9abc9-137">Range 示例</span><span class="sxs-lookup"><span data-stu-id="9abc9-137">Range sample</span></span>

<span data-ttu-id="9abc9-138">以下示例显示了如何创建销售记录。</span><span class="sxs-lookup"><span data-stu-id="9abc9-138">The following sample shows how to create sales records.</span></span> <span data-ttu-id="9abc9-139">该脚本使用 `Range` 对象来设置值、公式和部分格式。</span><span class="sxs-lookup"><span data-stu-id="9abc9-139">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

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

<span data-ttu-id="9abc9-140">运行此脚本将在当前工作表中创建以下数据：</span><span class="sxs-lookup"><span data-stu-id="9abc9-140">Running this script creates the following data in the current worksheet:</span></span>

:::image type="content" source="../images/range-sample.png" alt-text="包含由值行、公式列和带格式的标头组成的销售记录的工作表。":::

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="9abc9-142">Chart、Table 和其他数据对象</span><span class="sxs-lookup"><span data-stu-id="9abc9-142">Charts, tables, and other data objects</span></span>

<span data-ttu-id="9abc9-143">脚本可以在 Excel 中创建和设置数据结构和可视化效果。</span><span class="sxs-lookup"><span data-stu-id="9abc9-143">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="9abc9-144">Table 和 Chart 是最常用的两个对象，但是 API 支持数据透视表、形状和图像等。</span><span class="sxs-lookup"><span data-stu-id="9abc9-144">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="9abc9-145">这些都存储在集合中，本文后面将对该内容进行讨论。</span><span class="sxs-lookup"><span data-stu-id="9abc9-145">These are stored in collections, which will be discussed later in this article.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="9abc9-146">创建表</span><span class="sxs-lookup"><span data-stu-id="9abc9-146">Creating a table</span></span>

<span data-ttu-id="9abc9-147">通过使用数据填充范围创建表。</span><span class="sxs-lookup"><span data-stu-id="9abc9-147">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="9abc9-148">会将格式设置和表控件（如筛选器）自动应用到该范围。</span><span class="sxs-lookup"><span data-stu-id="9abc9-148">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="9abc9-149">以下脚本使用上一个示例中的范围创建一个表。</span><span class="sxs-lookup"><span data-stu-id="9abc9-149">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="9abc9-150">在工作表上使用之前的数据运行此脚本将创建下表：</span><span class="sxs-lookup"><span data-stu-id="9abc9-150">Running this script on the worksheet with the previous data creates the following table:</span></span>

:::image type="content" source="../images/table-sample.png" alt-text="包含根据以前销售记录所创建表的工作表。":::

#### <a name="creating-a-chart"></a><span data-ttu-id="9abc9-152">创建图表</span><span class="sxs-lookup"><span data-stu-id="9abc9-152">Creating a chart</span></span>

<span data-ttu-id="9abc9-153">创建图表以直观显示某个范围内的数据。</span><span class="sxs-lookup"><span data-stu-id="9abc9-153">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="9abc9-154">脚本支持数十种图表类型，每种都可以根据需要进行自定义。</span><span class="sxs-lookup"><span data-stu-id="9abc9-154">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="9abc9-155">下面的脚本为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方，并将其设置为 100 像素。</span><span class="sxs-lookup"><span data-stu-id="9abc9-155">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

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

<span data-ttu-id="9abc9-156">在工作表上使用上一个表运行此脚本将创建以下图表：</span><span class="sxs-lookup"><span data-stu-id="9abc9-156">Running this script on the worksheet with the previous table creates the following chart:</span></span>

:::image type="content" source="../images/chart-sample.png" alt-text="一个柱形图，显示上一个销售记录中三个项目的数量。":::

### <a name="collections-and-other-object-relations"></a><span data-ttu-id="9abc9-158">集合和其他对象关系</span><span class="sxs-lookup"><span data-stu-id="9abc9-158">Collections and other object relations</span></span>

<span data-ttu-id="9abc9-159">任何子对象都可通过其父对象访问。</span><span class="sxs-lookup"><span data-stu-id="9abc9-159">Any child object can be accessed through its parent object.</span></span> <span data-ttu-id="9abc9-160">例如，可从 `Workbook` 对象中读取 `Worksheets`。</span><span class="sxs-lookup"><span data-stu-id="9abc9-160">For example, you can read `Worksheets` from the `Workbook` object.</span></span> <span data-ttu-id="9abc9-161">父类上将会有一个相关的 `get` 方法（例如 `Workbook.getWorksheets()` 或 `Workbook.getWorksheet(name)` ）。</span><span class="sxs-lookup"><span data-stu-id="9abc9-161">There will be a related `get` method on the parent class that (e.g., `Workbook.getWorksheets()` or `Workbook.getWorksheet(name)`).</span></span> <span data-ttu-id="9abc9-162">单数形式的 `get` 方法将返回单个对象，并且需要特定对象的 ID 或名称（如工作表名称）。</span><span class="sxs-lookup"><span data-stu-id="9abc9-162">`get` methods that are singular return a single object and require an ID or name for the specific object (such as the name of a worksheet).</span></span> <span data-ttu-id="9abc9-163">复数形式的 `get` 方法会将整个对象集合作为数组返回。</span><span class="sxs-lookup"><span data-stu-id="9abc9-163">`get` methods that are plural return the entire object collection as an array.</span></span> <span data-ttu-id="9abc9-164">如果集合为空，将得到一个空数组 (`[]`)。</span><span class="sxs-lookup"><span data-stu-id="9abc9-164">If the collection is empty, you'll get an empty array (`[]`).</span></span>

<span data-ttu-id="9abc9-165">检索到相应集合后，可在其上面使用常规数组操作（如获取其 `length` 或使用 `for`、`for..of` 或 `while` 循环进行迭代）或使用 TypeScript 数组方法（如 `map` 或 `forEach`）。</span><span class="sxs-lookup"><span data-stu-id="9abc9-165">Once the collection is retrieved, you can use regular array operations such as getting its `length` or use `for`, `for..of`, `while` loops for iteration or use TypeScript array methods such as `map`, `forEach` on them.</span></span> <span data-ttu-id="9abc9-166">你还可以使用数组索引值访问集合中的单个对象。</span><span class="sxs-lookup"><span data-stu-id="9abc9-166">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="9abc9-167">例如，`workbook.getTables()[0]` 将返回集合中的第一个表格。</span><span class="sxs-lookup"><span data-stu-id="9abc9-167">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="9abc9-168">请阅读[《在 Office 脚本中使用内置的 JavaScript 对象》的“使用集合”部分](javascript-objects.md#working-with-collections)，深入了解如何在 Office 脚本框架中使用内置数组功能。</span><span class="sxs-lookup"><span data-stu-id="9abc9-168">Read the [Working with collections section of Using built-in JavaScript objects in Office Scripts](javascript-objects.md#working-with-collections) to learn more about using built-in array functionality with the Office Scripts framework.</span></span>

<span data-ttu-id="9abc9-169">以下脚本将获取工作簿中的所有表格。</span><span class="sxs-lookup"><span data-stu-id="9abc9-169">The following script gets all tables in the workbook.</span></span> <span data-ttu-id="9abc9-170">然后，它将确保显示标题、筛选按钮可见并且将表格样式设置为“TableStyleLight1”。</span><span class="sxs-lookup"><span data-stu-id="9abc9-170">It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

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

#### <a name="adding-excel-objects-with-a-script"></a><span data-ttu-id="9abc9-171">使用脚本添加 Excel 对象</span><span class="sxs-lookup"><span data-stu-id="9abc9-171">Adding Excel objects with a script</span></span>

<span data-ttu-id="9abc9-172">通过调用可在父对象上使用的相应 `add` 方法，可以以编程方式添加文档对象，如表格或图表。</span><span class="sxs-lookup"><span data-stu-id="9abc9-172">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!NOTE]
> <span data-ttu-id="9abc9-173">不要手动将对象添加到集合数组。</span><span class="sxs-lookup"><span data-stu-id="9abc9-173">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="9abc9-174">请在父对象上使用 `add` 方法。例如，使用 `Worksheet.addTable` 方法向 `Worksheet` 添加 `Table`。</span><span class="sxs-lookup"><span data-stu-id="9abc9-174">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="9abc9-175">以下脚本将在 Excel 工作簿中的第一个工作表上创建一个表格。</span><span class="sxs-lookup"><span data-stu-id="9abc9-175">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="9abc9-176">请注意，所创建的表格是通过 `addTable` 方法返回的。</span><span class="sxs-lookup"><span data-stu-id="9abc9-176">Note that the created table is returned by the `addTable` method.</span></span>

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

## <a name="removing-excel-objects-with-a-script"></a><span data-ttu-id="9abc9-177">使用脚本删除 Excel 对象</span><span class="sxs-lookup"><span data-stu-id="9abc9-177">Removing Excel objects with a script</span></span>

<span data-ttu-id="9abc9-178">若要删除对象，请调用对象的 `delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="9abc9-178">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="9abc9-179">与添加对象一样，不要手动从集合数组中删除对象。</span><span class="sxs-lookup"><span data-stu-id="9abc9-179">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="9abc9-180">请在集合类型的对象上使用 `delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="9abc9-180">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="9abc9-181">例如，使用 `Table.delete`从 `Worksheet` 中删除 `Table`。</span><span class="sxs-lookup"><span data-stu-id="9abc9-181">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="9abc9-182">以下脚本将删除工作簿中的第一个工作表。</span><span class="sxs-lookup"><span data-stu-id="9abc9-182">The following script removes the first worksheet in the workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="9abc9-183">进一步了解对象模型</span><span class="sxs-lookup"><span data-stu-id="9abc9-183">Further reading on the object model</span></span>

<span data-ttu-id="9abc9-184">[Office 脚本 API 参考文档](/javascript/api/office-scripts/overview)是 Office 脚本中使用的对象的完整列表。</span><span class="sxs-lookup"><span data-stu-id="9abc9-184">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="9abc9-185">在这里，可以使用目录导航到想进一步了解的任何课程。</span><span class="sxs-lookup"><span data-stu-id="9abc9-185">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="9abc9-186">以下是几个经常查看的页面。</span><span class="sxs-lookup"><span data-stu-id="9abc9-186">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="9abc9-187">Chart</span><span class="sxs-lookup"><span data-stu-id="9abc9-187">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="9abc9-188">Comment</span><span class="sxs-lookup"><span data-stu-id="9abc9-188">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="9abc9-189">PivotTable</span><span class="sxs-lookup"><span data-stu-id="9abc9-189">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="9abc9-190">区域</span><span class="sxs-lookup"><span data-stu-id="9abc9-190">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="9abc9-191">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="9abc9-191">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="9abc9-192">Shape</span><span class="sxs-lookup"><span data-stu-id="9abc9-192">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="9abc9-193">Table</span><span class="sxs-lookup"><span data-stu-id="9abc9-193">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="9abc9-194">Workbook</span><span class="sxs-lookup"><span data-stu-id="9abc9-194">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="9abc9-195">Worksheet</span><span class="sxs-lookup"><span data-stu-id="9abc9-195">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="9abc9-196">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9abc9-196">See also</span></span>

- [<span data-ttu-id="9abc9-197">在 Excel 网页版中录制、编辑和创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="9abc9-197">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="9abc9-198">在 Excel 网页版中使用 Office 脚本读取工作簿数据</span><span class="sxs-lookup"><span data-stu-id="9abc9-198">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="9abc9-199">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="9abc9-199">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="9abc9-200">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="9abc9-200">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
