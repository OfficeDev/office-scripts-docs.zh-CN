---
title: Office 脚本中的区域基础知识
description: 了解有关在 Office 脚本中使用 Range 对象的基础知识。
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 73eeba086aace6262c624de9074ffb301f6532bd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571104"
---
# <a name="range-basics"></a><span data-ttu-id="7a1f3-103">Range 基础知识</span><span class="sxs-lookup"><span data-stu-id="7a1f3-103">Range basics</span></span>

<span data-ttu-id="7a1f3-104">`Range` 是 Office Scripts Excel 对象模型中的基础对象。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-104">`Range` is the foundational object within the Office Scripts Excel object model.</span></span> <span data-ttu-id="7a1f3-105">[区域 API](/javascript/api/office-scripts/excelscript/excelscript.range) 允许访问网格上可用的数据和格式，并链接 Excel 内的其他关键对象，如工作表、表、图表等。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-105">[Range APIs](/javascript/api/office-scripts/excelscript/excelscript.range) allow access to both data and format available on the grid and link other key objects within Excel such as worksheets, tables, charts, etc.</span></span>

<span data-ttu-id="7a1f3-106">区域使用其地址（如"A1：B4"）或已命名项（它是一组给定单元格的命名键）进行标识。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-106">A range is identified using its address such as "A1:B4" or using a named-item, which is a named key for a given set of cells.</span></span> <span data-ttu-id="7a1f3-107">在 Excel 对象模型中，单元格和单元格组都称为 _range_。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-107">In the Excel object model, both a cell and group of cells are referred as _range_.</span></span> <span data-ttu-id="7a1f3-108">`Range` 可以包含单元格级属性（如单元格内的数据），还可以包含单元格和单元格级属性（如格式、边框等）。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-108">`Range` can contain cell-level attributes such as data within a cell and also cell and cells-level attributes such as format, borders, etc.</span></span>

<span data-ttu-id="7a1f3-109">`Range` 还可通过用户选择（至少包含一个单元格）获取。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-109">`Range` can also be obtained via user's selection that consists of at least one cell.</span></span> <span data-ttu-id="7a1f3-110">与区域交互时，必须明确这些单元格和范围关系。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-110">As you interact with the range, it's important to keep these cell and range relationships clear.</span></span>

<span data-ttu-id="7a1f3-111">以下是 getter、setter 和其他在脚本中最常用的有用方法的核心集。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-111">Following are the core set of getters, setters, and other useful methods most often used in scripts.</span></span> <span data-ttu-id="7a1f3-112">这是 API 旅程的一个很好起点。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-112">This is a great starting point for your API journey.</span></span> <span data-ttu-id="7a1f3-113">以下各节对方法进行分组，并有助于在开始解锁对象的 API 时 `Range` 构建一个精神模型。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-113">The later sections group the methods and help to build a mental model as you begin to unlock the `Range` object's APIs.</span></span>

## <a name="example-scripts"></a><span data-ttu-id="7a1f3-114">示例脚本</span><span class="sxs-lookup"><span data-stu-id="7a1f3-114">Example scripts</span></span>

* [<span data-ttu-id="7a1f3-115">基本读写</span><span class="sxs-lookup"><span data-stu-id="7a1f3-115">Basic read and write</span></span>](#basic-read-and-write)
* [<span data-ttu-id="7a1f3-116">在工作表末尾添加行</span><span class="sxs-lookup"><span data-stu-id="7a1f3-116">Add row at the end of worksheet</span></span>](#add-row-at-the-end-of-worksheet)
* [<span data-ttu-id="7a1f3-117">清除列筛选器</span><span class="sxs-lookup"><span data-stu-id="7a1f3-117">Clear column filter</span></span>](clear-table-filter-for-active-cell.md)
* [<span data-ttu-id="7a1f3-118">使用唯一的颜色为每个单元格设置颜色</span><span class="sxs-lookup"><span data-stu-id="7a1f3-118">Color each cell with unique color</span></span>](#color-each-cell-with-unique-color)
* [<span data-ttu-id="7a1f3-119">使用二维数组更新二维 (二维) 区域</span><span class="sxs-lookup"><span data-stu-id="7a1f3-119">Update range with values using 2-dimensional (2D) array</span></span>](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a><span data-ttu-id="7a1f3-120">基本读写</span><span class="sxs-lookup"><span data-stu-id="7a1f3-120">Basic read and write</span></span>

```TypeScript
/**
 * This script demonstrates basic read-write operations on the Range object.
 */
function main(workbook: ExcelScript.Workbook) {
  const cell = workbook.getActiveCell();
  const prevValue = cell.getValue();
  if (prevValue) {
      console.log(`Active cell's value is: ${prevValue}`);
  } else {
      console.log("Setting active cell's value..");
      cell.setValue("Sample");
  }

  // Get cell next to the right column and set its value and fill color.
  const nextCell = cell.getOffsetRange(0,1);
  nextCell.setValue("Next cell");
  console.log(`Next cell's address is: ${nextCell.getAddress()}`);
  console.log("Setting fill color and font color of next cell...");
  nextCell.getFormat().getFill().setColor("Magenta");
  nextCell.getFormat().getFill().setColor("Cyan");

  // Get the target range address to update with 2-dimensional value.
  const dataRange = nextCell.getOffsetRange(1, 0).getResizedRange(2, 1);
  const DATA = [
    [10, 7],
    [8, 15],
    [12, 1]
  ];
  console.log(`Updating range ${dataRange.getAddress()} with values: ${DATA}`);
  dataRange.setValues(DATA);

  // Formula range.
  const formulaRange = dataRange.getOffsetRange(3, 0).getRow(0);
  console.log(`Updating formula for range: ${formulaRange.getAddress()}`)
  // Since relative formula is being set, we can set the formula of the entire range to the same value.
  formulaRange.setFormulaR1C1("=SUM(R[-3]C:R[-1]C)");
  console.log(`Updating number format for range: ${formulaRange.getAddress()}`)
  // Since the number format is common to the entire range, we can set it to a common format.
  formulaRange.setNumberFormat("0.00");
  return;
}
```

### <a name="add-row-at-the-end-of-worksheet"></a><span data-ttu-id="7a1f3-121">在工作表末尾添加行</span><span class="sxs-lookup"><span data-stu-id="7a1f3-121">Add row at the end of worksheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for the update.
    if (usedRange) {
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);
    targetRange.setValues([data]);
    return;
}
```

### <a name="color-each-cell-with-unique-color"></a><span data-ttu-id="7a1f3-122">使用唯一的颜色为每个单元格设置颜色</span><span class="sxs-lookup"><span data-stu-id="7a1f3-122">Color each cell with unique color</span></span>

```TypeScript
/**
 * This sample demonstrates how to iterate over a selected range and set cell property.
   It colors each cell within the selected range with a random color.
 */
function main(workbook: ExcelScript.Workbook) {

    const syncStart = new Date().getTime();
    // Get selected range
    const range = workbook.getSelectedRange();
    const rows = range.getRowCount();
    const cols = range.getColumnCount();
    console.log("Start");

    // Color each cell with random color.
    for (let row = 0; row < rows; row++) {
        for (let col = 0; col < cols; col++) {
            range
                .getCell(row, col)
                .getFormat()
                .getFill()
                .setColor(`#${Math.random().toString(16).substr(-6)}`);
        }
    }

    console.log("End");
    const syncEnd = new Date().getTime();
    console.log("Completed, took: " + (syncEnd - syncStart) / 1000 + " Sec");
}
```

### <a name="update-range-with-values-using-2d-array"></a><span data-ttu-id="7a1f3-123">使用 2D 数组更新值的范围</span><span class="sxs-lookup"><span data-stu-id="7a1f3-123">Update range with values using 2D array</span></span>

<span data-ttu-id="7a1f3-124">根据 2D 数组值动态计算要更新的范围维度。</span><span class="sxs-lookup"><span data-stu-id="7a1f3-124">Dynamically calculates the range dimension to update based on 2D array values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const currentCell = workbook.getActiveCell();
  let inputRange = computeTargetRange(currentCell, DATA);
  // Set range values.
  console.log(inputRange.getAddress());
  inputRange.setValues(DATA);
  // Call a helper function to place border around the range.
  borderAround(inputRange);
}

/**
 * A helper function that computes the target range given the target range's starting cell and selected range. 
 */
function computeTargetRange(targetCell: ExcelScript.Range, data: string[][]): ExcelScript.Range {
  const targetRange = targetCell.getResizedRange(data.length - 1, data[0].length - 1);
  return targetRange;
}

/**
 * A helper function that places a border around the range.
 */
function borderAround(range: ExcelScript.Range): void {
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.dash);
  return;
}

// Values used for range setup.
const DATA = [
  ['Item', 'Bread', 'Donuts', 'Cookies', 'Cakes', 'Pies'],
  ['Amount', '2', '1.5', '4', '12', '26']
]
```

## <a name="training-videos-range-basics"></a><span data-ttu-id="7a1f3-125">培训视频：范围基础知识</span><span class="sxs-lookup"><span data-stu-id="7a1f3-125">Training videos: Range basics</span></span>

<span data-ttu-id="7a1f3-126">_Range 基础知识_</span><span class="sxs-lookup"><span data-stu-id="7a1f3-126">_Range basics_</span></span>

<span data-ttu-id="7a1f3-127">[![观看有关 Range 基础知识的分步视频](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "有关 Range 基础知识的分步视频")</span><span class="sxs-lookup"><span data-stu-id="7a1f3-127">[![Watch step-by-step video on Range basics](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Step-by-step video on Range basics")</span></span>

<span data-ttu-id="7a1f3-128">_在工作表末尾添加行_</span><span class="sxs-lookup"><span data-stu-id="7a1f3-128">_Add row at the end of worksheet_</span></span>

<span data-ttu-id="7a1f3-129">[![观看分步视频，了解如何在工作表末尾添加行](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "有关如何在工作表末尾添加行的分步视频")</span><span class="sxs-lookup"><span data-stu-id="7a1f3-129">[![Watch step-by-step video on how to add a row at the end of a worksheet](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Step-by-step video on how to add a row at the end of a worksheet")</span></span>

## <a name="methods-that-return-some-range-metadata"></a><span data-ttu-id="7a1f3-130">返回某些区域元数据的方法</span><span class="sxs-lookup"><span data-stu-id="7a1f3-130">Methods that return some range metadata</span></span>

* <span data-ttu-id="7a1f3-131">getAddress () getAddressLocal () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-131">getAddress(), getAddressLocal()</span></span>
* <span data-ttu-id="7a1f3-132">getCellCount () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-132">getCellCount()</span></span>
* <span data-ttu-id="7a1f3-133">getRowCount () getColumnCount () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-133">getRowCount(), getColumnCount()</span></span>

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a><span data-ttu-id="7a1f3-134">返回与给定区域关联的数据/常量的方法</span><span class="sxs-lookup"><span data-stu-id="7a1f3-134">Methods that return data/constants associated with a given range</span></span>

### <a name="returned-as-single-cell-value"></a><span data-ttu-id="7a1f3-135">作为单个单元格值返回</span><span class="sxs-lookup"><span data-stu-id="7a1f3-135">Returned as single cell value</span></span>

* <span data-ttu-id="7a1f3-136">getFormula () getFormulaLocal () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-136">getFormula(), getFormulaLocal()</span></span>
* <span data-ttu-id="7a1f3-137">getFormulaR1C1 () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-137">getFormulaR1C1()</span></span>
* <span data-ttu-id="7a1f3-138">getNumberFormat () getNumberFormatLocal () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-138">getNumberFormat(), getNumberFormatLocal()</span></span>
* <span data-ttu-id="7a1f3-139">getText()</span><span class="sxs-lookup"><span data-stu-id="7a1f3-139">getText()</span></span>
* <span data-ttu-id="7a1f3-140">getValue () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-140">getValue()</span></span>
* <span data-ttu-id="7a1f3-141">getValueType () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-141">getValueType()</span></span>

### <a name="returned-as-2d-arrays-whole-range"></a><span data-ttu-id="7a1f3-142">作为 2D 数组返回 (整个区域) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-142">Returned as 2D arrays (whole range)</span></span>

* <span data-ttu-id="7a1f3-143">getFormulas () getFormulasLocal () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-143">getFormulas(), getFormulasLocal()</span></span>
* <span data-ttu-id="7a1f3-144">getFormulasR1C1 () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-144">getFormulasR1C1()</span></span>
* <span data-ttu-id="7a1f3-145">getNumberFormatCategories () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-145">getNumberFormatCategories()</span></span>
* <span data-ttu-id="7a1f3-146">getNumberFormats () getNumberFormatsLocal () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-146">getNumberFormats(), getNumberFormatsLocal()</span></span>
* <span data-ttu-id="7a1f3-147">getTexts () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-147">getTexts()</span></span>
* <span data-ttu-id="7a1f3-148">getValues () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-148">getValues()</span></span>
* <span data-ttu-id="7a1f3-149">getValueTypes () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-149">getValueTypes()</span></span>
* <span data-ttu-id="7a1f3-150">getHidden () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-150">getHidden()</span></span>
* <span data-ttu-id="7a1f3-151">getIsEntireRow () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-151">getIsEntireRow()</span></span>
* <span data-ttu-id="7a1f3-152">getIsEntireColumn () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-152">getIsEntireColumn()</span></span>

## <a name="methods-that-return-other-range-object"></a><span data-ttu-id="7a1f3-153">返回其他 range 对象的方法</span><span class="sxs-lookup"><span data-stu-id="7a1f3-153">Methods that return other range object</span></span>

* <span data-ttu-id="7a1f3-154">getSurroundingRegion () - 类似于 VBA 中的 CurrentRegion</span><span class="sxs-lookup"><span data-stu-id="7a1f3-154">getSurroundingRegion() -- similar to CurrentRegion in VBA</span></span>
* <span data-ttu-id="7a1f3-155">getCell (行、列) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-155">getCell(row, column)</span></span>
* <span data-ttu-id="7a1f3-156">getColumn (列) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-156">getColumn(column)</span></span>
* <span data-ttu-id="7a1f3-157">getColumnHidden () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-157">getColumnHidden()</span></span>
* <span data-ttu-id="7a1f3-158">getColumnsAfter (count) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-158">getColumnsAfter(count)</span></span>
* <span data-ttu-id="7a1f3-159">getColumnsBefore (count) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-159">getColumnsBefore(count)</span></span>
* <span data-ttu-id="7a1f3-160">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="7a1f3-160">getEntireColumn()</span></span>
* <span data-ttu-id="7a1f3-161">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="7a1f3-161">getEntireRow()</span></span>
* <span data-ttu-id="7a1f3-162">getLastCell () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-162">getLastCell()</span></span>
* <span data-ttu-id="7a1f3-163">getLastColumn () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-163">getLastColumn()</span></span>
* <span data-ttu-id="7a1f3-164">getLastRow () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-164">getLastRow()</span></span>
* <span data-ttu-id="7a1f3-165">getRow (行) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-165">getRow(row)</span></span>
* <span data-ttu-id="7a1f3-166">getRowHidden () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-166">getRowHidden()</span></span>
* <span data-ttu-id="7a1f3-167">getRowsAbove (count) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-167">getRowsAbove(count)</span></span>
* <span data-ttu-id="7a1f3-168">getRowsBelow (count) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-168">getRowsBelow(count)</span></span>

<span data-ttu-id="7a1f3-169">**重要/有趣**</span><span class="sxs-lookup"><span data-stu-id="7a1f3-169">**Important/Interesting**</span></span>

* <span data-ttu-id="7a1f3-170">_workbook_.getSelectedRange () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-170">_workbook_.getSelectedRange()</span></span>
* <span data-ttu-id="7a1f3-171">_workbook_.getActiveCell () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-171">_workbook_.getActiveCell()</span></span>
* <span data-ttu-id="7a1f3-172">getUsedRange (valuesOnly) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-172">getUsedRange(valuesOnly)</span></span>
* <span data-ttu-id="7a1f3-173">getAbsoluteResizedRange (numRows、 numColumns) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-173">getAbsoluteResizedRange(numRows, numColumns)</span></span>
* <span data-ttu-id="7a1f3-174">getOffsetRange (rowOffset、 columnOffset) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-174">getOffsetRange(rowOffset, columnOffset)</span></span>
* <span data-ttu-id="7a1f3-175">getResizedRange (deltaRows、deltaColumns) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-175">getResizedRange(deltaRows, deltaColumns)</span></span>

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a><span data-ttu-id="7a1f3-176">返回与另一个 range 对象相关的 range 对象的方法</span><span class="sxs-lookup"><span data-stu-id="7a1f3-176">Methods that return a range object in relation to another range object</span></span>

* <span data-ttu-id="7a1f3-177">getBoundingRect (anotherRange) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-177">getBoundingRect(anotherRange)</span></span>
* <span data-ttu-id="7a1f3-178">getIntersection (anotherRange) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-178">getIntersection(anotherRange)</span></span>

## <a name="methods-that-return-other-objects-non-range-objects"></a><span data-ttu-id="7a1f3-179">返回非 range 对象 (对象的方法) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-179">Methods that return other objects (non-range objects)</span></span>

* <span data-ttu-id="7a1f3-180">getDirectPrecedents () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-180">getDirectPrecedents()</span></span>
* <span data-ttu-id="7a1f3-181">getWorksheet () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-181">getWorksheet()</span></span>
* <span data-ttu-id="7a1f3-182">getTables (完全包含) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-182">getTables(fullyContained)</span></span>
* <span data-ttu-id="7a1f3-183">getPivotTables (fullyContained) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-183">getPivotTables(fullyContained)</span></span>
* <span data-ttu-id="7a1f3-184">getDataValidation () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-184">getDataValidation()</span></span>
* <span data-ttu-id="7a1f3-185">getPredefinedCellStyle () </span><span class="sxs-lookup"><span data-stu-id="7a1f3-185">getPredefinedCellStyle()</span></span>

## <a name="set-methods"></a><span data-ttu-id="7a1f3-186">Set 方法</span><span class="sxs-lookup"><span data-stu-id="7a1f3-186">Set methods</span></span>

### <a name="singular-cell-set-methods"></a><span data-ttu-id="7a1f3-187">单数单元格集方法</span><span class="sxs-lookup"><span data-stu-id="7a1f3-187">Singular cell set methods</span></span>

* <span data-ttu-id="7a1f3-188">setFormula (公式) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-188">setFormula(formula)</span></span>
* <span data-ttu-id="7a1f3-189">setFormulaLocal (formulaLocal) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-189">setFormulaLocal(formulaLocal)</span></span>
* <span data-ttu-id="7a1f3-190">setFormulaR1C1 (formulaR1C1) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-190">setFormulaR1C1(formulaR1C1)</span></span>
* <span data-ttu-id="7a1f3-191">setNumberFormatLocal (numberFormatLocal) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-191">setNumberFormatLocal(numberFormatLocal)</span></span>
* <span data-ttu-id="7a1f3-192">setValue (值) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-192">setValue(value)</span></span>

### <a name="2d--entire-range-set-methods"></a><span data-ttu-id="7a1f3-193">2D / 整个范围集方法</span><span class="sxs-lookup"><span data-stu-id="7a1f3-193">2D / entire range set methods</span></span>

* <span data-ttu-id="7a1f3-194">setFormulas (公式) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-194">setFormulas(formulas)</span></span>
* <span data-ttu-id="7a1f3-195">setFormulasLocal (formulasLocal) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-195">setFormulasLocal(formulasLocal)</span></span>
* <span data-ttu-id="7a1f3-196">setFormulasR1C1 (formulasR1C1) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-196">setFormulasR1C1(formulasR1C1)</span></span>
* <span data-ttu-id="7a1f3-197">setNumberFormat (numberFormat) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-197">setNumberFormat(numberFormat)</span></span>
* <span data-ttu-id="7a1f3-198">setNumberFormats (numberFormats) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-198">setNumberFormats(numberFormats)</span></span>
* <span data-ttu-id="7a1f3-199">setNumberFormatsLocal (numberFormatsLocal) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-199">setNumberFormatsLocal(numberFormatsLocal)</span></span>
* <span data-ttu-id="7a1f3-200">setValues (值) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-200">setValues(values)</span></span>

## <a name="other-methods"></a><span data-ttu-id="7a1f3-201">其他方法</span><span class="sxs-lookup"><span data-stu-id="7a1f3-201">Other methods</span></span>

* <span data-ttu-id="7a1f3-202">跨 (合并) </span><span class="sxs-lookup"><span data-stu-id="7a1f3-202">merge(across)</span></span>
* <span data-ttu-id="7a1f3-203">unmerge()</span><span class="sxs-lookup"><span data-stu-id="7a1f3-203">unmerge()</span></span>

## <a name="coming-soon"></a><span data-ttu-id="7a1f3-204">即将推出</span><span class="sxs-lookup"><span data-stu-id="7a1f3-204">Coming soon</span></span>

* <span data-ttu-id="7a1f3-205">范围边缘 API</span><span class="sxs-lookup"><span data-stu-id="7a1f3-205">Range edge APIs</span></span>
