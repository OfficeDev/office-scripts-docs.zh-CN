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
# <a name="range-basics"></a>Range 基础知识

`Range` 是 Office Scripts Excel 对象模型中的基础对象。 [区域 API](/javascript/api/office-scripts/excelscript/excelscript.range) 允许访问网格上可用的数据和格式，并链接 Excel 内的其他关键对象，如工作表、表、图表等。

区域使用其地址（如"A1：B4"）或已命名项（它是一组给定单元格的命名键）进行标识。 在 Excel 对象模型中，单元格和单元格组都称为 _range_。 `Range` 可以包含单元格级属性（如单元格内的数据），还可以包含单元格和单元格级属性（如格式、边框等）。

`Range` 还可通过用户选择（至少包含一个单元格）获取。 与区域交互时，必须明确这些单元格和范围关系。

以下是 getter、setter 和其他在脚本中最常用的有用方法的核心集。 这是 API 旅程的一个很好起点。 以下各节对方法进行分组，并有助于在开始解锁对象的 API 时 `Range` 构建一个精神模型。

## <a name="example-scripts"></a>示例脚本

* [基本读写](#basic-read-and-write)
* [在工作表末尾添加行](#add-row-at-the-end-of-worksheet)
* [清除列筛选器](clear-table-filter-for-active-cell.md)
* [使用唯一的颜色为每个单元格设置颜色](#color-each-cell-with-unique-color)
* [使用二维数组更新二维 (二维) 区域](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a>基本读写

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

### <a name="add-row-at-the-end-of-worksheet"></a>在工作表末尾添加行

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

### <a name="color-each-cell-with-unique-color"></a>使用唯一的颜色为每个单元格设置颜色

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

### <a name="update-range-with-values-using-2d-array"></a>使用 2D 数组更新值的范围

根据 2D 数组值动态计算要更新的范围维度。

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

## <a name="training-videos-range-basics"></a>培训视频：范围基础知识

_Range 基础知识_

[![观看有关 Range 基础知识的分步视频](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "有关 Range 基础知识的分步视频")

_在工作表末尾添加行_

[![观看分步视频，了解如何在工作表末尾添加行](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "有关如何在工作表末尾添加行的分步视频")

## <a name="methods-that-return-some-range-metadata"></a>返回某些区域元数据的方法

* getAddress () getAddressLocal () 
* getCellCount () 
* getRowCount () getColumnCount () 

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a>返回与给定区域关联的数据/常量的方法

### <a name="returned-as-single-cell-value"></a>作为单个单元格值返回

* getFormula () getFormulaLocal () 
* getFormulaR1C1 () 
* getNumberFormat () getNumberFormatLocal () 
* getText()
* getValue () 
* getValueType () 

### <a name="returned-as-2d-arrays-whole-range"></a>作为 2D 数组返回 (整个区域) 

* getFormulas () getFormulasLocal () 
* getFormulasR1C1 () 
* getNumberFormatCategories () 
* getNumberFormats () getNumberFormatsLocal () 
* getTexts () 
* getValues () 
* getValueTypes () 
* getHidden () 
* getIsEntireRow () 
* getIsEntireColumn () 

## <a name="methods-that-return-other-range-object"></a>返回其他 range 对象的方法

* getSurroundingRegion () - 类似于 VBA 中的 CurrentRegion
* getCell (行、列) 
* getColumn (列) 
* getColumnHidden () 
* getColumnsAfter (count) 
* getColumnsBefore (count) 
* getEntireColumn()
* getEntireRow()
* getLastCell () 
* getLastColumn () 
* getLastRow () 
* getRow (行) 
* getRowHidden () 
* getRowsAbove (count) 
* getRowsBelow (count) 

**重要/有趣**

* _workbook_.getSelectedRange () 
* _workbook_.getActiveCell () 
* getUsedRange (valuesOnly) 
* getAbsoluteResizedRange (numRows、 numColumns) 
* getOffsetRange (rowOffset、 columnOffset) 
* getResizedRange (deltaRows、deltaColumns) 

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a>返回与另一个 range 对象相关的 range 对象的方法

* getBoundingRect (anotherRange) 
* getIntersection (anotherRange) 

## <a name="methods-that-return-other-objects-non-range-objects"></a>返回非 range 对象 (对象的方法) 

* getDirectPrecedents () 
* getWorksheet () 
* getTables (完全包含) 
* getPivotTables (fullyContained) 
* getDataValidation () 
* getPredefinedCellStyle () 

## <a name="set-methods"></a>Set 方法

### <a name="singular-cell-set-methods"></a>单数单元格集方法

* setFormula (公式) 
* setFormulaLocal (formulaLocal) 
* setFormulaR1C1 (formulaR1C1) 
* setNumberFormatLocal (numberFormatLocal) 
* setValue (值) 

### <a name="2d--entire-range-set-methods"></a>2D / 整个范围集方法

* setFormulas (公式) 
* setFormulasLocal (formulasLocal) 
* setFormulasR1C1 (formulasR1C1) 
* setNumberFormat (numberFormat) 
* setNumberFormats (numberFormats) 
* setNumberFormatsLocal (numberFormatsLocal) 
* setValues (值) 

## <a name="other-methods"></a>其他方法

* 跨 (合并) 
* unmerge()

## <a name="coming-soon"></a>即将推出

* 范围边缘 API
