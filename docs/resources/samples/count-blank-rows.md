---
title: 对工作表中的空行计数
description: 了解如何使用 Office 脚本检测工作表中是否有空行而不是数据，然后报告要用于数据流的空白Power Automate计数。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: e5b60779d2ca2de5f4cf4e03ddd6ff7372515ad6
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313804"
---
# <a name="count-blank-rows-on-sheets"></a>对工作表中的空行计数

此项目包括两个脚本：

* [对给定工作表上的空](#sample-code-count-blank-rows-on-a-given-sheet)行进行计数：遍历给定工作表上的已用区域并返回空行数。
* [统计所有工作表上的](#sample-code-count-blank-rows-on-all-sheets)空行数：遍历所有工作表上的已用区域并返回空行数。

> [!NOTE]
> 对于我们的脚本，空行是没有任何数据的任何行。 行可以具有格式。

_此工作表返回 4 个空行的计数_

:::image type="content" source="../../images/blank-rows.png" alt-text="显示包含空白行的数据的工作表。":::

_此工作表返回 0 个空 (所有行都有一些数据)_

:::image type="content" source="../../images/no-blank-rows.png" alt-text="显示不带空白行的数据的工作表。":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a>示例代码：对给定工作表上的空白行计数

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Get the worksheet named "Sheet1".
  const sheet = workbook.getWorksheet('Sheet1'); 
  
  // Get the entire data range.
  const range = sheet.getUsedRange(true);

  // If the used range is empty, end the script.
  if (!range) {
    console.log(`No data on this sheet.`);
    return;
  }
  
  // Log the address of the used range.
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
    
  // Look through the values in the range for blank rows.
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let emptyRow = true;
    
    // Look at every cell in the row for one with a value.
    for (let cell of row) {
      if (cell.toString().length > 0) {
        emptyRow = false
      }
    }

    // If no cell had a value, the row is empty.
    if (emptyRow) {
      emptyRows++;
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a>示例代码：统计所有工作表上的空行数

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Loop through every worksheet in the workbook.
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) {     
    // Get the entire data range.
    const range = sheet.getUsedRange(true);
  
    // If the used range is empty, skip to the next worksheet.
    if (!range) {
      console.log(`No data on this sheet.`);
      continue;
    }
    
    // Log the address of the used range.
    console.log(`Used range for the worksheet: ${range.getAddress()}`);
      
    // Look through the values in the range for blank rows.
    const values = range.getValues();
    for (let row of values) {
      let emptyRow = true;
      
      // Look at every cell in the row for one with a value.
      for (let cell of row) {
        if (cell.toString().length > 0) {
          emptyRow = false
        }
      }
  
      // If no cell had a value, the row is empty.
      if (emptyRow) {
        emptyRows++;
      }
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```
