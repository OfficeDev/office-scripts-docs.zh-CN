---
title: 对工作表中的空行计数
description: 了解如何使用 Office 脚本检测工作表中是否有空行而不是数据，然后报告要用于数据流的空白Power Automate计数。
ms.date: 05/04/2021
localization_priority: Normal
ms.openlocfilehash: 73fe0f995ee6ccaa1328b68983f0ec6887d96a09
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074576"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="c2d53-103">对工作表中的空行计数</span><span class="sxs-lookup"><span data-stu-id="c2d53-103">Count blank rows on sheets</span></span>

<span data-ttu-id="c2d53-104">此项目包括两个脚本：</span><span class="sxs-lookup"><span data-stu-id="c2d53-104">This project includes two scripts:</span></span>

* <span data-ttu-id="c2d53-105">[对给定工作表上的空](#sample-code-count-blank-rows-on-a-given-sheet)行进行计数：遍历给定工作表上的已用区域并返回空行数。</span><span class="sxs-lookup"><span data-stu-id="c2d53-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="c2d53-106">[统计所有工作表上的](#sample-code-count-blank-rows-on-all-sheets)空行数：遍历所有工作表上的已用区域并返回空行数。</span><span class="sxs-lookup"><span data-stu-id="c2d53-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="c2d53-107">对于我们的脚本，空行是没有任何数据的任何行。</span><span class="sxs-lookup"><span data-stu-id="c2d53-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="c2d53-108">行可以具有格式。</span><span class="sxs-lookup"><span data-stu-id="c2d53-108">The row can have formatting.</span></span>

<span data-ttu-id="c2d53-109">_此工作表返回 4 个空行的计数_</span><span class="sxs-lookup"><span data-stu-id="c2d53-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="显示包含空白行的数据的工作表。":::

<span data-ttu-id="c2d53-111">_此工作表返回 0 个空 (所有行都有一些数据)_</span><span class="sxs-lookup"><span data-stu-id="c2d53-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="显示不带空白行的数据的工作表。":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="c2d53-113">示例代码：对给定工作表上的空白行计数</span><span class="sxs-lookup"><span data-stu-id="c2d53-113">Sample code: Count blank rows on a given sheet</span></span>

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

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="c2d53-114">示例代码：统计所有工作表上的空行数</span><span class="sxs-lookup"><span data-stu-id="c2d53-114">Sample code: Count blank rows on all sheets</span></span>

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

## <a name="use-with-power-automate"></a><span data-ttu-id="c2d53-115">与 Power Automate</span><span class="sxs-lookup"><span data-stu-id="c2d53-115">Use with Power Automate</span></span>

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="显示Power Automate运行脚本的一个Office流。":::
