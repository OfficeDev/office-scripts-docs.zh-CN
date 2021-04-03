---
title: 对工作表中的空行计数
description: 了解如何使用 Office 脚本检测工作表中是否有空行而不是数据，然后报告要用于 Power Automate 流的空白行数。
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: 1f52b9c4d538d5d3e64dc61dae3e27d046b56862
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571136"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="2a918-103">对工作表中的空行计数</span><span class="sxs-lookup"><span data-stu-id="2a918-103">Count blank rows on sheets</span></span>

<span data-ttu-id="2a918-104">此项目包括两个脚本：</span><span class="sxs-lookup"><span data-stu-id="2a918-104">This project includes two scripts:</span></span>

* <span data-ttu-id="2a918-105">[对给定工作表上的空](#sample-code-count-blank-rows-on-a-given-sheet)行进行计数：遍历给定工作表上的已用区域并返回空行数。</span><span class="sxs-lookup"><span data-stu-id="2a918-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="2a918-106">[统计所有工作表上的](#sample-code-count-blank-rows-on-all-sheets)空行数：遍历所有工作表上的已用区域并返回空行数。</span><span class="sxs-lookup"><span data-stu-id="2a918-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="2a918-107">对于我们的脚本，空行是没有任何数据的任何行。</span><span class="sxs-lookup"><span data-stu-id="2a918-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="2a918-108">行可以具有格式。</span><span class="sxs-lookup"><span data-stu-id="2a918-108">The row can have formatting.</span></span>

<span data-ttu-id="2a918-109">_此工作表返回 4 个空行的计数_</span><span class="sxs-lookup"><span data-stu-id="2a918-109">_This sheet returns count of 4 blank rows_</span></span>

![具有空行的数据](../../images/blank-rows.png)

<span data-ttu-id="2a918-111">_此工作表返回 0 个空 (所有行都有一些数据)_</span><span class="sxs-lookup"><span data-stu-id="2a918-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

![不含空白行的数据](../../images/no-blank-rows.png)

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="2a918-113">示例代码：对给定工作表上的空白行计数</span><span class="sxs-lookup"><span data-stu-id="2a918-113">Sample code: Count blank rows on a given sheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheet = workbook.getWorksheet('Sheet1'); 
  // Getting the active worksheet is not suitable for a script used by Power Automate.
  // const sheet = workbook.getActiveWorksheet();
  
  const range = sheet.getUsedRange(true); // Get value only.
  if (!range) {
    console.log(`No data on this sheet. `);
    return;
  }
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let len = 0; 
    for (let cell of row) {
      len = len + cell.toString().length;
    }
    if (len === 0) { 
      emptyRows++;
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="2a918-114">示例代码：统计所有工作表上的空行数</span><span class="sxs-lookup"><span data-stu-id="2a918-114">Sample code: Count blank rows on all sheets</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) { 
    const range = sheet.getUsedRange(true); // Get value only.
    if (!range) {
      console.log(`No data on this sheet. `);
      continue;
    }
    console.log(`Used range for the worksheet ${sheet.getName()}: ${range.getAddress()}`);
    const values = range.getValues();

    for (let row of values) {
      let len = 0;
      for (let cell of row) {
        len = len + cell.toString().length;
      }
      if (len === 0) {
        emptyRows++;
      }
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="use-with-power-automate"></a><span data-ttu-id="2a918-115">与 Power Automate 一同使用</span><span class="sxs-lookup"><span data-stu-id="2a918-115">Use with Power Automate</span></span>

![显示如何在 Power Automate 中设置的屏幕截图](../../images/use-in-power-automate.png)
