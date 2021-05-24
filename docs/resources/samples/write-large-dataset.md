---
title: 编写大型数据集
description: 了解如何在脚本中将大型数据集拆分为Office操作。
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 06abb58c61c18620d638ab3eb61ea68398bf20aa
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545621"
---
# <a name="write-a-large-dataset"></a><span data-ttu-id="a11f2-103">编写大型数据集</span><span class="sxs-lookup"><span data-stu-id="a11f2-103">Write a large dataset</span></span>

<span data-ttu-id="a11f2-104">`Range.setValues()`API 将数据放在一个范围中。</span><span class="sxs-lookup"><span data-stu-id="a11f2-104">The `Range.setValues()` API puts data in a range.</span></span> <span data-ttu-id="a11f2-105">此 API 有一些限制，具体取决于各种因素，例如数据大小和网络设置。</span><span class="sxs-lookup"><span data-stu-id="a11f2-105">This API has limitations depending on various factors, such as data size and network settings.</span></span> <span data-ttu-id="a11f2-106">这意味着，如果您尝试将大量信息作为单个操作写入工作簿，则需要以较小的批次写入数据，以便可靠地更新 [较大范围](../../testing/platform-limits.md)。</span><span class="sxs-lookup"><span data-stu-id="a11f2-106">This means that if you attempt to write a massive amount of information to a workbook as a single operation, you'll need to write the data in smaller batches in order to reliably update a [large range](../../testing/platform-limits.md).</span></span>

<span data-ttu-id="a11f2-107">有关脚本Office基础知识，请阅读提高脚本[Office性能](../../develop/web-client-performance.md)。</span><span class="sxs-lookup"><span data-stu-id="a11f2-107">For performance basics in Office Scripts, please read [Improve the performance of your Office Scripts](../../develop/web-client-performance.md).</span></span>

## <a name="sample-code-write-a-large-dataset"></a><span data-ttu-id="a11f2-108">示例代码：编写大型数据集</span><span class="sxs-lookup"><span data-stu-id="a11f2-108">Sample code: Write a large dataset</span></span>

<span data-ttu-id="a11f2-109">此脚本以较小的部分写入区域行。</span><span class="sxs-lookup"><span data-stu-id="a11f2-109">This script writes rows of a range in smaller parts.</span></span> <span data-ttu-id="a11f2-110">它选择一次写入 1000 个单元格。</span><span class="sxs-lookup"><span data-stu-id="a11f2-110">It selects 1000 cells to write at a time.</span></span> <span data-ttu-id="a11f2-111">在空白工作表上运行脚本以查看更新批处理的运行情况。</span><span class="sxs-lookup"><span data-stu-id="a11f2-111">Run the script on a blank worksheet to see the update batches in action.</span></span> <span data-ttu-id="a11f2-112">控制台输出进一步深入了解发生了什么。</span><span class="sxs-lookup"><span data-stu-id="a11f2-112">The console output gives further insight into what's happening.</span></span>

> [!NOTE]
> <span data-ttu-id="a11f2-113">可以通过更改 的值来更改正在写入的总行数 `SAMPLE_ROWS` 。</span><span class="sxs-lookup"><span data-stu-id="a11f2-113">You can change the number of total rows being written by changing the value of `SAMPLE_ROWS`.</span></span> <span data-ttu-id="a11f2-114">可以通过更改 的值，将要写入的单元格数更改为单个操作 `CELLS_IN_BATCH` 。</span><span class="sxs-lookup"><span data-stu-id="a11f2-114">You can change the number of cells to write as a single action by changing the value of `CELLS_IN_BATCH`.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const SAMPLE_ROWS = 100000;
  const CELLS_IN_BATCH = 10000;

  // Get the current worksheet.
  const sheet = workbook.getActiveWorksheet();

  console.log(`Generating data...`)
  let data: (string | number | boolean)[][] = [];
  // Generate six columns of random data per row. 
  for (let i = 0; i < SAMPLE_ROWS; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInBatches(sheet.getRange("B2"), data, CELLS_IN_BATCH);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`);
  }
}

function updateRangeInBatches(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInBatch: number
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per batch setting: ${cellsInBatch}`);

  // Determine the total number of cells to write.
  const totalCells = values.length * values[0].length;
  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInBatch) {
    console.log(`No need to batch -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  // Determine how many rows to write at once.
  const rowsPerBatch = Math.floor(cellsInBatch / values[0].length);
  console.log("Rows per batch: " + rowsPerBatch);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let batchCount = 0;

  // Write each batch of rows.
  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerBatch) {
      batchCount++;
      console.log(`Calling update next batch function. Batch#: ${batchCount}`);
      updateNextBatch(startCell, values, rowsPerBatch, totalRowsUpdated);

      // Write a completion percentage to help the user understand the progress.
      rowCount = 0;
      totalRowsUpdated += rowsPerBatch;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);
    }
  }
  
  console.log(`Updating remaining rows -- last batch: ${rowCount}`)
  if (rowCount > 0) {
    updateNextBatch(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInBatch).toFixed(8)} seconds per ${cellsInBatch} cells-batch.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */
function updateNextBatch(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerBatch: number,
  totalRowsUpdated: number
) {
  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerBatch - 1, data[0].length - 1);
  console.log(`Updating batch at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerBatch);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the batch range: ${JSON.stringify(e)}`;
  }
  return;
}

/**
 * A helper function that computes the target range given the target range's starting cell
 * and selected range and updates the values.
 */
function updateTargetRange(
  targetCell: ExcelScript.Range,
  values: (string | boolean | number)[][]
) {
  const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
  console.log(`Updating the range: ${targetRange.getAddress()}`);
  try {
    targetRange.setValues(values);
  } catch (e) {
    throw `Error while updating the whole range: ${JSON.stringify(e)}`;
  }
  return;
}

// Credit: https://www.codegrepper.com/code-examples/javascript/random+text+generator+javascript
function getRandomString(length: number): string {
  var randomChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var result = '';
  for (var i = 0; i < length; i++) {
    result += randomChars.charAt(Math.floor(Math.random() * randomChars.length));
  }
  return result;
}
```

## <a name="training-video-write-a-large-dataset"></a><span data-ttu-id="a11f2-115">培训视频：编写大型数据集</span><span class="sxs-lookup"><span data-stu-id="a11f2-115">Training video: Write a large dataset</span></span>

<span data-ttu-id="a11f2-116">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/BP9Kp0Ltj7U)。</span><span class="sxs-lookup"><span data-stu-id="a11f2-116">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/BP9Kp0Ltj7U).</span></span>
