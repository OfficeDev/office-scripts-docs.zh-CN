---
title: 编写大型数据集时的性能优化
description: 了解如何在脚本中编写大型数据集时优化Office性能。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: dcbcf156ef624c4c5ce35c44d501286d507d9c40
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232716"
---
# <a name="performance-optimization-when-writing-a-large-dataset"></a><span data-ttu-id="f7a57-103">编写大型数据集时的性能优化</span><span class="sxs-lookup"><span data-stu-id="f7a57-103">Performance optimization when writing a large dataset</span></span>

## <a name="basic-performance-optimization"></a><span data-ttu-id="f7a57-104">基本性能优化</span><span class="sxs-lookup"><span data-stu-id="f7a57-104">Basic performance optimization</span></span>

<span data-ttu-id="f7a57-105">有关脚本Office基础知识，请参阅[入门文章](getting-started.md#basic-performance-considerations)的性能部分。</span><span class="sxs-lookup"><span data-stu-id="f7a57-105">For performance basics in Office Scripts, see the [performance section](getting-started.md#basic-performance-considerations) of the Getting Started article.</span></span>

## <a name="sample-code-optimize-performance-of-a-large-dataset"></a><span data-ttu-id="f7a57-106">示例代码：优化大型数据集的性能</span><span class="sxs-lookup"><span data-stu-id="f7a57-106">Sample code: Optimize performance of a large dataset</span></span>

<span data-ttu-id="f7a57-107">区域 `setValues()` API 允许设置区域的值。</span><span class="sxs-lookup"><span data-stu-id="f7a57-107">The `setValues()` Range API allows setting the values of a range.</span></span> <span data-ttu-id="f7a57-108">此 API 具有数据限制，具体取决于各种因素，如数据大小、网络设置等。为了可靠地更新大量数据，您需要考虑以较小的区块执行数据更新。</span><span class="sxs-lookup"><span data-stu-id="f7a57-108">This API has data limitations depending on various factors such as data size, network settings, etc. In order to reliably update a large range of data, you'll need to think about doing data updates in smaller chunks.</span></span> <span data-ttu-id="f7a57-109">此脚本尝试这样做，并写入区块中的区域行，以便如果需要更新较大区域，可以在较小的部件中完成。</span><span class="sxs-lookup"><span data-stu-id="f7a57-109">This script attempts to do this and writes rows of a range in chunks so that if a large range needs to be updated, it can be done in smaller parts.</span></span> <span data-ttu-id="f7a57-110">**警告**：尚未跨各种大小进行测试，因此如果要在脚本中使用它，请注意这一点。</span><span class="sxs-lookup"><span data-stu-id="f7a57-110">**Warning**: It has not been tested across various sizes so be aware of that if you want to use this in your script.</span></span> <span data-ttu-id="f7a57-111">由于我们有机会进行测试，我们将更新有关它在各种数据大小下如何执行的结果。</span><span class="sxs-lookup"><span data-stu-id="f7a57-111">As we have opportunity to test, we'll update with findings around how it performs for various data sizes.</span></span>

<span data-ttu-id="f7a57-112">此脚本选择每个区块 1K 个单元格，但你可以重写以测试它是如何工作的。</span><span class="sxs-lookup"><span data-stu-id="f7a57-112">This script selects 1K cells per chunk but you can override to test out how it works for you.</span></span> <span data-ttu-id="f7a57-113">它使用 6 列数据更新 100k 行。</span><span class="sxs-lookup"><span data-stu-id="f7a57-113">It updates 100k rows with 6 columns of data.</span></span> <span data-ttu-id="f7a57-114">在空白工作表上运行此代码以检查。</span><span class="sxs-lookup"><span data-stu-id="f7a57-114">Run this on a blank sheet to examine.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  let data: (string | number | boolean)[][] = [];
  // Number of rows in the random data (x 6 columns).
  const sampleRows = 100000;

  console.log(`Generating data...`)
  // Dynamically generate some random data for testing purpose. 
  for (let i = 0; i < sampleRows; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInChunks(sheet.getRange("B2"), data);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`)
  }

  return;
}

function updateRangeInChunks(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInChunk: number = 10000
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per chunk setting: ${cellsInChunk}`);
  if (!values) {
    console.log(`Invalid input values to update.`);
    return false;
  }
  if (values.length === 0 || values[0].length === 0) {
    console.log(`Empty data -- nothing to update.`);
    return true;
  }
  const totalCells = values.length * values[0].length;

  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInChunk) {
    console.log(`No need to chunk -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  const rowsPerChunk = Math.floor(cellsInChunk / values[0].length);
  console.log("Rows per chunk: " + rowsPerChunk);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let chunkCount = 0;

  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerChunk) {
      chunkCount++;
      console.log(`Calling update next chunk function. Chunk#: ${chunkCount}`);
      updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
      rowCount = 0;
      totalRowsUpdated += rowsPerChunk;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);

    }
  }
  console.log(`Updating remaining rows -- last chunk: ${rowCount}`)
  if (rowCount > 0) {
    updateNextChunk(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInChunk).toFixed(8)} seconds per ${cellsInChunk} cells-chunk.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */

function updateNextChunk(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerChunk: number,
  totalRowsUpdated: number
) {

  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerChunk - 1, data[0].length - 1);
  console.log(`Updating chunk at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerChunk);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the chunk range: ${JSON.stringify(e)}`;
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

## <a name="training-video-optimize-performance-when-writing-a-large-dataset"></a><span data-ttu-id="f7a57-115">培训视频：在编写大型数据集时优化性能</span><span class="sxs-lookup"><span data-stu-id="f7a57-115">Training video: Optimize performance when writing a large dataset</span></span>

<span data-ttu-id="f7a57-116">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/BP9Kp0Ltj7U)。</span><span class="sxs-lookup"><span data-stu-id="f7a57-116">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/BP9Kp0Ltj7U).</span></span>
