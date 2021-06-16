---
title: 提高脚本Office性能
description: 通过了解工作簿和脚本之间的通信Excel创建更快的脚本。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: a5bd879625b9c3bac0caa621dde312f7c961dd5c
ms.sourcegitcommit: 2aaf7dc527cb6c9f1206550b2c5745280503b2a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/16/2021
ms.locfileid: "52957698"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="b48f2-103">提高脚本Office性能</span><span class="sxs-lookup"><span data-stu-id="b48f2-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="b48f2-104">使用脚本Office自动执行一系列任务，以节省时间。</span><span class="sxs-lookup"><span data-stu-id="b48f2-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="b48f2-105">较慢的脚本可能会感觉它无法加快工作流的速度。</span><span class="sxs-lookup"><span data-stu-id="b48f2-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="b48f2-106">大多数情况下，脚本将完全正常并如期运行。</span><span class="sxs-lookup"><span data-stu-id="b48f2-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="b48f2-107">但是，有一些可避免的场景可能会影响性能。</span><span class="sxs-lookup"><span data-stu-id="b48f2-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="b48f2-108">脚本运行缓慢的最常见原因是与工作簿的通信过多。</span><span class="sxs-lookup"><span data-stu-id="b48f2-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="b48f2-109">当工作簿存在于云中时，脚本将在本地计算机上运行。</span><span class="sxs-lookup"><span data-stu-id="b48f2-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="b48f2-110">在某些时候，脚本会将其本地数据与工作簿的本地数据同步。</span><span class="sxs-lookup"><span data-stu-id="b48f2-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="b48f2-111">这意味着，当 (同步时，) 写入操作（如) ） `workbook.addWorksheet()` 都只应用于工作簿。</span><span class="sxs-lookup"><span data-stu-id="b48f2-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="b48f2-112">同样，任何读取操作 (，) 这些时间仅从脚本的 `myRange.getValues()` 工作簿获取数据。</span><span class="sxs-lookup"><span data-stu-id="b48f2-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="b48f2-113">在任一情况下，脚本都先提取信息，然后再处理数据。</span><span class="sxs-lookup"><span data-stu-id="b48f2-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="b48f2-114">例如，以下代码将准确记录已用区域中的行数。</span><span class="sxs-lookup"><span data-stu-id="b48f2-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="b48f2-115">Office脚本 API 确保工作簿或脚本中任何数据都准确且在必要时是最新的。</span><span class="sxs-lookup"><span data-stu-id="b48f2-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="b48f2-116">无需担心这些同步，脚本就能够正常运行。</span><span class="sxs-lookup"><span data-stu-id="b48f2-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="b48f2-117">但是，了解此脚本到云通信可以帮助您避免不必要的网络调用。</span><span class="sxs-lookup"><span data-stu-id="b48f2-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="b48f2-118">性能优化</span><span class="sxs-lookup"><span data-stu-id="b48f2-118">Performance optimizations</span></span>

<span data-ttu-id="b48f2-119">你可以应用简单的技术来帮助减少与云的通信。</span><span class="sxs-lookup"><span data-stu-id="b48f2-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="b48f2-120">以下模式有助于加快脚本速度。</span><span class="sxs-lookup"><span data-stu-id="b48f2-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="b48f2-121">读取工作簿数据一次，而不是在循环中重复读取。</span><span class="sxs-lookup"><span data-stu-id="b48f2-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="b48f2-122">删除不必要的 `console.log` 语句。</span><span class="sxs-lookup"><span data-stu-id="b48f2-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="b48f2-123">避免使用 try/catch 块。</span><span class="sxs-lookup"><span data-stu-id="b48f2-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="b48f2-124">在循环之外读取工作簿数据</span><span class="sxs-lookup"><span data-stu-id="b48f2-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="b48f2-125">从工作簿获取数据的任何方法都可以触发网络调用。</span><span class="sxs-lookup"><span data-stu-id="b48f2-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="b48f2-126">应尽量在本地保存数据，而不是重复进行同一调用。</span><span class="sxs-lookup"><span data-stu-id="b48f2-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="b48f2-127">在处理循环时尤其如此。</span><span class="sxs-lookup"><span data-stu-id="b48f2-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="b48f2-128">请考虑使用脚本获取工作表的已用区域中的负数计数。</span><span class="sxs-lookup"><span data-stu-id="b48f2-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="b48f2-129">脚本需要对已用区域内每个单元格进行访问。</span><span class="sxs-lookup"><span data-stu-id="b48f2-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="b48f2-130">为此，它需要范围、行数和列数。</span><span class="sxs-lookup"><span data-stu-id="b48f2-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="b48f2-131">您应该在启动循环之前，将那些变量存储为本地变量。</span><span class="sxs-lookup"><span data-stu-id="b48f2-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="b48f2-132">否则，循环的每个迭代将强制返回到工作簿。</span><span class="sxs-lookup"><span data-stu-id="b48f2-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> <span data-ttu-id="b48f2-133">作为实验，请尝试在 `usedRangeValues` 循环中用 替换 `usedRange.getValues()` 。</span><span class="sxs-lookup"><span data-stu-id="b48f2-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="b48f2-134">您可能会注意到，在处理较大范围时，脚本需要更长的时间运行。</span><span class="sxs-lookup"><span data-stu-id="b48f2-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a><span data-ttu-id="b48f2-135">避免 `try...catch` 在循环或周围循环中使用块</span><span class="sxs-lookup"><span data-stu-id="b48f2-135">Avoid using `try...catch` blocks in or surrounding loops</span></span>

<span data-ttu-id="b48f2-136">我们不建议在循环 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) 或周围循环中使用语句。</span><span class="sxs-lookup"><span data-stu-id="b48f2-136">We don't recommend using [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statements either in loops or surrounding loops.</span></span> <span data-ttu-id="b48f2-137">这是您避免在循环中读取数据的相同原因：每次迭代都强制脚本与工作簿同步，以确保未引发任何错误。</span><span class="sxs-lookup"><span data-stu-id="b48f2-137">This is for the same reason you should avoid reading data in a loop: each iteration forces the script to synchronize with the workbook to make sure no error has been thrown.</span></span> <span data-ttu-id="b48f2-138">通过检查从工作簿返回的对象，可以避免大多数错误。</span><span class="sxs-lookup"><span data-stu-id="b48f2-138">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="b48f2-139">例如，以下脚本在尝试添加行之前检查工作簿返回的表是否存在。</span><span class="sxs-lookup"><span data-stu-id="b48f2-139">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="b48f2-140">删除不必要的 `console.log` 语句</span><span class="sxs-lookup"><span data-stu-id="b48f2-140">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="b48f2-141">控制台日志记录是调试 [脚本的重要工具](../testing/troubleshooting.md)。</span><span class="sxs-lookup"><span data-stu-id="b48f2-141">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="b48f2-142">但是，它会强制脚本与工作簿同步，以确保记录的信息是最新的。</span><span class="sxs-lookup"><span data-stu-id="b48f2-142">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="b48f2-143">请考虑在共享脚本 (不必要的日志记录语句，例如) 测试日志的日志记录语句。</span><span class="sxs-lookup"><span data-stu-id="b48f2-143">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="b48f2-144">这通常不会导致明显的性能问题，除非语句 `console.log()` 位于循环中。</span><span class="sxs-lookup"><span data-stu-id="b48f2-144">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

## <a name="case-by-case-help"></a><span data-ttu-id="b48f2-145">按案例帮助</span><span class="sxs-lookup"><span data-stu-id="b48f2-145">Case-by-case help</span></span>

<span data-ttu-id="b48f2-146">随着 Office 脚本平台的扩展以使用[Power Automate、](https://flow.microsoft.com/)自适应卡片和其他跨产品[](/adaptive-cards)功能，脚本工作簿通信的细节变得更加复杂。</span><span class="sxs-lookup"><span data-stu-id="b48f2-146">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="b48f2-147">如果您需要有关加快脚本运行速度的帮助，请通过 [Microsoft Q&A 联系](/answers/topics/office-scripts-excel-dev.html)。</span><span class="sxs-lookup"><span data-stu-id="b48f2-147">If you need help making your script run faster, please reach out through [Microsoft Q&A](/answers/topics/office-scripts-excel-dev.html).</span></span> <span data-ttu-id="b48f2-148">请务必使用"office-scripts-dev"标记你的问题，以便专家可以找到它并提供帮助。</span><span class="sxs-lookup"><span data-stu-id="b48f2-148">Be sure to tag your question with "office-scripts-dev" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="b48f2-149">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b48f2-149">See also</span></span>

- [<span data-ttu-id="b48f2-150">Excel 网页版中 Office 脚本的脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="b48f2-150">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="b48f2-151">MDN Web 文档：循环和迭代</span><span class="sxs-lookup"><span data-stu-id="b48f2-151">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
