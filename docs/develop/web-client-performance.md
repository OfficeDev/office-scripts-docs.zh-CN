---
title: 提高 Office 脚本的性能
description: 通过了解 Excel 工作簿和脚本之间的通信来创建更快的脚本。
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878749"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="9fd4f-103">提高 Office 脚本的性能</span><span class="sxs-lookup"><span data-stu-id="9fd4f-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="9fd4f-104">Office 脚本的用途是自动化通常执行的一系列任务以节省时间。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="9fd4f-105">较慢的脚本可能感觉不会加快工作流的速度。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="9fd4f-106">大多数情况下，您的脚本完全正常，并按预期运行。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="9fd4f-107">但是，有几个可能会影响性能的 avoidable 方案。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="9fd4f-108">速度较慢的脚本的最常见原因是与工作簿之间的通信过多。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="9fd4f-109">您的脚本在本地计算机上运行，而该工作簿存在于云中。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="9fd4f-110">在某些情况下，您的脚本会将其本地数据与工作簿的同步。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="9fd4f-111">这意味着， `workbook.addWorksheet()` 只有在发生这种幕后同步时，才会将任何写入操作 (例如，) 仅应用于工作簿。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="9fd4f-112">同样，任何读操作 (例如， `myRange.getValues()`) 仅在这些时间从脚本的工作簿中获取数据。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="9fd4f-113">在这两种情况下，脚本都会在对数据进行操作之前提取信息。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="9fd4f-114">例如，以下代码将准确记录所用区域中的行数。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="9fd4f-115">Office 脚本 Api 可确保工作簿或脚本中的任何数据在必要时都是准确且最新的。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="9fd4f-116">您无需担心这些同步即可正确运行脚本。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="9fd4f-117">但是，对此脚本到云通信的感知可帮助您避免不需要的网络调用。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="9fd4f-118">性能优化</span><span class="sxs-lookup"><span data-stu-id="9fd4f-118">Performance optimizations</span></span>

<span data-ttu-id="9fd4f-119">您可以应用简单的技术来帮助减少与云的通信。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="9fd4f-120">下面的模式可帮助您提高脚本速度。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="9fd4f-121">读取一次工作簿数据，而不是重复执行循环。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="9fd4f-122">删除不必要 `console.log` 的语句。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="9fd4f-123">避免使用 try/catch 块。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="9fd4f-124">读取循环外部的工作簿数据</span><span class="sxs-lookup"><span data-stu-id="9fd4f-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="9fd4f-125">从工作簿中获取数据的任何方法都可以触发网络调用。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="9fd4f-126">应尽可能在本地保存数据，而不是反复进行相同的调用。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="9fd4f-127">在处理循环时尤其如此。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="9fd4f-128">考虑使用脚本来获取工作表的所用区域中的负数计数。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="9fd4f-129">脚本需要循环访问所使用区域中的每个单元格。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="9fd4f-130">若要执行此操作，它需要范围、行数和列数。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="9fd4f-131">在开始循环之前，应将这些变量存储为局部变量。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="9fd4f-132">否则，循环的每个迭代都将强制返回工作簿。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="9fd4f-133">作为实验，请尝试 `usedRangeValues` 将循环中的替换为 `usedRange.getValues()` 。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="9fd4f-134">您可能会注意到，在处理大型区域时脚本运行时间要长得多。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="9fd4f-135">删除不必要的 `console.log` 语句</span><span class="sxs-lookup"><span data-stu-id="9fd4f-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="9fd4f-136">控制台日志记录是 [调试脚本](../testing/troubleshooting.md)的重要工具。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="9fd4f-137">但是，它确实强制脚本与工作簿同步，以确保记录的信息是最新的。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="9fd4f-138">请考虑删除不必要的日志记录语句 (如用于在共享脚本之前测试) 的记录语句。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="9fd4f-139">除非语句在循环中，否则通常不会引起显著的性能问题 `console.log()` 。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="9fd4f-140">避免使用 try/catch 块</span><span class="sxs-lookup"><span data-stu-id="9fd4f-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="9fd4f-141">建议不要将[ `try` / `catch` 块](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch)用作脚本的预期控制流的一部分。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="9fd4f-142">通过检查从工作簿返回的对象，可以避免大多数错误。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="9fd4f-143">例如，下面的脚本在尝试添加行之前检查工作簿返回的表是否存在。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

## <a name="case-by-case-help"></a><span data-ttu-id="9fd4f-144">按大小写帮助</span><span class="sxs-lookup"><span data-stu-id="9fd4f-144">Case-by-case help</span></span>

<span data-ttu-id="9fd4f-145">随着 Office 脚本平台扩展以配合使用 [电源自动化](https://flow.microsoft.com/)、 [自适应卡](https://docs.microsoft.com/adaptive-cards)和其他跨产品功能，脚本工作簿通信的详细信息会变得更加复杂。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="9fd4f-146">如果需要帮助使脚本运行得更快，请通过 [堆栈溢出](https://stackoverflow.com/questions/tagged/office-scripts)。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="9fd4f-147">请务必使用 "office-scripts" 标记你的问题，以便专家可以找到它和帮助。</span><span class="sxs-lookup"><span data-stu-id="9fd4f-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="9fd4f-148">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9fd4f-148">See also</span></span>

- [<span data-ttu-id="9fd4f-149">Excel 网页版中 Office 脚本的脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="9fd4f-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="9fd4f-150">MDN web 文档：循环和迭代</span><span class="sxs-lookup"><span data-stu-id="9fd4f-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
