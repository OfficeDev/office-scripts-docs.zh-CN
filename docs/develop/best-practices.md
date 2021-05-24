---
title: Office 脚本中的最佳实践
description: 如何防止常见问题，并编写可Office输入或数据的稳固脚本。
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546020"
---
# <a name="best-practices-in-office-scripts"></a><span data-ttu-id="98050-103">Office 脚本中的最佳实践</span><span class="sxs-lookup"><span data-stu-id="98050-103">Best practices in Office Scripts</span></span>

<span data-ttu-id="98050-104">这些模式和做法旨在帮助脚本每次成功运行。</span><span class="sxs-lookup"><span data-stu-id="98050-104">These patterns and practices are designed to help your scripts run successfully every time.</span></span> <span data-ttu-id="98050-105">使用它们可以避免开始自动执行工作流时出现Excel错误。</span><span class="sxs-lookup"><span data-stu-id="98050-105">Use them to avoid common pitfalls as you start automating your Excel workflow.</span></span>

## <a name="verify-an-object-is-present"></a><span data-ttu-id="98050-106">验证对象是否存在</span><span class="sxs-lookup"><span data-stu-id="98050-106">Verify an object is present</span></span>

<span data-ttu-id="98050-107">脚本通常依赖于工作簿中呈现的某个工作表或表。</span><span class="sxs-lookup"><span data-stu-id="98050-107">Scripts often rely on a certain worksheet or table being present in the workbook.</span></span> <span data-ttu-id="98050-108">但是，在脚本运行之间，它们可能会重命名或删除。</span><span class="sxs-lookup"><span data-stu-id="98050-108">However, they might get renamed or removed between script runs.</span></span> <span data-ttu-id="98050-109">通过先检查这些表或工作表是否存在，然后再对它们调用方法，您可以确保脚本不会突然结束。</span><span class="sxs-lookup"><span data-stu-id="98050-109">By checking if those tables or worksheets exist before calling methods on them, you can make sure the script doesn't end abruptly.</span></span>

<span data-ttu-id="98050-110">以下示例代码检查工作簿中是否包含"索引"工作表。</span><span class="sxs-lookup"><span data-stu-id="98050-110">The following sample code checks if the "Index" worksheet is present in the workbook.</span></span> <span data-ttu-id="98050-111">如果工作表存在，脚本将获取一个范围并继续。</span><span class="sxs-lookup"><span data-stu-id="98050-111">If the worksheet is present, the script gets a range and proceeds.</span></span> <span data-ttu-id="98050-112">如果不存在，脚本将记录自定义错误消息。</span><span class="sxs-lookup"><span data-stu-id="98050-112">If it isn't present, the script logs a custom error message.</span></span>

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

<span data-ttu-id="98050-113">TypeScript `?` 运算符在调用方法之前检查对象是否存在。</span><span class="sxs-lookup"><span data-stu-id="98050-113">The TypeScript `?` operator checks if the object exists before calling a method.</span></span> <span data-ttu-id="98050-114">如果不需要在对象不存在时执行任何特殊操作，这可以使代码更加简化。</span><span class="sxs-lookup"><span data-stu-id="98050-114">This can make your code more streamlined if you don't need to do anything special when the object doesn't exist.</span></span>

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a><span data-ttu-id="98050-115">首先验证数据和工作簿状态</span><span class="sxs-lookup"><span data-stu-id="98050-115">Validate data and workbook state first</span></span>

<span data-ttu-id="98050-116">处理数据之前，请确保存在所有工作表、表、形状和其他对象。</span><span class="sxs-lookup"><span data-stu-id="98050-116">Make sure all your worksheets, tables, shapes, and other objects are present before working on the data.</span></span> <span data-ttu-id="98050-117">使用以前的模式，检查所有内容是否都位于工作簿中并符合您的预期。</span><span class="sxs-lookup"><span data-stu-id="98050-117">Using the previous pattern, check to see if everything is in the workbook and matches your expectations.</span></span> <span data-ttu-id="98050-118">在写入任何数据之前执行此操作可确保脚本不会使工作簿保持部分状态。</span><span class="sxs-lookup"><span data-stu-id="98050-118">Doing this before any data is written ensures your script doesn't leave the workbook in a partial state.</span></span>

<span data-ttu-id="98050-119">以下脚本要求存在两个名为"Table1"和"Table2"的表。</span><span class="sxs-lookup"><span data-stu-id="98050-119">The following script requires two tables named "Table1" and "Table2" to be present.</span></span> <span data-ttu-id="98050-120">该脚本首先检查表是否存在，然后以 语句和相应的消息结尾（如果没有 `return` ）。</span><span class="sxs-lookup"><span data-stu-id="98050-120">The script first checks if the tables are present and then ends with the `return` statement and an appropriate message if they're not.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="98050-121">如果验证发生在单独的函数中，您仍必须通过从 函数发出 语句 `return` 来结束 `main` 脚本。</span><span class="sxs-lookup"><span data-stu-id="98050-121">If the verification is happening in a separate function, you still must end the script by issuing the `return` statement from the `main` function.</span></span> <span data-ttu-id="98050-122">从子函数返回不会结束脚本。</span><span class="sxs-lookup"><span data-stu-id="98050-122">Returning from the subfunction doesn't end the script.</span></span>

<span data-ttu-id="98050-123">以下脚本与上一脚本具有相同的行为。</span><span class="sxs-lookup"><span data-stu-id="98050-123">The following script has the same behavior as the previous one.</span></span> <span data-ttu-id="98050-124">区别在于函数 `main` 调用 函数 `inputPresent` 来验证所有内容。</span><span class="sxs-lookup"><span data-stu-id="98050-124">The difference is that the `main` function calls the `inputPresent` function to verify everything.</span></span> <span data-ttu-id="98050-125">`inputPresent` 返回一个 boolean (`true` 或 `false`) 以指示是否存在所有必需的输入。</span><span class="sxs-lookup"><span data-stu-id="98050-125">`inputPresent` returns a boolean (`true` or `false`) to indicate whether all required inputs are present.</span></span> <span data-ttu-id="98050-126">函数 `main` 使用该布尔值决定继续或结束脚本。</span><span class="sxs-lookup"><span data-stu-id="98050-126">The `main` function uses that boolean to decide on continuing or ending the script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a><span data-ttu-id="98050-127">何时使用 `throw` 语句</span><span class="sxs-lookup"><span data-stu-id="98050-127">When to use a `throw` statement</span></span>

<span data-ttu-id="98050-128">语句 [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) 指示发生了意外错误。</span><span class="sxs-lookup"><span data-stu-id="98050-128">A [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) statement indicates an unexpected error has occurred.</span></span> <span data-ttu-id="98050-129">它立即结束代码。</span><span class="sxs-lookup"><span data-stu-id="98050-129">It ends the code immediately.</span></span> <span data-ttu-id="98050-130">大多数情况下，不需要从 `throw` 脚本执行。</span><span class="sxs-lookup"><span data-stu-id="98050-130">For the most part, you don't need to `throw` from your script.</span></span> <span data-ttu-id="98050-131">通常，脚本会自动通知用户脚本由于问题无法运行。</span><span class="sxs-lookup"><span data-stu-id="98050-131">Usually, the script automatically informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="98050-132">在大多数情况下，用一条错误消息和函数中的语句结束 `return` 脚本 `main` 就足够了。</span><span class="sxs-lookup"><span data-stu-id="98050-132">In most cases, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="98050-133">但是，如果您的脚本作为流Power Automate运行，您可能需要阻止该流继续运行。</span><span class="sxs-lookup"><span data-stu-id="98050-133">However, if your script is running as part of a Power Automate flow, you may want to stop the flow from continuing.</span></span> <span data-ttu-id="98050-134">`throw`语句将停止脚本，并指示流也停止。</span><span class="sxs-lookup"><span data-stu-id="98050-134">A `throw` statement stops the script and tells the flow to stop as well.</span></span>

<span data-ttu-id="98050-135">以下脚本显示如何使用表 `throw` 检查示例中的 语句。</span><span class="sxs-lookup"><span data-stu-id="98050-135">The following script shows how to use the `throw` statement in our table checking example.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a><span data-ttu-id="98050-136">何时使用 `try...catch` 语句</span><span class="sxs-lookup"><span data-stu-id="98050-136">When to use a `try...catch` statement</span></span>

<span data-ttu-id="98050-137">语句 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) 是检测 API 调用是否失败并继续运行脚本的方法。</span><span class="sxs-lookup"><span data-stu-id="98050-137">The [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statement is a way to detect if an API call fails and continue running the script.</span></span>

<span data-ttu-id="98050-138">请考虑以下对区域执行大型数据更新的代码段。</span><span class="sxs-lookup"><span data-stu-id="98050-138">Consider the following snippet that performs a large data update on a range.</span></span>

```TypeScript
range.setValues(someLargeValues);
```

<span data-ttu-id="98050-139">如果 `someLargeValues` 大于 web Excel处理， `setValues()` 调用将失败。</span><span class="sxs-lookup"><span data-stu-id="98050-139">If `someLargeValues` is larger than Excel for the web can handle, the `setValues()` call fails.</span></span> <span data-ttu-id="98050-140">脚本随后也会失败，出现 [运行时错误](../testing/troubleshooting.md#runtime-errors)。</span><span class="sxs-lookup"><span data-stu-id="98050-140">The script then also fails with a [runtime error](../testing/troubleshooting.md#runtime-errors).</span></span> <span data-ttu-id="98050-141">语句 `try...catch` 使脚本能够识别此情况，而不会立即结束脚本并显示默认错误。</span><span class="sxs-lookup"><span data-stu-id="98050-141">The `try...catch` statement lets your script recognize this condition, without immediately ending the script and showing the default error.</span></span>

<span data-ttu-id="98050-142">为脚本用户提供更好的体验的一个方法是向用户显示自定义错误消息。</span><span class="sxs-lookup"><span data-stu-id="98050-142">One approach for giving the script user a better experience is to present them a custom error message.</span></span> <span data-ttu-id="98050-143">以下代码段显示了 `try...catch` 一个语句，它记录更多的错误信息，以更好地帮助读者。</span><span class="sxs-lookup"><span data-stu-id="98050-143">The following snippet shows a `try...catch` statement logging more error information to better help the reader.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

<span data-ttu-id="98050-144">处理错误的另一个方法是具有处理错误案例的回退行为。</span><span class="sxs-lookup"><span data-stu-id="98050-144">Another approach to dealing with errors is to have fallback behavior that handles the error case.</span></span> <span data-ttu-id="98050-145">以下代码段使用 `catch` 块尝试备用方法将更新分解为较小的部分，并避免错误。</span><span class="sxs-lookup"><span data-stu-id="98050-145">The following snippet uses the `catch` block to try an alternate method break up the update into smaller pieces and avoid the error.</span></span>

> [!TIP]
> <span data-ttu-id="98050-146">有关如何更新较大区域的完整示例，请参阅编写 [大型数据集](../resources/samples/write-large-dataset.md)。</span><span class="sxs-lookup"><span data-stu-id="98050-146">For a full example on how to update a large range, see [Write a large dataset](../resources/samples/write-large-dataset.md).</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> <span data-ttu-id="98050-147">在 `try...catch` 循环内部或周围使用会降低脚本的速度。</span><span class="sxs-lookup"><span data-stu-id="98050-147">Using `try...catch` inside or around a loop slows down your script.</span></span> <span data-ttu-id="98050-148">有关更多性能信息，请参阅 [避免使用 `try...catch` 块](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)。</span><span class="sxs-lookup"><span data-stu-id="98050-148">For more performance information, see [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span></span>

## <a name="see-also"></a><span data-ttu-id="98050-149">另请参阅</span><span class="sxs-lookup"><span data-stu-id="98050-149">See also</span></span>

- [<span data-ttu-id="98050-150">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="98050-150">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="98050-151">有关使用脚本Power Automate疑Office信息</span><span class="sxs-lookup"><span data-stu-id="98050-151">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="98050-152">Office 脚本的平台限制</span><span class="sxs-lookup"><span data-stu-id="98050-152">Platform limits with Office Scripts</span></span>](../testing/platform-limits.md)
- [<span data-ttu-id="98050-153">提高脚本Office性能</span><span class="sxs-lookup"><span data-stu-id="98050-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
