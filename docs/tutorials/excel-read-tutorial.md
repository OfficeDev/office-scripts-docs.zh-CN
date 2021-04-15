---
title: 在 Excel 网页版中使用 Office 脚本读取工作簿数据
description: 有关从工作簿中读取数据并评估脚本中的数据的 Office 脚本教程。
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: d6321cb91a425da3fd45329d5171f1d5694b2b99
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754852"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="58581-103">在 Excel 网页版中使用 Office 脚本读取工作簿数据</span><span class="sxs-lookup"><span data-stu-id="58581-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="58581-104">本教程将介绍如何在 Excel 网页版中使用 Office 脚本从工作簿中读取数据。</span><span class="sxs-lookup"><span data-stu-id="58581-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="58581-105">你将编写一个新脚本，该脚本可设置银行对帐单的格式并规范化该对帐单中的数据。</span><span class="sxs-lookup"><span data-stu-id="58581-105">You'll be writing a new script that formats a bank statement and normalizes the data in that statement.</span></span> <span data-ttu-id="58581-106">在此数据清理过程中，你的脚本将从事务单元格中读取值，将一个简单的公式应用到每个值，并将生成的答案写入工作簿。</span><span class="sxs-lookup"><span data-stu-id="58581-106">As part of that data clean-up, your script will read values from the transaction cells, apply a simple formula to each value, and write the resulting answer to the workbook.</span></span> <span data-ttu-id="58581-107">通过从工作簿中读取数据，可在脚本中自动执行某些决策过程。</span><span class="sxs-lookup"><span data-stu-id="58581-107">Reading data from the workbook lets you automate some of your decision making processes in the script.</span></span>

> [!TIP]
> <span data-ttu-id="58581-108">如果你不熟悉 Office 脚本，建议先查看[在 Excel 网页版中录制、编辑和创建 Office 脚本](excel-tutorial.md)教程。</span><span class="sxs-lookup"><span data-stu-id="58581-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="58581-109">[Office 脚本使用 TypeScript](../overview/code-editor-environment.md)，本教程面向在 JavaScript 或 TypeScript 方面具备初级到中级知识的人员。</span><span class="sxs-lookup"><span data-stu-id="58581-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="58581-110">如果你不熟悉 JavaScript，建议从 [Mozilla JavaScript 教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)入手。</span><span class="sxs-lookup"><span data-stu-id="58581-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="58581-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="58581-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## <a name="read-a-cell"></a><span data-ttu-id="58581-112">读取单元格</span><span class="sxs-lookup"><span data-stu-id="58581-112">Read a cell</span></span>

<span data-ttu-id="58581-113">使用操作录制器创建的脚本只能将信息写入工作簿。</span><span class="sxs-lookup"><span data-stu-id="58581-113">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="58581-114">借助代码编辑器，可以编辑并创建也从工作簿中读取数据的脚本。</span><span class="sxs-lookup"><span data-stu-id="58581-114">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="58581-115">我们来创建一个读取数据并根据读取的数据执行操作的脚本。</span><span class="sxs-lookup"><span data-stu-id="58581-115">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="58581-116">我们将使用示例银行帐单。</span><span class="sxs-lookup"><span data-stu-id="58581-116">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="58581-117">此帐单是结合了支票和信贷的帐单。</span><span class="sxs-lookup"><span data-stu-id="58581-117">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="58581-118">遗憾的是，它们会以不同的方式报告余额变化。</span><span class="sxs-lookup"><span data-stu-id="58581-118">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="58581-119">支票帐单将收入作为正面信贷，将费用作为负面借记。</span><span class="sxs-lookup"><span data-stu-id="58581-119">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="58581-120">信贷帐单与之相反。</span><span class="sxs-lookup"><span data-stu-id="58581-120">The credit statement does the opposite.</span></span>

<span data-ttu-id="58581-121">在本教程的其余部分中，我们将使用脚本对此数据进行标准化。</span><span class="sxs-lookup"><span data-stu-id="58581-121">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="58581-122">首先，让我们来了解如何从工作簿中读取数据。</span><span class="sxs-lookup"><span data-stu-id="58581-122">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="58581-123">在用于教程其余部分的工作簿中创建新工作表。</span><span class="sxs-lookup"><span data-stu-id="58581-123">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="58581-124">复制以下数据，并将其粘贴到新工作表中，从单元格 **A1** 开始。</span><span class="sxs-lookup"><span data-stu-id="58581-124">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="58581-125">日期</span><span class="sxs-lookup"><span data-stu-id="58581-125">Date</span></span> |<span data-ttu-id="58581-126">帐户</span><span class="sxs-lookup"><span data-stu-id="58581-126">Account</span></span> |<span data-ttu-id="58581-127">说明</span><span class="sxs-lookup"><span data-stu-id="58581-127">Description</span></span> |<span data-ttu-id="58581-128">借记</span><span class="sxs-lookup"><span data-stu-id="58581-128">Debit</span></span> |<span data-ttu-id="58581-129">信贷</span><span class="sxs-lookup"><span data-stu-id="58581-129">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="58581-130">2019 年 10 月 10 日</span><span class="sxs-lookup"><span data-stu-id="58581-130">10/10/2019</span></span> |<span data-ttu-id="58581-131">支票</span><span class="sxs-lookup"><span data-stu-id="58581-131">Checking</span></span> |<span data-ttu-id="58581-132">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="58581-132">Coho Vineyard</span></span> |<span data-ttu-id="58581-133">-20.05</span><span class="sxs-lookup"><span data-stu-id="58581-133">-20.05</span></span> | |
    |<span data-ttu-id="58581-134">2019 年 10 月 11 日</span><span class="sxs-lookup"><span data-stu-id="58581-134">10/11/2019</span></span> |<span data-ttu-id="58581-135">信贷</span><span class="sxs-lookup"><span data-stu-id="58581-135">Credit</span></span> |<span data-ttu-id="58581-136">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="58581-136">The Phone Company</span></span> |<span data-ttu-id="58581-137">99.95</span><span class="sxs-lookup"><span data-stu-id="58581-137">99.95</span></span> | |
    |<span data-ttu-id="58581-138">2019 年 10 月 13 日</span><span class="sxs-lookup"><span data-stu-id="58581-138">10/13/2019</span></span> |<span data-ttu-id="58581-139">信贷</span><span class="sxs-lookup"><span data-stu-id="58581-139">Credit</span></span> |<span data-ttu-id="58581-140">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="58581-140">Coho Vineyard</span></span> |<span data-ttu-id="58581-141">154.43</span><span class="sxs-lookup"><span data-stu-id="58581-141">154.43</span></span> | |
    |<span data-ttu-id="58581-142">2019 年 10 月 15 日</span><span class="sxs-lookup"><span data-stu-id="58581-142">10/15/2019</span></span> |<span data-ttu-id="58581-143">支票</span><span class="sxs-lookup"><span data-stu-id="58581-143">Checking</span></span> |<span data-ttu-id="58581-144">外部存款</span><span class="sxs-lookup"><span data-stu-id="58581-144">External Deposit</span></span> | |<span data-ttu-id="58581-145">1000</span><span class="sxs-lookup"><span data-stu-id="58581-145">1000</span></span> |
    |<span data-ttu-id="58581-146">2019 年 10 月 20 日</span><span class="sxs-lookup"><span data-stu-id="58581-146">10/20/2019</span></span> |<span data-ttu-id="58581-147">信贷</span><span class="sxs-lookup"><span data-stu-id="58581-147">Credit</span></span> |<span data-ttu-id="58581-148">Coho Vineyard - 退款</span><span class="sxs-lookup"><span data-stu-id="58581-148">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="58581-149">-35.45</span><span class="sxs-lookup"><span data-stu-id="58581-149">-35.45</span></span> |
    |<span data-ttu-id="58581-150">2019 年 10 月 25 日</span><span class="sxs-lookup"><span data-stu-id="58581-150">10/25/2019</span></span> |<span data-ttu-id="58581-151">支票</span><span class="sxs-lookup"><span data-stu-id="58581-151">Checking</span></span> |<span data-ttu-id="58581-152">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="58581-152">Best For You Organics Company</span></span> | <span data-ttu-id="58581-153">-85.64</span><span class="sxs-lookup"><span data-stu-id="58581-153">-85.64</span></span> | |
    |<span data-ttu-id="58581-154">2019 年 11 月 1 日</span><span class="sxs-lookup"><span data-stu-id="58581-154">11/01/2019</span></span> |<span data-ttu-id="58581-155">支票</span><span class="sxs-lookup"><span data-stu-id="58581-155">Checking</span></span> |<span data-ttu-id="58581-156">外部存款</span><span class="sxs-lookup"><span data-stu-id="58581-156">External Deposit</span></span> | |<span data-ttu-id="58581-157">1000</span><span class="sxs-lookup"><span data-stu-id="58581-157">1000</span></span> |

3. <span data-ttu-id="58581-158">打开“**所有脚本**”，然后选择“**新脚本**”。</span><span class="sxs-lookup"><span data-stu-id="58581-158">Open **All Scripts** and select **New Script**.</span></span>
4. <span data-ttu-id="58581-159">让我们来清理格式。</span><span class="sxs-lookup"><span data-stu-id="58581-159">Let's clean up the formatting.</span></span> <span data-ttu-id="58581-160">这是一个财务文档，因此更改“借记”和“信贷”列中的数字格式以将值显示为美元金额。</span><span class="sxs-lookup"><span data-stu-id="58581-160">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="58581-161">我们还调整列宽以适应数据。</span><span class="sxs-lookup"><span data-stu-id="58581-161">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="58581-162">将脚本内容替换为以下代码：</span><span class="sxs-lookup"><span data-stu-id="58581-162">Replace the script contents with the following code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. <span data-ttu-id="58581-163">现在，让我们从数字列之一中读取一个值。</span><span class="sxs-lookup"><span data-stu-id="58581-163">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="58581-164">将以下代码添加到脚本的末尾（在结束 `}` 之前）：</span><span class="sxs-lookup"><span data-stu-id="58581-164">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="58581-165">运行脚本。</span><span class="sxs-lookup"><span data-stu-id="58581-165">Run the script.</span></span>
7. <span data-ttu-id="58581-166">应在控制台中看到 `[Array[1]]`。</span><span class="sxs-lookup"><span data-stu-id="58581-166">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="58581-167">这不是数字，因为区域是数据的二维数组。</span><span class="sxs-lookup"><span data-stu-id="58581-167">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="58581-168">该二维区域直接记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="58581-168">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="58581-169">幸运的是，代码编辑器让你能够看到数组的内容。</span><span class="sxs-lookup"><span data-stu-id="58581-169">Luckily, the Code Editor lets you see the contents of the array.</span></span>
8. <span data-ttu-id="58581-170">将二维数组记录到控制台时，它会对每行下面的列值进行分组。</span><span class="sxs-lookup"><span data-stu-id="58581-170">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="58581-171">按蓝色三角形展开数组日志。</span><span class="sxs-lookup"><span data-stu-id="58581-171">Expand the array log by pressing the blue triangle.</span></span>
9. <span data-ttu-id="58581-172">按新出现的蓝色三角形展开数组的第二级别。</span><span class="sxs-lookup"><span data-stu-id="58581-172">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="58581-173">现在，你应该会看到：</span><span class="sxs-lookup"><span data-stu-id="58581-173">You should now see this:</span></span>

    :::image type="content" source="../images/tutorial-4.png" alt-text="控制台日志显示输出&quot;-20.05&quot;，嵌套在两数组":::

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="58581-175">修改单元格的值</span><span class="sxs-lookup"><span data-stu-id="58581-175">Modify the value of a cell</span></span>

<span data-ttu-id="58581-176">现在，我们可以读取数据，让我们使用该数据来修改工作簿。</span><span class="sxs-lookup"><span data-stu-id="58581-176">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="58581-177">使单元格 **D2** 的值与 `Math.abs` 函数呈正相关。</span><span class="sxs-lookup"><span data-stu-id="58581-177">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="58581-178">[Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) 对象包含许多脚本具有访问权限的函数。</span><span class="sxs-lookup"><span data-stu-id="58581-178">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="58581-179">可在[使用 Office 脚本中的内置 JavaScript 对象](../develop/javascript-objects.md)中找到有关 `Math` 和其他内置对象的详细信息。</span><span class="sxs-lookup"><span data-stu-id="58581-179">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="58581-180">我们将使用 `getValue` 和 `setValue` 方法更改单元格的值。</span><span class="sxs-lookup"><span data-stu-id="58581-180">We'll use `getValue` and `setValue` methods to change the value of the cell.</span></span> <span data-ttu-id="58581-181">这些方法适用于单个单元格。</span><span class="sxs-lookup"><span data-stu-id="58581-181">These methods work on a single cell.</span></span> <span data-ttu-id="58581-182">处理多单元格区域时，需使用 `getValues` 和 `setValues`。</span><span class="sxs-lookup"><span data-stu-id="58581-182">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span> <span data-ttu-id="58581-183">将以下代码添加到脚本末尾：</span><span class="sxs-lookup"><span data-stu-id="58581-183">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue() as number);
    range.setValue(positiveValue);
    ```

    > [!NOTE]
    > <span data-ttu-id="58581-184">我们正使用 `as` 关键字将 `range.getValue()` 的返回值 [转换](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) 为 `number`。 </span><span class="sxs-lookup"><span data-stu-id="58581-184">We are [casting](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) the returned value of `range.getValue()` to a `number` by using the `as` keyword.</span></span> <span data-ttu-id="58581-185">这样做很有必要，因为区域可能是字符串、数字或布尔值。</span><span class="sxs-lookup"><span data-stu-id="58581-185">This is necessary because a range could be strings, numbers, or booleans.</span></span> <span data-ttu-id="58581-186">在本实例中，我们明确需要数字。</span><span class="sxs-lookup"><span data-stu-id="58581-186">In this instance, we explicitly need a number.</span></span>

2. <span data-ttu-id="58581-187">单元格 **D2** 的值现在应为正值。</span><span class="sxs-lookup"><span data-stu-id="58581-187">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="58581-188">修改列的值</span><span class="sxs-lookup"><span data-stu-id="58581-188">Modify the values of a column</span></span>

<span data-ttu-id="58581-189">现在，我们知道如何读取和写入单个单元格，让我们对脚本进行一般化，使其适用于整个“借记”和“信贷”列。</span><span class="sxs-lookup"><span data-stu-id="58581-189">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="58581-190">删除仅影响单个单元格的代码（先前的绝对值代码），以便你的脚本现在如下所示：</span><span class="sxs-lookup"><span data-stu-id="58581-190">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. <span data-ttu-id="58581-191">在脚本末尾添加循环访问最后两列中的行的循环。</span><span class="sxs-lookup"><span data-stu-id="58581-191">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="58581-192">对于每个单元格，脚本将值设置为当前值的绝对值。</span><span class="sxs-lookup"><span data-stu-id="58581-192">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="58581-193">请注意，定义单元格位置的数组是从零开始的。</span><span class="sxs-lookup"><span data-stu-id="58581-193">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="58581-194">这意味着单元格 **A1** 为 `range[0][0]`。</span><span class="sxs-lookup"><span data-stu-id="58581-194">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3] as number);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4] as number);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    <span data-ttu-id="58581-195">此部分的脚本执行几项重要任务。</span><span class="sxs-lookup"><span data-stu-id="58581-195">This portion of the script does several important tasks.</span></span> <span data-ttu-id="58581-196">首先，获取已用区域的值和行计数。</span><span class="sxs-lookup"><span data-stu-id="58581-196">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="58581-197">这样，我们就可以查看值并知道何时停止。</span><span class="sxs-lookup"><span data-stu-id="58581-197">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="58581-198">其次，循环访问已用区域，检查“借记”或“信贷”列中的每个单元格。</span><span class="sxs-lookup"><span data-stu-id="58581-198">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="58581-199">最后，如果单元格中的值不为 0，则该值将替换为其绝对值。</span><span class="sxs-lookup"><span data-stu-id="58581-199">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="58581-200">我们正在避免使用零，因此可以将空白单元格保留原样。</span><span class="sxs-lookup"><span data-stu-id="58581-200">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="58581-201">运行脚本。</span><span class="sxs-lookup"><span data-stu-id="58581-201">Run the script.</span></span>

    <span data-ttu-id="58581-202">现在，你的银行帐单如下所示：</span><span class="sxs-lookup"><span data-stu-id="58581-202">Your banking statement should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-5.png" alt-text="一个工作表，显示银行对账单为仅具有正值的带格式的表格。":::

## <a name="next-steps"></a><span data-ttu-id="58581-204">后续步骤</span><span class="sxs-lookup"><span data-stu-id="58581-204">Next steps</span></span>

<span data-ttu-id="58581-205">打开“代码编辑器”，然后尝试使用一些 [Excel 网页版中的 Office 脚本的示例脚本](../resources/excel-samples.md)。</span><span class="sxs-lookup"><span data-stu-id="58581-205">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="58581-206">还可以访问 [Excel 网页版中的 Office 脚本的脚本基础知识](../develop/scripting-fundamentals.md)，了解有关创建 Office 脚本的详细信息。</span><span class="sxs-lookup"><span data-stu-id="58581-206">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>

<span data-ttu-id="58581-207">下一系列的 Office 脚本教程重点介绍如何将 Office 脚本与 Power Automate 一起使用。</span><span class="sxs-lookup"><span data-stu-id="58581-207">The next series of Office Scripts tutorials focus on using Office Scripts with Power Automate.</span></span> <span data-ttu-id="58581-208">在[使用 Power Automate 运行 Office 脚本](../develop/power-automate-integration.md)中了解有关结合两个平台的优势的更多信息，或尝试[通过 Power Automate 手动流调用脚本](excel-power-automate-manual.md)教程来创建使用 Office 脚本的 Power Automate 流。</span><span class="sxs-lookup"><span data-stu-id="58581-208">Learn more about the advantages combining the two platforms in [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) or try the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial to create a Power Automate flow that uses an Office Script.</span></span>
