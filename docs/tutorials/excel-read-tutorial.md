---
title: 在 Excel 网页版中使用 Office 脚本读取工作簿数据
description: 有关从工作簿中读取数据并评估脚本中的数据的 Office 脚本教程。
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700103"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="dd57a-103">在 Excel 网页版中使用 Office 脚本读取工作簿数据</span><span class="sxs-lookup"><span data-stu-id="dd57a-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="dd57a-104">本教程将介绍如何在 Excel 网页版中使用 Office 脚本从工作簿中读取数据。</span><span class="sxs-lookup"><span data-stu-id="dd57a-104">This tutorial will teach you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="dd57a-105">然后，你将编辑所读取的数据，并将其放回工作簿中。</span><span class="sxs-lookup"><span data-stu-id="dd57a-105">You'll then edit the data you read and put it back in the workbook.</span></span>

> [!TIP]
> <span data-ttu-id="dd57a-106">如果你不熟悉 Office 脚本，建议先查看[在 Excel 网页版中录制、编辑和创建 Office 脚本](excel-tutorial.md)教程。</span><span class="sxs-lookup"><span data-stu-id="dd57a-106">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="dd57a-107">先决条件</span><span class="sxs-lookup"><span data-stu-id="dd57a-107">Prerequisites</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

<span data-ttu-id="dd57a-108">在开始本教程之前，你需要具有 Office 脚本的访问权限，要求如下：</span><span class="sxs-lookup"><span data-stu-id="dd57a-108">Before starting this tutorial, you'll need access to Office Scripts, which requires the following:</span></span>

- <span data-ttu-id="dd57a-109">[Excel 网页版](https://www.office.com/launch/excel)。</span><span class="sxs-lookup"><span data-stu-id="dd57a-109">[Excel on the web](https://www.office.com/launch/excel).</span></span>
- <span data-ttu-id="dd57a-110">要求管理员[为组织启用 Office 脚本](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)，这会将“自动”选项卡添加到功能区\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="dd57a-110">Ask your administrator to [enable Office Scripts for your organization](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), which adds the **Automate** tab to the ribbon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dd57a-111">本教程面向在 JavaScript 或 TypeScript 方面具备初级到中级知识的人员。</span><span class="sxs-lookup"><span data-stu-id="dd57a-111">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="dd57a-112">如果你不熟悉 JavaScript，建议查看 [Mozilla JavaScript 教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)。</span><span class="sxs-lookup"><span data-stu-id="dd57a-112">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="dd57a-113">请访问 [Excel 网页版中的 Office 脚本](../overview/excel.md)，以了解有关脚本环境的详细信息。</span><span class="sxs-lookup"><span data-stu-id="dd57a-113">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="read-a-cell"></a><span data-ttu-id="dd57a-114">读取单元格</span><span class="sxs-lookup"><span data-stu-id="dd57a-114">Read a cell</span></span>

<span data-ttu-id="dd57a-115">使用操作录制器创建的脚本只能将信息写入工作簿。</span><span class="sxs-lookup"><span data-stu-id="dd57a-115">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="dd57a-116">借助代码编辑器，可以编辑并创建也从工作簿中读取数据的脚本。</span><span class="sxs-lookup"><span data-stu-id="dd57a-116">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="dd57a-117">我们来创建一个读取数据并根据读取的数据执行操作的脚本。</span><span class="sxs-lookup"><span data-stu-id="dd57a-117">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="dd57a-118">我们将使用示例银行帐单。</span><span class="sxs-lookup"><span data-stu-id="dd57a-118">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="dd57a-119">此帐单是结合了支票和信贷的帐单。</span><span class="sxs-lookup"><span data-stu-id="dd57a-119">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="dd57a-120">遗憾的是，它们会以不同的方式报告余额变化。</span><span class="sxs-lookup"><span data-stu-id="dd57a-120">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="dd57a-121">支票帐单将收入作为正面信贷，将费用作为负面借记。</span><span class="sxs-lookup"><span data-stu-id="dd57a-121">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="dd57a-122">信贷帐单与之相反。</span><span class="sxs-lookup"><span data-stu-id="dd57a-122">The credit statement does the opposite.</span></span>

<span data-ttu-id="dd57a-123">在本教程的其余部分中，我们将使用脚本对此数据进行标准化。</span><span class="sxs-lookup"><span data-stu-id="dd57a-123">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="dd57a-124">首先，让我们来了解如何从工作簿中读取数据。</span><span class="sxs-lookup"><span data-stu-id="dd57a-124">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="dd57a-125">在用于教程其余部分的工作簿中创建新工作表。</span><span class="sxs-lookup"><span data-stu-id="dd57a-125">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="dd57a-126">复制以下数据，并将其粘贴到新工作表中，从单元格 **A1** 开始。</span><span class="sxs-lookup"><span data-stu-id="dd57a-126">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="dd57a-127">日期</span><span class="sxs-lookup"><span data-stu-id="dd57a-127">Date</span></span> |<span data-ttu-id="dd57a-128">帐户</span><span class="sxs-lookup"><span data-stu-id="dd57a-128">Account</span></span> |<span data-ttu-id="dd57a-129">说明</span><span class="sxs-lookup"><span data-stu-id="dd57a-129">Description</span></span> |<span data-ttu-id="dd57a-130">借记</span><span class="sxs-lookup"><span data-stu-id="dd57a-130">Debit</span></span> |<span data-ttu-id="dd57a-131">信贷</span><span class="sxs-lookup"><span data-stu-id="dd57a-131">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="dd57a-132">2019 年 10 月 10 日</span><span class="sxs-lookup"><span data-stu-id="dd57a-132">10/10/2019</span></span> |<span data-ttu-id="dd57a-133">支票</span><span class="sxs-lookup"><span data-stu-id="dd57a-133">Checking</span></span> |<span data-ttu-id="dd57a-134">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="dd57a-134">Coho Vineyard</span></span> |<span data-ttu-id="dd57a-135">-20.05</span><span class="sxs-lookup"><span data-stu-id="dd57a-135">-20.05</span></span> | |
    |<span data-ttu-id="dd57a-136">2019 年 10 月 11 日</span><span class="sxs-lookup"><span data-stu-id="dd57a-136">10/11/2019</span></span> |<span data-ttu-id="dd57a-137">信贷</span><span class="sxs-lookup"><span data-stu-id="dd57a-137">Credit</span></span> |<span data-ttu-id="dd57a-138">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="dd57a-138">The Phone Company</span></span> |<span data-ttu-id="dd57a-139">99.95</span><span class="sxs-lookup"><span data-stu-id="dd57a-139">99.95</span></span> | |
    |<span data-ttu-id="dd57a-140">2019 年 10 月 13 日</span><span class="sxs-lookup"><span data-stu-id="dd57a-140">10/13/2019</span></span> |<span data-ttu-id="dd57a-141">信贷</span><span class="sxs-lookup"><span data-stu-id="dd57a-141">Credit</span></span> |<span data-ttu-id="dd57a-142">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="dd57a-142">Coho Vineyard</span></span> |<span data-ttu-id="dd57a-143">154.43</span><span class="sxs-lookup"><span data-stu-id="dd57a-143">154.43</span></span> | |
    |<span data-ttu-id="dd57a-144">2019 年 10 月 15 日</span><span class="sxs-lookup"><span data-stu-id="dd57a-144">10/15/2019</span></span> |<span data-ttu-id="dd57a-145">支票</span><span class="sxs-lookup"><span data-stu-id="dd57a-145">Checking</span></span> |<span data-ttu-id="dd57a-146">外部存款</span><span class="sxs-lookup"><span data-stu-id="dd57a-146">External Deposit</span></span> | |<span data-ttu-id="dd57a-147">1000</span><span class="sxs-lookup"><span data-stu-id="dd57a-147">1000</span></span> |
    |<span data-ttu-id="dd57a-148">2019 年 10 月 20 日</span><span class="sxs-lookup"><span data-stu-id="dd57a-148">10/20/2019</span></span> |<span data-ttu-id="dd57a-149">信贷</span><span class="sxs-lookup"><span data-stu-id="dd57a-149">Credit</span></span> |<span data-ttu-id="dd57a-150">Coho Vineyard - 退款</span><span class="sxs-lookup"><span data-stu-id="dd57a-150">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="dd57a-151">-35.45</span><span class="sxs-lookup"><span data-stu-id="dd57a-151">-35.45</span></span> |
    |<span data-ttu-id="dd57a-152">2019 年 10 月 25 日</span><span class="sxs-lookup"><span data-stu-id="dd57a-152">10/25/2019</span></span> |<span data-ttu-id="dd57a-153">支票</span><span class="sxs-lookup"><span data-stu-id="dd57a-153">Checking</span></span> |<span data-ttu-id="dd57a-154">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="dd57a-154">Best For You Organics Company</span></span> | <span data-ttu-id="dd57a-155">-85.64</span><span class="sxs-lookup"><span data-stu-id="dd57a-155">-85.64</span></span> | |
    |<span data-ttu-id="dd57a-156">2019 年 11 月 1 日</span><span class="sxs-lookup"><span data-stu-id="dd57a-156">11/01/2019</span></span> |<span data-ttu-id="dd57a-157">支票</span><span class="sxs-lookup"><span data-stu-id="dd57a-157">Checking</span></span> |<span data-ttu-id="dd57a-158">外部存款</span><span class="sxs-lookup"><span data-stu-id="dd57a-158">External Deposit</span></span> | |<span data-ttu-id="dd57a-159">1000</span><span class="sxs-lookup"><span data-stu-id="dd57a-159">1000</span></span> |

3. <span data-ttu-id="dd57a-160">打开“代码编辑器”，然后选择“新建脚本”\*\*\*\*\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="dd57a-160">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="dd57a-161">让我们来清理格式。</span><span class="sxs-lookup"><span data-stu-id="dd57a-161">Let's clean up the formatting.</span></span> <span data-ttu-id="dd57a-162">这是一个财务文档，因此更改“借记”和“信贷”列中的数字格式以将值显示为美元金额\*\*\*\*\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="dd57a-162">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="dd57a-163">我们还调整列宽以适应数据。</span><span class="sxs-lookup"><span data-stu-id="dd57a-163">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="dd57a-164">将脚本内容替换为以下代码：</span><span class="sxs-lookup"><span data-stu-id="dd57a-164">Replace the script contents with the following code:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. <span data-ttu-id="dd57a-165">现在，让我们从数字列之一中读取一个值。</span><span class="sxs-lookup"><span data-stu-id="dd57a-165">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="dd57a-166">将以下代码添加到脚本末尾：</span><span class="sxs-lookup"><span data-stu-id="dd57a-166">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    <span data-ttu-id="dd57a-167">请注意对 `load` 和 `sync` 的调用。</span><span class="sxs-lookup"><span data-stu-id="dd57a-167">Note the calls to `load` and `sync`.</span></span> <span data-ttu-id="dd57a-168">你可以在 [Excel 网页版中的 Office 脚本的脚本基础知识](../develop/scripting-fundamentals.md#sync-and-load)中了解这些方法的详细信息。</span><span class="sxs-lookup"><span data-stu-id="dd57a-168">You can learn the details of those methods in [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md#sync-and-load).</span></span> <span data-ttu-id="dd57a-169">现在，我们知道你必须请求要读取的数据，然后将脚本与工作簿同步来读取该数据。</span><span class="sxs-lookup"><span data-stu-id="dd57a-169">For now, know that you must request data to be read and then sync your script with the workbook to read that data.</span></span>

6. <span data-ttu-id="dd57a-170">运行脚本。</span><span class="sxs-lookup"><span data-stu-id="dd57a-170">Run the script.</span></span>
7. <span data-ttu-id="dd57a-171">打开控制台。</span><span class="sxs-lookup"><span data-stu-id="dd57a-171">Open the console.</span></span> <span data-ttu-id="dd57a-172">转到“省略号”菜单，然后按“日志...”\*\*\*\*\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="dd57a-172">Go to the **Ellipses** menu and press **Logs...**.</span></span>
8. <span data-ttu-id="dd57a-173">应在控制台中看到 `[Array[1]]`。</span><span class="sxs-lookup"><span data-stu-id="dd57a-173">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="dd57a-174">这不是数字，因为区域是数据的二维数组。</span><span class="sxs-lookup"><span data-stu-id="dd57a-174">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="dd57a-175">该二维区域直接记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="dd57a-175">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="dd57a-176">幸运的是，代码编辑器可以让你看到数组的内容。</span><span class="sxs-lookup"><span data-stu-id="dd57a-176">Luckily, the Code Editor does let you see the contents of the array.</span></span>
9. <span data-ttu-id="dd57a-177">将二维数组记录到控制台时，它会对每行下面的列值进行分组。</span><span class="sxs-lookup"><span data-stu-id="dd57a-177">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="dd57a-178">按蓝色三角形展开数组日志。</span><span class="sxs-lookup"><span data-stu-id="dd57a-178">Expand the array log by pressing the blue triangle.</span></span>
10. <span data-ttu-id="dd57a-179">按新出现的蓝色三角形展开数组的第二级别。</span><span class="sxs-lookup"><span data-stu-id="dd57a-179">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="dd57a-180">现在，你应该会看到：</span><span class="sxs-lookup"><span data-stu-id="dd57a-180">You should now see this:</span></span>

    ![控制台日志显示嵌套在两个数组下的输出“-20.05”。](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="dd57a-182">修改单元格的值</span><span class="sxs-lookup"><span data-stu-id="dd57a-182">Modify the value of a cell</span></span>

<span data-ttu-id="dd57a-183">现在，我们可以读取数据，让我们使用该数据来修改工作簿。</span><span class="sxs-lookup"><span data-stu-id="dd57a-183">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="dd57a-184">使单元格 **D2** 的值与 `Math.abs` 函数呈正相关。</span><span class="sxs-lookup"><span data-stu-id="dd57a-184">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="dd57a-185">[Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) 对象包含许多脚本具有访问权限的函数。</span><span class="sxs-lookup"><span data-stu-id="dd57a-185">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="dd57a-186">可在[使用 Office 脚本中的内置 JavaScript 对象](../develop/javascript-objects.md)中找到有关 `Math` 和其他内置对象的详细信息。</span><span class="sxs-lookup"><span data-stu-id="dd57a-186">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="dd57a-187">将以下代码添加到脚本末尾：</span><span class="sxs-lookup"><span data-stu-id="dd57a-187">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. <span data-ttu-id="dd57a-188">单元格 **D2** 的值现在应为正值。</span><span class="sxs-lookup"><span data-stu-id="dd57a-188">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="dd57a-189">修改列的值</span><span class="sxs-lookup"><span data-stu-id="dd57a-189">Modify the values of a column</span></span>

<span data-ttu-id="dd57a-190">现在，我们知道如何读取和写入单个单元格，让我们对脚本进行一般化，使其适用于整个“借记”和“信贷”列\*\*\*\*\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="dd57a-190">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="dd57a-191">删除仅影响单个单元格的代码（先前的绝对值代码），以便你的脚本现在如下所示：</span><span class="sxs-lookup"><span data-stu-id="dd57a-191">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. <span data-ttu-id="dd57a-192">添加循环访问最后两列中的行的循环。</span><span class="sxs-lookup"><span data-stu-id="dd57a-192">Add a loop that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="dd57a-193">对于每个单元格，脚本将值设置为当前值的绝对值。</span><span class="sxs-lookup"><span data-stu-id="dd57a-193">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="dd57a-194">请注意，定义单元格位置的数组是从零开始的。</span><span class="sxs-lookup"><span data-stu-id="dd57a-194">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="dd57a-195">这意味着单元格 **A1** 为 `range[0][0]`。</span><span class="sxs-lookup"><span data-stu-id="dd57a-195">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    <span data-ttu-id="dd57a-196">此部分的脚本执行几项重要任务。</span><span class="sxs-lookup"><span data-stu-id="dd57a-196">This portion of the script does several important tasks.</span></span> <span data-ttu-id="dd57a-197">首先，加载已用区域的值和行计数。</span><span class="sxs-lookup"><span data-stu-id="dd57a-197">First, it loads the values and row count of the used range.</span></span> <span data-ttu-id="dd57a-198">这样，我们就可以查看值并知道何时停止。</span><span class="sxs-lookup"><span data-stu-id="dd57a-198">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="dd57a-199">其次，循环访问已用区域，检查“借记”或“信贷”列中的每个单元格\*\*\*\*\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="dd57a-199">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="dd57a-200">最后，如果单元格中的值不为 0，则该值将替换为其绝对值。</span><span class="sxs-lookup"><span data-stu-id="dd57a-200">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="dd57a-201">我们正在避免使用零，因此可以将空白单元格保留原样。</span><span class="sxs-lookup"><span data-stu-id="dd57a-201">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="dd57a-202">运行脚本。</span><span class="sxs-lookup"><span data-stu-id="dd57a-202">Run the script.</span></span>

    <span data-ttu-id="dd57a-203">现在，你的银行帐单如下所示：</span><span class="sxs-lookup"><span data-stu-id="dd57a-203">Your banking statement should now look like this:</span></span>

    ![银行帐单作为仅包含正值的格式表。](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="dd57a-205">后续步骤</span><span class="sxs-lookup"><span data-stu-id="dd57a-205">Next steps</span></span>

<span data-ttu-id="dd57a-206">打开“代码编辑器”，然后尝试使用一些 [Excel 网页版中的 Office 脚本的示例脚本](../resources/excel-samples.md)。</span><span class="sxs-lookup"><span data-stu-id="dd57a-206">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="dd57a-207">还可以访问 [Excel 网页版中的 Office 脚本的脚本基础知识](../develop/scripting-fundamentals.md)，了解有关创建 Office 脚本的详细信息。</span><span class="sxs-lookup"><span data-stu-id="dd57a-207">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>
