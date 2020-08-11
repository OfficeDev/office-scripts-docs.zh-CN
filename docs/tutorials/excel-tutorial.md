---
title: 在 Excel 网页版中录制、编辑和创建 Office 脚本
description: 有关 Office 脚本基础知识的教程，包括使用操作录制器录制脚本以及将数据写入工作簿。
ms.date: 07/21/2020
localization_priority: Priority
ms.openlocfilehash: 96bdc286883d87249de260666c7c8ffe2c94cc0f
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616771"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="c7014-103">在 Excel 网页版中录制、编辑和创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="c7014-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="c7014-104">本教程将提供有关为 Excel 网页版录制、编辑和编写 Office 脚本的基础知识。</span><span class="sxs-lookup"><span data-stu-id="c7014-104">This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span> <span data-ttu-id="c7014-105">你将录制一个脚本，以便将某些格式应用于销售记录工作表。</span><span class="sxs-lookup"><span data-stu-id="c7014-105">You'll record a script that applies some formatting to a sales record worksheet.</span></span> <span data-ttu-id="c7014-106">然后，可编辑录制的脚本以应用更多格式、创建表格并对表格进行排序。</span><span class="sxs-lookup"><span data-stu-id="c7014-106">You'll then edit the recorded script to apply more formatting, create a table, and sort that table.</span></span> <span data-ttu-id="c7014-107">这种“先记录后编辑”模式是查看 Microsoft Excel 操作作为代码的外观的重要工具。</span><span class="sxs-lookup"><span data-stu-id="c7014-107">This record-then-edit pattern is an important tool to see what your Excel actions look like as code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c7014-108">先决条件</span><span class="sxs-lookup"><span data-stu-id="c7014-108">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="c7014-109">本教程面向在 JavaScript 或 TypeScript 方面具备初级到中级知识的人员。</span><span class="sxs-lookup"><span data-stu-id="c7014-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="c7014-110">如果你不熟悉 JavaScript，建议从 [Mozilla JavaScript 教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)入手。</span><span class="sxs-lookup"><span data-stu-id="c7014-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="c7014-111">请访问 [Office 脚本代码编辑器环境](../overview/code-editor-environment.md)，以了解有关脚本环境的详细信息。</span><span class="sxs-lookup"><span data-stu-id="c7014-111">Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="c7014-112">添加数据和录制基本脚本</span><span class="sxs-lookup"><span data-stu-id="c7014-112">Add data and record a basic script</span></span>

<span data-ttu-id="c7014-113">首先，我们需要一些数据和一个小的启动脚本。</span><span class="sxs-lookup"><span data-stu-id="c7014-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="c7014-114">在 Excel 网页版中创建新的工作簿。</span><span class="sxs-lookup"><span data-stu-id="c7014-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="c7014-115">复制以下水果销售数据，并将其粘贴到工作表中，从单元格 **A1** 开始。</span><span class="sxs-lookup"><span data-stu-id="c7014-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="c7014-116">水果</span><span class="sxs-lookup"><span data-stu-id="c7014-116">Fruit</span></span> |<span data-ttu-id="c7014-117">2018 年</span><span class="sxs-lookup"><span data-stu-id="c7014-117">2018</span></span> |<span data-ttu-id="c7014-118">2019 年</span><span class="sxs-lookup"><span data-stu-id="c7014-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="c7014-119">橙子</span><span class="sxs-lookup"><span data-stu-id="c7014-119">Oranges</span></span> |<span data-ttu-id="c7014-120">1000</span><span class="sxs-lookup"><span data-stu-id="c7014-120">1000</span></span> |<span data-ttu-id="c7014-121">1200</span><span class="sxs-lookup"><span data-stu-id="c7014-121">1200</span></span> |
    |<span data-ttu-id="c7014-122">柠檬</span><span class="sxs-lookup"><span data-stu-id="c7014-122">Lemons</span></span> |<span data-ttu-id="c7014-123">800</span><span class="sxs-lookup"><span data-stu-id="c7014-123">800</span></span> |<span data-ttu-id="c7014-124">900</span><span class="sxs-lookup"><span data-stu-id="c7014-124">900</span></span> |
    |<span data-ttu-id="c7014-125">酸橙</span><span class="sxs-lookup"><span data-stu-id="c7014-125">Limes</span></span> |<span data-ttu-id="c7014-126">600</span><span class="sxs-lookup"><span data-stu-id="c7014-126">600</span></span> |<span data-ttu-id="c7014-127">500</span><span class="sxs-lookup"><span data-stu-id="c7014-127">500</span></span> |
    |<span data-ttu-id="c7014-128">葡萄柚</span><span class="sxs-lookup"><span data-stu-id="c7014-128">Grapefruits</span></span> |<span data-ttu-id="c7014-129">900</span><span class="sxs-lookup"><span data-stu-id="c7014-129">900</span></span> |<span data-ttu-id="c7014-130">700</span><span class="sxs-lookup"><span data-stu-id="c7014-130">700</span></span> |

3. <span data-ttu-id="c7014-131">打开“**自动**”选项卡。如果未看到“**自动**”选项卡，请通过按下拉箭头来检查功能区溢出。</span><span class="sxs-lookup"><span data-stu-id="c7014-131">Open the **Automate** tab. If you do not see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.</span></span>
4. <span data-ttu-id="c7014-132">按“**录制操作**”按钮。</span><span class="sxs-lookup"><span data-stu-id="c7014-132">Press the **Record Actions** button.</span></span>
5. <span data-ttu-id="c7014-133">选择单元格 **A2:C2**（“橙子”行），并将填充颜色设置为橙色。</span><span class="sxs-lookup"><span data-stu-id="c7014-133">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="c7014-134">通过按“**停止**”按钮来停止录制。</span><span class="sxs-lookup"><span data-stu-id="c7014-134">Stop the recording by pressing the **Stop** button.</span></span>
7. <span data-ttu-id="c7014-135">在“**脚本名称**”字段中填写一个便于记忆的名称。</span><span class="sxs-lookup"><span data-stu-id="c7014-135">Fill in the **Script Name** field with a memorable name.</span></span>
8. <span data-ttu-id="c7014-136">*可选：* 在“**描述**”字段中填写有意义的描述。</span><span class="sxs-lookup"><span data-stu-id="c7014-136">*Optional:* Fill in the **Description** field with a meaningful description.</span></span> <span data-ttu-id="c7014-137">这用于提供有关脚本功能的上下文。</span><span class="sxs-lookup"><span data-stu-id="c7014-137">This is used to provide context as to what the script does.</span></span> <span data-ttu-id="c7014-138">在本教程中，你可以使用“表格的颜色代码行”。</span><span class="sxs-lookup"><span data-stu-id="c7014-138">For the tutorial, you can use "Color-codes rows of a table".</span></span>

   > [!TIP]
   > <span data-ttu-id="c7014-139">稍后可以从“**脚本详细信息**”窗格编辑脚本的描述，该窗格位于代码编辑器的“**...**”菜单下。</span><span class="sxs-lookup"><span data-stu-id="c7014-139">You can edit a script's description later from the **Script Details** pane, which is located under the Code Editor's **...** menu.</span></span>

9. <span data-ttu-id="c7014-140">通过按“**保存**”按钮来保存脚本。</span><span class="sxs-lookup"><span data-stu-id="c7014-140">Save the script by pressing the **Save** button.</span></span>

    <span data-ttu-id="c7014-141">你的工作表应如下所示（不要担心颜色是否不同）:</span><span class="sxs-lookup"><span data-stu-id="c7014-141">Your worksheet should look like this (don't worry if the color is different):</span></span>

    ![水果销售数据行，其中”橙子”行突出显示为橙色。](../images/tutorial-1.png)

## <a name="edit-an-existing-script"></a><span data-ttu-id="c7014-143">编辑现有脚本</span><span class="sxs-lookup"><span data-stu-id="c7014-143">Edit an existing script</span></span>

<span data-ttu-id="c7014-144">前面的脚本将“橙子”行的颜色设置为橙色。</span><span class="sxs-lookup"><span data-stu-id="c7014-144">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="c7014-145">让我们为“柠檬”添加黄色行。</span><span class="sxs-lookup"><span data-stu-id="c7014-145">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="c7014-146">从立即打开的“**详细信息**”窗格中，按“**编辑**”按钮。</span><span class="sxs-lookup"><span data-stu-id="c7014-146">From the now-open **Details** pane, press the **Edit** button.</span></span>
2. <span data-ttu-id="c7014-147">你应该会该看到与此代码类似的内容：</span><span class="sxs-lookup"><span data-stu-id="c7014-147">You should see something similar to this code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    <span data-ttu-id="c7014-148">此代码从工作簿中获取当前工作表。</span><span class="sxs-lookup"><span data-stu-id="c7014-148">This code gets the current worksheet from the workbook.</span></span> <span data-ttu-id="c7014-149">然后，它将设置区域 **A2:C2** 的填充颜色。</span><span class="sxs-lookup"><span data-stu-id="c7014-149">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="c7014-150">区域是 Excel 网页版中的 Office 脚本的基本组成部分。</span><span class="sxs-lookup"><span data-stu-id="c7014-150">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="c7014-151">区域是一个连续的矩形单元格块，其中包含值、公式和格式。</span><span class="sxs-lookup"><span data-stu-id="c7014-151">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="c7014-152">它们是单元格的基本结构，你可以通过它们执行大多数脚本编写任务。</span><span class="sxs-lookup"><span data-stu-id="c7014-152">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

3. <span data-ttu-id="c7014-153">将以下行添加到脚本的末尾（在 `color` 设置位置和结束 `}` 之间）：</span><span class="sxs-lookup"><span data-stu-id="c7014-153">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. <span data-ttu-id="c7014-154">通过按“**运行**”来测试脚本。</span><span class="sxs-lookup"><span data-stu-id="c7014-154">Test the script by pressing **Run**.</span></span> <span data-ttu-id="c7014-155">工作簿现在应如下所示：</span><span class="sxs-lookup"><span data-stu-id="c7014-155">Your workbook should now look like this:</span></span>

    ![水果销售数据行，其中“橙子”行突出显示为橙色，而“柠檬”行则突出显示为黄色。](../images/tutorial-2.png)

## <a name="create-a-table"></a><span data-ttu-id="c7014-157">创建表格</span><span class="sxs-lookup"><span data-stu-id="c7014-157">Create a table</span></span>

<span data-ttu-id="c7014-158">让我们将此水果销售数据转换为表格。</span><span class="sxs-lookup"><span data-stu-id="c7014-158">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="c7014-159">我们将在整个过程中使用自己的脚本。</span><span class="sxs-lookup"><span data-stu-id="c7014-159">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="c7014-160">将以下行添加到脚本的末尾（在结束 `}` 之前）：</span><span class="sxs-lookup"><span data-stu-id="c7014-160">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. <span data-ttu-id="c7014-161">该调用将返回 `Table` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7014-161">That call returns a `Table` object.</span></span> <span data-ttu-id="c7014-162">让我们使用该表对数据进行排序。</span><span class="sxs-lookup"><span data-stu-id="c7014-162">Let's use that table to sort the data.</span></span> <span data-ttu-id="c7014-163">我们将根据“水果”列中的值按升序对数据进行排序。</span><span class="sxs-lookup"><span data-stu-id="c7014-163">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="c7014-164">在创建表格后添加以下行：</span><span class="sxs-lookup"><span data-stu-id="c7014-164">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="c7014-165">你的脚本应如下所示：</span><span class="sxs-lookup"><span data-stu-id="c7014-165">Your script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet12!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    <span data-ttu-id="c7014-166">表格具有 `TableSort` 对象，可通过 `Table.getSort` 方法进行访问。</span><span class="sxs-lookup"><span data-stu-id="c7014-166">Tables have a `TableSort` object, accessed through the `Table.getSort` method.</span></span> <span data-ttu-id="c7014-167">可以对该对象应用排序条件。</span><span class="sxs-lookup"><span data-stu-id="c7014-167">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="c7014-168">`apply` 方法接受 `SortField` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="c7014-168">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="c7014-169">在本示例中，我们只有一个排序条件，因此只使用一个 `SortField`。</span><span class="sxs-lookup"><span data-stu-id="c7014-169">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="c7014-170">`key: 0` 将具有排序定义值的列设置为“0”（这是表格上的第一列，在本示例中为 **A**）。</span><span class="sxs-lookup"><span data-stu-id="c7014-170">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="c7014-171">`ascending: true` 以升序（而不是降序）对数据进行排序。</span><span class="sxs-lookup"><span data-stu-id="c7014-171">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="c7014-172">运行脚本。</span><span class="sxs-lookup"><span data-stu-id="c7014-172">Run the script.</span></span> <span data-ttu-id="c7014-173">你看到的表格应类似于：</span><span class="sxs-lookup"><span data-stu-id="c7014-173">You should see a table like this:</span></span>

    ![已排序的水果销售表格。](../images/tutorial-3.png)

    > [!NOTE]
    > <span data-ttu-id="c7014-175">如果重新运行该脚本，将会收到错误消息。</span><span class="sxs-lookup"><span data-stu-id="c7014-175">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="c7014-176">这是因为不能在另一个表格的顶部创建表格。</span><span class="sxs-lookup"><span data-stu-id="c7014-176">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="c7014-177">但是，可以在其他工作表或工作簿上运行脚本。</span><span class="sxs-lookup"><span data-stu-id="c7014-177">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="c7014-178">重新运行脚本</span><span class="sxs-lookup"><span data-stu-id="c7014-178">Re-run the script</span></span>

1. <span data-ttu-id="c7014-179">在当前工作簿中创建一个新的工作表。</span><span class="sxs-lookup"><span data-stu-id="c7014-179">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="c7014-180">从教程开头复制水果数据，并将其粘贴到新的工作表中，从单元格 **A1** 开始。</span><span class="sxs-lookup"><span data-stu-id="c7014-180">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="c7014-181">运行脚本。</span><span class="sxs-lookup"><span data-stu-id="c7014-181">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="c7014-182">后续步骤</span><span class="sxs-lookup"><span data-stu-id="c7014-182">Next steps</span></span>

<span data-ttu-id="c7014-183">完成[在 Excel 网页版中使用 Office 脚本读取工作簿数据](excel-read-tutorial.md)教程。</span><span class="sxs-lookup"><span data-stu-id="c7014-183">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="c7014-184">它指导你如何使用 Office 脚本从工作簿中读取数据。</span><span class="sxs-lookup"><span data-stu-id="c7014-184">It teaches you how to read data from a workbook with an Office Script.</span></span>
