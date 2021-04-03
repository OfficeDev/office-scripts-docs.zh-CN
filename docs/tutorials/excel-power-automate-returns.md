---
title: 从脚本返回数据到自动运行 Power Automated 流
description: 本教程演示了如何通过 Power Automate 运行适用于 Excel 网页版的 Office 脚本来发送提醒电子邮件。
ms.date: 12/15/2020
localization_priority: Priority
ms.openlocfilehash: 31ba31ddbfb36f20087be6aa7d83b1b896a698d1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570528"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="a5761-103">从脚本返回数据到自动运行 Power Automated 流（预览） </span><span class="sxs-lookup"><span data-stu-id="a5761-103">Return data from a script to an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="a5761-104">本教程将教你如何从适用于 Excel 网页版的 Office 脚本中将信息作为自动 [Power Automate](https://flow.microsoft.com) 工作流的一部分返回。</span><span class="sxs-lookup"><span data-stu-id="a5761-104">This tutorial teaches you how to return information from an Office Script for Excel on the web as part of an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="a5761-105">将创建一个脚本，它可以查看时间表并与流一起发送提醒电子邮件。</span><span class="sxs-lookup"><span data-stu-id="a5761-105">You'll make a script that looks through a schedule and works with a flow to send reminder emails.</span></span> <span data-ttu-id="a5761-106">此流程将按常规计划运行，代表你提供这些提醒。</span><span class="sxs-lookup"><span data-stu-id="a5761-106">This flow will run on a regular schedule, providing these reminders on your behalf.</span></span>

> [!TIP]
> <span data-ttu-id="a5761-107">如果你不熟悉 Office 脚本，建议先查看[在 Excel 网页版中录制、编辑和创建 Office 脚本](excel-tutorial.md)教程。</span><span class="sxs-lookup"><span data-stu-id="a5761-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>
>
> <span data-ttu-id="a5761-108">如果你没有使用过 Power Automate，建议你从[手动 Power Automated 流中调用脚本](excel-power-automate-manual.md)和[在自动运行 Power Automated 流中将数据传递到脚本](excel-power-automate-trigger.md)教程开始。</span><span class="sxs-lookup"><span data-stu-id="a5761-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) and [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorials.</span></span>
>
> <span data-ttu-id="a5761-109">[Office 脚本使用 TypeScript](../overview/code-editor-environment.md)，本教程面向在 JavaScript 或 TypeScript 方面具备初级到中级知识的人员。</span><span class="sxs-lookup"><span data-stu-id="a5761-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="a5761-110">如果你不熟悉 JavaScript，建议从 [Mozilla JavaScript 教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)入手。</span><span class="sxs-lookup"><span data-stu-id="a5761-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a5761-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="a5761-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="a5761-112">准备工作簿</span><span class="sxs-lookup"><span data-stu-id="a5761-112">Prepare the workbook</span></span>

1. <span data-ttu-id="a5761-113">随时下载工作簿 - <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> 到 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="a5761-113">Download the workbook <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="a5761-114">在 Excel 网页版中打开 **on-call-rotation.xlsx**。</span><span class="sxs-lookup"><span data-stu-id="a5761-114">Open **on-call-rotation.xlsx** in Excel on the web.</span></span>

1. <span data-ttu-id="a5761-115">在表中添加行，其中包含姓名、电子邮件地址以及与当前日期重叠的开始和结束日期。</span><span class="sxs-lookup"><span data-stu-id="a5761-115">Add a row to the table with your name, email address, and start and end dates that overlap with the current date.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="a5761-116">要编写的脚本使用表中第一个匹配的条目，因此请确保你的名称位于当前周的任何行的上方。</span><span class="sxs-lookup"><span data-stu-id="a5761-116">The script you'll write uses the first matching entry in the table, so make sure your name is above any row with the current week.</span></span>

    ![Excel 电子表格中的待命轮换表屏幕截图](../images/power-automate-return-tutorial-1.png)

## <a name="create-an-office-script"></a><span data-ttu-id="a5761-118">创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="a5761-118">Create an Office Script</span></span>

1. <span data-ttu-id="a5761-119">转到“**自动**”选项卡，然后选择“**所有脚本**”。</span><span class="sxs-lookup"><span data-stu-id="a5761-119">Go to the **Automate** tab and select **All Scripts**.</span></span>

1. <span data-ttu-id="a5761-120">选择“**新建脚本**”。</span><span class="sxs-lookup"><span data-stu-id="a5761-120">Select **New Script**.</span></span>

1. <span data-ttu-id="a5761-121">将脚本命名为“**获取待命人员**”。</span><span class="sxs-lookup"><span data-stu-id="a5761-121">Name the script **Get On-Call Person**.</span></span>

1. <span data-ttu-id="a5761-122">现在应该有一个空脚本。</span><span class="sxs-lookup"><span data-stu-id="a5761-122">You should now have an empty script.</span></span> <span data-ttu-id="a5761-123">我们希望使用脚本从电子表格中获取电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="a5761-123">We want to use the script to get an email address from the spreadsheet.</span></span> <span data-ttu-id="a5761-124">更改 `main` 以返回字符串，如下所示：</span><span class="sxs-lookup"><span data-stu-id="a5761-124">Change `main` to return a string, like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. <span data-ttu-id="a5761-125">接下来，需要从表中获取所有数据。</span><span class="sxs-lookup"><span data-stu-id="a5761-125">Next, we need to get all the data from the table.</span></span> <span data-ttu-id="a5761-126">这样就可以通过脚本查看每一行。</span><span class="sxs-lookup"><span data-stu-id="a5761-126">That lets us look through each row with the script.</span></span> <span data-ttu-id="a5761-127">在 `main` 函数中添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="a5761-127">Add the following code inside the `main` function.</span></span>

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. <span data-ttu-id="a5761-128">表中的日期使用 [Excel 的日期序列号](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487)存储。</span><span class="sxs-lookup"><span data-stu-id="a5761-128">The dates in the table are stored using [Excel's date serial number](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487).</span></span> <span data-ttu-id="a5761-129">需要将这些日期转换为 JavaScript 日期以便进行比较。</span><span class="sxs-lookup"><span data-stu-id="a5761-129">We need to convert those dates to JavaScript dates in order to compare them.</span></span> <span data-ttu-id="a5761-130">将在脚本中添加帮助程序函数。</span><span class="sxs-lookup"><span data-stu-id="a5761-130">We'll add a helper function to our script.</span></span> <span data-ttu-id="a5761-131">在 `main` 函数外添加以下代码：</span><span class="sxs-lookup"><span data-stu-id="a5761-131">Add the following code outside of the `main` function:</span></span>

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. <span data-ttu-id="a5761-132">现在，我们需要弄清楚谁在待命。</span><span class="sxs-lookup"><span data-stu-id="a5761-132">Now, we need to figure out which person is on call right now.</span></span> <span data-ttu-id="a5761-133">他们的行将具有围绕当前日期的开始和结束日期。</span><span class="sxs-lookup"><span data-stu-id="a5761-133">Their row will have a start and end date surrounding the current date.</span></span> <span data-ttu-id="a5761-134">我们将编写脚本，假设一次只有一个人待命。</span><span class="sxs-lookup"><span data-stu-id="a5761-134">We'll write the script to assume only one person is on call at a time.</span></span> <span data-ttu-id="a5761-135">脚本可以返回数组来处理多个值，但现在我们将返回第一个匹配的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="a5761-135">Scripts can return arrays to handle multiple values, but for now we'll return the first matching email address.</span></span> <span data-ttu-id="a5761-136">将以下代码添加到`main` 函数末尾。</span><span class="sxs-lookup"><span data-stu-id="a5761-136">Add the following code to the end of the `main` function.</span></span>

    ```TypeScript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. <span data-ttu-id="a5761-137">最后的脚本应该如下所示：</span><span class="sxs-lookup"><span data-stu-id="a5761-137">The final script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="a5761-138">使用 Power Automate 功能创建自动工作流</span><span class="sxs-lookup"><span data-stu-id="a5761-138">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="a5761-139">登录 [Power Automate 网站](https://flow.microsoft.com)。</span><span class="sxs-lookup"><span data-stu-id="a5761-139">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

1. <span data-ttu-id="a5761-140">在屏幕左侧显示的菜单中，按 "**创建**"。</span><span class="sxs-lookup"><span data-stu-id="a5761-140">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="a5761-141">这将带你进入创建新工作流的方式列表。</span><span class="sxs-lookup"><span data-stu-id="a5761-141">This brings you to list of ways to create new workflows.</span></span>

    ![Power Automate 中的“创建”按钮。](../images/power-automate-tutorial-1.png)

1. <span data-ttu-id="a5761-143">在“**从空白开始**”部分下，选择“**计划云流**”。</span><span class="sxs-lookup"><span data-stu-id="a5761-143">Under the **Start from blank** section, select **Scheduled cloud flow**.</span></span>

    ![Power Automate 中的“已计划云流”按钮](../images/power-automate-return-tutorial-2.png)

1. <span data-ttu-id="a5761-145">现在需要为这个流程设置时间表。</span><span class="sxs-lookup"><span data-stu-id="a5761-145">Now we need to set the schedule for this flow.</span></span> <span data-ttu-id="a5761-146">从 2021 年上半年开始，电子表格在每周一都有一个新的待命任务。</span><span class="sxs-lookup"><span data-stu-id="a5761-146">Our spreadsheet has a new on-call assignment starting every Monday in the first half of 2021.</span></span> <span data-ttu-id="a5761-147">把流设置为星期一早上的首个运行的项。</span><span class="sxs-lookup"><span data-stu-id="a5761-147">Let's set the flow to run first thing Monday mornings.</span></span> <span data-ttu-id="a5761-148">使用以下选项将流配置为每周星期一运行。</span><span class="sxs-lookup"><span data-stu-id="a5761-148">Use the following options to configure the flow to run on Monday each week.</span></span>

    - <span data-ttu-id="a5761-149">**流名称**：通知待命人</span><span class="sxs-lookup"><span data-stu-id="a5761-149">**Flow name**: Notify On-Call Person</span></span>
    - <span data-ttu-id="a5761-150">**开始时间**：1/4/21 凌晨 1:00</span><span class="sxs-lookup"><span data-stu-id="a5761-150">**Starting**: 1/4/21 at 1:00am</span></span>
    - <span data-ttu-id="a5761-151">**重复间隔**：1 周</span><span class="sxs-lookup"><span data-stu-id="a5761-151">**Repeat every**: 1 Week</span></span>
    - <span data-ttu-id="a5761-152">**这些日期**：星期一</span><span class="sxs-lookup"><span data-stu-id="a5761-152">**On these days**: M</span></span>

    ![显示已计划流的指定选项的窗口](../images/power-automate-return-tutorial-3.png)

1. <span data-ttu-id="a5761-154">按“**创建**”。</span><span class="sxs-lookup"><span data-stu-id="a5761-154">Press **Create**.</span></span>

1. <span data-ttu-id="a5761-155">按 **"新建步骤"**。</span><span class="sxs-lookup"><span data-stu-id="a5761-155">Press **New step**.</span></span>

1. <span data-ttu-id="a5761-156">选择 "**标准**" 选项卡，然后选择 "**Excel Online （企业）**"。</span><span class="sxs-lookup"><span data-stu-id="a5761-156">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Power Automate 中的 Excel Online（商业版）选项](../images/power-automate-tutorial-4.png)

1. <span data-ttu-id="a5761-158">在 "**操作**"下，选择 **运行脚本（预览版）**。</span><span class="sxs-lookup"><span data-stu-id="a5761-158">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Power Automate 中的“运行脚本”（预览版）操作选项](../images/power-automate-tutorial-5.png)

1. <span data-ttu-id="a5761-160">接下来，选择要在流步骤中使用的工作簿和脚本。</span><span class="sxs-lookup"><span data-stu-id="a5761-160">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="a5761-161">使用 **on-call-rotation.xlsx** 在 OneDrive 中创建的工作簿。</span><span class="sxs-lookup"><span data-stu-id="a5761-161">Use the **on-call-rotation.xlsx** workbook you created in your OneDrive.</span></span> <span data-ttu-id="a5761-162">为 **运行脚本** 连接器指定以下设置：</span><span class="sxs-lookup"><span data-stu-id="a5761-162">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="a5761-163">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="a5761-163">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="a5761-164">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="a5761-164">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="a5761-165">**文件**: on-call-rotation.xlsx *（通过文件浏览器选择）*</span><span class="sxs-lookup"><span data-stu-id="a5761-165">**File**: on-call-rotation.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="a5761-166">**脚本**：获取待命人员</span><span class="sxs-lookup"><span data-stu-id="a5761-166">**Script**: Get On-Call Person</span></span>

    ![Power Automate 中用于运行脚本的连接器设置](../images/power-automate-return-tutorial-4.png)

1. <span data-ttu-id="a5761-168">按 **"新建步骤"**。</span><span class="sxs-lookup"><span data-stu-id="a5761-168">Press **New step**.</span></span>

1. <span data-ttu-id="a5761-169">我们将通过发送提醒邮件来结束流。</span><span class="sxs-lookup"><span data-stu-id="a5761-169">We'll end the flow by sending the reminder email.</span></span> <span data-ttu-id="a5761-170">使用连接器的搜索栏选择“**发送电子邮件 (V2)**”。</span><span class="sxs-lookup"><span data-stu-id="a5761-170">Select **Send an email (V2)** by using the connector's search bar.</span></span> <span data-ttu-id="a5761-171">使用“**新增动态内容**”控件添加脚本返回的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="a5761-171">Use the **Add dynamic content** control to add the email address returned by the script.</span></span> <span data-ttu-id="a5761-172">这将被标记为 **结果**，旁边有 Excel 图标。</span><span class="sxs-lookup"><span data-stu-id="a5761-172">This will be labelled **result** with the Excel icon next to it.</span></span> <span data-ttu-id="a5761-173">可以提供你想要的任何主题和正文。</span><span class="sxs-lookup"><span data-stu-id="a5761-173">You can provide whatever subject and body text you'd like.</span></span>

    ![在 Power Automate 中发送电子邮件的连接器设置](../images/power-automate-return-tutorial-5.png)

    > [!NOTE]
    > <span data-ttu-id="a5761-175">本教程使用 Outlook。</span><span class="sxs-lookup"><span data-stu-id="a5761-175">This tutorial uses Outlook.</span></span> <span data-ttu-id="a5761-176">可改为使用你喜欢的电子邮件服务，但某些选项可能不同。</span><span class="sxs-lookup"><span data-stu-id="a5761-176">Feel free to use your preferred email service instead, though some options may be different.</span></span>

1. <span data-ttu-id="a5761-177">按“**保存**”。</span><span class="sxs-lookup"><span data-stu-id="a5761-177">Press **Save**.</span></span>

## <a name="test-the-script-in-power-automate"></a><span data-ttu-id="a5761-178">在 Power Automate 功能中测试脚本</span><span class="sxs-lookup"><span data-stu-id="a5761-178">Test the script in Power Automate</span></span>

<span data-ttu-id="a5761-179">你的流将在每周一早上运行。</span><span class="sxs-lookup"><span data-stu-id="a5761-179">Your flow will run every Monday morning.</span></span> <span data-ttu-id="a5761-180">现在可以通过按屏幕右上角的“**测试**”按钮来测试脚本。</span><span class="sxs-lookup"><span data-stu-id="a5761-180">You can test the script now by pressing the **Test** button in the upper-right corner of the screen.</span></span> <span data-ttu-id="a5761-181">选择“**手动**”并按 **“运行测试”** 来立即运行流并测试行为。</span><span class="sxs-lookup"><span data-stu-id="a5761-181">Select **Manually** and press **Run Test** to run the flow now and test the behavior.</span></span> <span data-ttu-id="a5761-182">可能需要向 Excel 和 Outlook 授予权限才能继续。</span><span class="sxs-lookup"><span data-stu-id="a5761-182">You may need to grant permissions to Excel and Outlook to continue.</span></span>

![Power Automate 测试按钮](../images/power-automate-return-tutorial-6.png)

> [!TIP]
> <span data-ttu-id="a5761-184">如果流无法发送电子邮件，请在电子表格中仔细检查是否在表格顶部列出了当前日期范围的有效电子邮件。</span><span class="sxs-lookup"><span data-stu-id="a5761-184">If your flow fails to send an email, double-check in the spreadsheet that a valid email is listed for the current date range at the top of the table.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a5761-185">后续步骤</span><span class="sxs-lookup"><span data-stu-id="a5761-185">Next steps</span></span>

<span data-ttu-id="a5761-186">访问[使用 Power Automate 运行 Office 脚本](../develop/power-automate-integration.md)，以了解有关将 Office Script 与 Power Automate 连接的更多信息。</span><span class="sxs-lookup"><span data-stu-id="a5761-186">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="a5761-187">你还可以查看[自动任务提醒示例场景](../resources/scenarios/task-reminders.md)，以了解如何将 Office 脚本和 Power Automate 与 Team Adaptive Cards 结合使用。</span><span class="sxs-lookup"><span data-stu-id="a5761-187">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
