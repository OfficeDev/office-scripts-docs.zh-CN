---
title: 通过手动 Power Automate 流呼叫脚本
description: 有关通过手动触发器在 Power Automate 中使用 Office 脚本的教程。
ms.date: 06/29/2021
localization_priority: Priority
ms.openlocfilehash: 1a8b9659ec6f6354d583496ba0f3e94d4a13c01b
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313993"
---
# <a name="call-scripts-from-a-manual-power-automate-flow"></a><span data-ttu-id="91045-103">通过手动 Power Automate 流呼叫脚本</span><span class="sxs-lookup"><span data-stu-id="91045-103">Call scripts from a manual Power Automate flow</span></span>

<span data-ttu-id="91045-104">本教程将指导你如何通过 [Power Automate](https://flow.microsoft.com) 在 web 上运行 Office Script for Excel。</span><span class="sxs-lookup"><span data-stu-id="91045-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="91045-105">你将创建一个脚本，以当前时间更新两个单元格的值。</span><span class="sxs-lookup"><span data-stu-id="91045-105">You'll make a script that updates the values of two cells with the current time.</span></span> <span data-ttu-id="91045-106">然后，你可以将该脚本连接到手动触发的 Power Automate 流，以便每当选择 Power Automate 中的按钮时，脚本就会运行。</span><span class="sxs-lookup"><span data-stu-id="91045-106">You'll then connect that script to a manually triggered Power Automate flow, so that the script is run whenever a button in Power Automate is selected.</span></span> <span data-ttu-id="91045-107">了解基本模式后，可展开流以包括其他应用程序，并自动执行更多日常工作流。</span><span class="sxs-lookup"><span data-stu-id="91045-107">Once you understand the basic pattern, you can expand the flow to include other applications and automate more of your daily workflow.</span></span>

> [!TIP]
> <span data-ttu-id="91045-108">如果你不熟悉 Office 脚本，建议先查看[在 Excel 网页版中录制、编辑和创建 Office 脚本](excel-tutorial.md)教程。</span><span class="sxs-lookup"><span data-stu-id="91045-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="91045-109">[Office 脚本使用 TypeScript](../overview/code-editor-environment.md)，本教程面向在 JavaScript 或 TypeScript 方面具备初级到中级知识的人员。</span><span class="sxs-lookup"><span data-stu-id="91045-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="91045-110">如果你不熟悉 JavaScript，建议从 [Mozilla JavaScript 教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)入手。</span><span class="sxs-lookup"><span data-stu-id="91045-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="91045-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="91045-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="91045-112">准备工作簿</span><span class="sxs-lookup"><span data-stu-id="91045-112">Prepare the workbook</span></span>

<span data-ttu-id="91045-113">Power Automate 不应使用`Workbook.getActiveWorksheet`之类的[相对引用](../testing/power-automate-troubleshooting.md#avoid-relative-references)访问工作簿组件。</span><span class="sxs-lookup"><span data-stu-id="91045-113">Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="91045-114">因此，我们需要一个具有 Power Automate 可以引用的一致名称的工作簿和工作表。</span><span class="sxs-lookup"><span data-stu-id="91045-114">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="91045-115">创建名为 **MyWorkbook** 的新工作簿。</span><span class="sxs-lookup"><span data-stu-id="91045-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="91045-116">在 **MyWorkbook** 工作簿中，创建一个名为 **TutorialWorksheet** 的工作表。</span><span class="sxs-lookup"><span data-stu-id="91045-116">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="91045-117">创建 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="91045-117">Create an Office Script</span></span>

1. <span data-ttu-id="91045-118">转到“**自动**”选项卡，然后选择“**所有脚本**”。</span><span class="sxs-lookup"><span data-stu-id="91045-118">Go to the **Automate** tab and select **All Scripts**.</span></span>

2. <span data-ttu-id="91045-119">选择 "**New Script**"。</span><span class="sxs-lookup"><span data-stu-id="91045-119">Select **New Script**.</span></span>

3. <span data-ttu-id="91045-120">将默认脚本替换为以下脚本。</span><span class="sxs-lookup"><span data-stu-id="91045-120">Replace the default script with the following script.</span></span> <span data-ttu-id="91045-121">此脚本将当前日期和时间添加到 **TutorialWorksheet** 工作表的前两个单元格。</span><span class="sxs-lookup"><span data-stu-id="91045-121">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. <span data-ttu-id="91045-122">将脚本重命名为 "**设置日期和时间**"。</span><span class="sxs-lookup"><span data-stu-id="91045-122">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="91045-123">选择脚本名以进行更改。</span><span class="sxs-lookup"><span data-stu-id="91045-123">Select the script name to change it.</span></span>

5. <span data-ttu-id="91045-124">通过选择“**保存脚本**”来保存脚本。</span><span class="sxs-lookup"><span data-stu-id="91045-124">Save the script by selecting **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="91045-125">使用 Power Automate 功能创建自动工作流</span><span class="sxs-lookup"><span data-stu-id="91045-125">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="91045-126">登录 [Power Automate 网站](https://flow.microsoft.com)。</span><span class="sxs-lookup"><span data-stu-id="91045-126">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="91045-127">在屏幕左侧显示的菜单中，选择“**创建**”。</span><span class="sxs-lookup"><span data-stu-id="91045-127">In the menu that's displayed on the left side of the screen, select **Create**.</span></span> <span data-ttu-id="91045-128">这将带你进入创建新工作流的方式列表。</span><span class="sxs-lookup"><span data-stu-id="91045-128">This brings you to list of ways to create new workflows.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Power Automate“创建”按钮。":::

3. <span data-ttu-id="91045-130">在 **从空白开始** 部分中，选择 **即时流**。</span><span class="sxs-lookup"><span data-stu-id="91045-130">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="91045-131">这将创建手动激活的工作流。</span><span class="sxs-lookup"><span data-stu-id="91045-131">This creates a manually activated workflow.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="用于创建新工作流的 Power Automate 即时流选项。":::

4. <span data-ttu-id="91045-133">在出现的对话框窗口中，在“**流名称**”文本框中输入流的名称，从“**选择如何触发流**”下的选项列表中，选择“**手动触发流**”，然后选择“**创建**”。</span><span class="sxs-lookup"><span data-stu-id="91045-133">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and then select **Create**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="Power Automate &quot;手动触发流&quot;选项。":::

    <span data-ttu-id="91045-135">请注意，手动触发流仅是许多类型流中的一种。</span><span class="sxs-lookup"><span data-stu-id="91045-135">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="91045-136">在下一个教程中，你将创建收到电子邮件时自动运行的流程。</span><span class="sxs-lookup"><span data-stu-id="91045-136">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="91045-137">选择“**新建步骤**”。</span><span class="sxs-lookup"><span data-stu-id="91045-137">Select **New step**.</span></span>

6. <span data-ttu-id="91045-138">选择 "**标准**" 选项卡，然后选择 "**Excel Online （企业）**"。</span><span class="sxs-lookup"><span data-stu-id="91045-138">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text=" Power Automate 中的 Excel Online (商业版)选项。":::

7. <span data-ttu-id="91045-140">在 **操作** 下，选择 **运行脚本**。</span><span class="sxs-lookup"><span data-stu-id="91045-140">Under **Actions**, select **Run script**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text=" Power Automate 中的运行脚本操作选项。":::

8. <span data-ttu-id="91045-142">接下来，选择要在流步骤中使用的工作簿和脚本。</span><span class="sxs-lookup"><span data-stu-id="91045-142">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="91045-143">对于本教程，你将使用在 OneDrive 中创建的工作簿，但可以在 OneDrive 或 SharePoint 网站中使用任何工作簿。</span><span class="sxs-lookup"><span data-stu-id="91045-143">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="91045-144">为 **运行脚本** 连接器指定以下设置：</span><span class="sxs-lookup"><span data-stu-id="91045-144">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="91045-145">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="91045-145">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="91045-146">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="91045-146">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="91045-147">**文件**: MyWorkbook.xlsx *（通过文件浏览器选择）*</span><span class="sxs-lookup"><span data-stu-id="91045-147">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="91045-148">**脚本**：设置日期和时间</span><span class="sxs-lookup"><span data-stu-id="91045-148">**Script**: Set date and time</span></span>

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="用于运行脚本的 Power Automate 连接器设置。":::

9. <span data-ttu-id="91045-150">选择“**保存**”。</span><span class="sxs-lookup"><span data-stu-id="91045-150">Select **Save**.</span></span>

<span data-ttu-id="91045-151">现在，你的流程可以通过 Power Automate 运行。</span><span class="sxs-lookup"><span data-stu-id="91045-151">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="91045-152">可使用流编辑器中的 "**测试**" 按钮对其进行测试，或按照其余教程步骤运行流集合中的流程。</span><span class="sxs-lookup"><span data-stu-id="91045-152">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="91045-153">通过 Power Automate 运行脚本</span><span class="sxs-lookup"><span data-stu-id="91045-153">Run the script through Power Automate</span></span>

1. <span data-ttu-id="91045-154">在 Power Automate 主页面上，选择 **我的流**。</span><span class="sxs-lookup"><span data-stu-id="91045-154">From the main Power Automate page, select **My flows**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text=" Power Automate 中的&quot;我的流程&quot;按钮。":::

2. <span data-ttu-id="91045-156">从 "**我的流**" 选项卡中显示的流列表中选择 **我的教程流**。这将显示之前创建的流程的详细信息。</span><span class="sxs-lookup"><span data-stu-id="91045-156">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="91045-157">选择“**运行**”。</span><span class="sxs-lookup"><span data-stu-id="91045-157">Select **Run**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text=" Power Automate 中的“运行”按钮。":::

4. <span data-ttu-id="91045-159">将显示用于运行流的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="91045-159">A task pane will appear for running the flow.</span></span> <span data-ttu-id="91045-160">如果系统要求 **登录** 到 Excel Online，请通过选择“**继续**”来执行操作。</span><span class="sxs-lookup"><span data-stu-id="91045-160">If you are asked to **Sign in** to Excel Online, do so by selecting **Continue**.</span></span>

5. <span data-ttu-id="91045-161">选择“**运行流**”。</span><span class="sxs-lookup"><span data-stu-id="91045-161">Select **Run flow**.</span></span> <span data-ttu-id="91045-162">此时将运行流，该流将运行相关的 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="91045-162">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="91045-163">选择“**完成**”。</span><span class="sxs-lookup"><span data-stu-id="91045-163">Select **Done**.</span></span> <span data-ttu-id="91045-164">你应该看到 **运行** 部分进行了相应的更新。</span><span class="sxs-lookup"><span data-stu-id="91045-164">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="91045-165">刷新页面，查看 Power Automate 的结果。</span><span class="sxs-lookup"><span data-stu-id="91045-165">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="91045-166">如果成功，请转到工作簿查看已更新的单元格。</span><span class="sxs-lookup"><span data-stu-id="91045-166">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="91045-167">如果失败，请验证流的设置并再次运行。</span><span class="sxs-lookup"><span data-stu-id="91045-167">If it failed, verify the flow's settings and run it a second time.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text=" Power Automate 输出显示成功流运行。":::

## <a name="next-steps"></a><span data-ttu-id="91045-169">后续步骤</span><span class="sxs-lookup"><span data-stu-id="91045-169">Next steps</span></span>

<span data-ttu-id="91045-170">完成[将数据传递到自动运行的 Power Automate 流中的脚本](excel-power-automate-trigger.md)教程。</span><span class="sxs-lookup"><span data-stu-id="91045-170">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="91045-171">它教你如何将数据从工作流服务传递到你的 Office 脚本，并在发生特定事件时运行 Power Automate 流。</span><span class="sxs-lookup"><span data-stu-id="91045-171">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
