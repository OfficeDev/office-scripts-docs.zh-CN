---
title: Office脚本示例方案：自动任务提醒
description: 一个使用 Power Automate 自适应卡片在项目管理电子表格中自动执行任务提醒的示例。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: cf25b81ad44bbe963083f6a8346c0fd59a514305
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313979"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="104fd-103">Office脚本示例方案：自动任务提醒</span><span class="sxs-lookup"><span data-stu-id="104fd-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="104fd-104">在此方案中，你将管理项目。</span><span class="sxs-lookup"><span data-stu-id="104fd-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="104fd-105">每月使用Excel一个工作表跟踪员工的状态。</span><span class="sxs-lookup"><span data-stu-id="104fd-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="104fd-106">你经常需要提醒用户填写其状态，因此你已决定自动执行该提醒过程。</span><span class="sxs-lookup"><span data-stu-id="104fd-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="104fd-107">您将创建一个Power Automate流，以向缺少状态字段的人发送消息，然后向电子表格应用其响应。</span><span class="sxs-lookup"><span data-stu-id="104fd-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="104fd-108">为此，您将开发一对脚本来处理工作簿处理。</span><span class="sxs-lookup"><span data-stu-id="104fd-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="104fd-109">第一个脚本获取具有空白状态的人的列表，第二个脚本将状态字符串添加到右侧行。</span><span class="sxs-lookup"><span data-stu-id="104fd-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="104fd-110">你还将使用自适应卡片[Teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards)让员工直接从通知中输入其状态。</span><span class="sxs-lookup"><span data-stu-id="104fd-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="104fd-111">涵盖的脚本编写技能</span><span class="sxs-lookup"><span data-stu-id="104fd-111">Scripting skills covered</span></span>

- <span data-ttu-id="104fd-112">在 Power Automate</span><span class="sxs-lookup"><span data-stu-id="104fd-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="104fd-113">将数据传递到脚本</span><span class="sxs-lookup"><span data-stu-id="104fd-113">Pass data to scripts</span></span>
- <span data-ttu-id="104fd-114">从脚本返回数据</span><span class="sxs-lookup"><span data-stu-id="104fd-114">Return data from scripts</span></span>
- <span data-ttu-id="104fd-115">Teams自适应卡片</span><span class="sxs-lookup"><span data-stu-id="104fd-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="104fd-116">表格</span><span class="sxs-lookup"><span data-stu-id="104fd-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="104fd-117">先决条件</span><span class="sxs-lookup"><span data-stu-id="104fd-117">Prerequisites</span></span>

<span data-ttu-id="104fd-118">此方案使用[Power Automate](https://flow.microsoft.com)和[Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)。</span><span class="sxs-lookup"><span data-stu-id="104fd-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="104fd-119">你将需要与用于开发脚本的帐户关联的Office脚本。</span><span class="sxs-lookup"><span data-stu-id="104fd-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="104fd-120">若要免费访问 Microsoft 开发人员订阅以了解这些应用程序并使用这些应用程序，请考虑加入Microsoft 365[计划](https://developer.microsoft.com/microsoft-365/dev-program)。</span><span class="sxs-lookup"><span data-stu-id="104fd-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="104fd-121">设置说明</span><span class="sxs-lookup"><span data-stu-id="104fd-121">Setup instructions</span></span>

1. <span data-ttu-id="104fd-122">将<a href="task-reminders.xlsx">task-reminders.xlsx</a>下载到OneDrive。</span><span class="sxs-lookup"><span data-stu-id="104fd-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="104fd-123">在工作簿中打开Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="104fd-123">Open the workbook in Excel on the web.</span></span>

1. <span data-ttu-id="104fd-124">首先，我们需要一个脚本，用于获取电子表格中缺少状态报告的所有员工。</span><span class="sxs-lookup"><span data-stu-id="104fd-124">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="104fd-125">在" **自动化"** 选项卡下，选择 **"新建脚本** "，然后将以下脚本粘贴到编辑器中。</span><span class="sxs-lookup"><span data-stu-id="104fd-125">Under the **Automate** tab, select **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script looks for missing status reports in a project management table.
     *
     * @returns An array of Employee objects (containing their names and emails).
     */
    function main(workbook: ExcelScript.Workbook): Employee[] {
      // Get the first worksheet and the first table on that worksheet.
      let sheet = workbook.getFirstWorksheet()
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the data for the whole table.
      let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

      // Create the array of Employee objects to return.
      let people: Employee[] = [];

      // Loop through the table and check each row for completion.
      for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "") {
          // Save the email to return.
          people.push({ name: row[NAME_INDEX].toString(), email: row[EMAIL_INDEX].toString() });
        }
      }

      // Log the array to verify we're getting the right rows.
      console.log(people);

      // Return the array of Employees.
      return people;
    }

    /**
     * An interface representing an employee.
     * An array of Employees will be returned from the script
     * for the Power Automate flow.
     */
    interface Employee {
      name: string;
      email: string;
    }
    ```

1. <span data-ttu-id="104fd-126">保存名称为"获取人员" **的脚本**。</span><span class="sxs-lookup"><span data-stu-id="104fd-126">Save the script with the name **Get People**.</span></span>

1. <span data-ttu-id="104fd-127">接下来，我们需要第二个脚本处理状态报告卡，将新信息放入电子表格中。</span><span class="sxs-lookup"><span data-stu-id="104fd-127">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="104fd-128">在"代码编辑器"任务窗格中，选择" **新建脚本** "，然后将以下脚本粘贴到编辑器中。</span><span class="sxs-lookup"><span data-stu-id="104fd-128">In the Code Editor task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script applies the results of a Teams Adaptive Card about
     * a status update to a project management table.
     *
     * @param senderEmail - The email address of the employee updating their status.
     * @param statusReportResponse - The employee's status report.
     */
    function main(workbook: ExcelScript.Workbook,
      senderEmail: string,
      statusReportResponse: string) {

      // Get the first worksheet and the first table in that worksheet.
      let sheet = workbook.getFirstWorksheet();
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the range and data for the whole table.
      let bodyRange = table.getRangeBetweenHeaderAndTotal();
      let tableRowCount = bodyRange.getRowCount();
      let bodyRangeValues = bodyRange.getValues();

      // Create a flag to denote success.
      let statusAdded = false;

      // Loop through the table and check each row for a matching email address.
      for (let i = 0; i < tableRowCount && !statusAdded; i++) {
        let row = bodyRangeValues[i];

        // Check if the row's email address matches.
        if (row[EMAIL_INDEX] === senderEmail) {
          // Add the Teams Adaptive Card response to the table.
          bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
            [statusReportResponse]
          ]);
          statusAdded = true;
        }
      }

      // If successful, log the status update.
      if (statusAdded) {
        console.log(
          `Successfully added status report for ${senderEmail} containing: ${statusReportResponse}`
        );
      }
    }
    ```

1. <span data-ttu-id="104fd-129">使用名称保存状态 **保存脚本**。</span><span class="sxs-lookup"><span data-stu-id="104fd-129">Save the script with the name **Save Status**.</span></span>

1. <span data-ttu-id="104fd-130">现在，我们需要创建流。</span><span class="sxs-lookup"><span data-stu-id="104fd-130">Now, we need to create the flow.</span></span> <span data-ttu-id="104fd-131">打开[Power Automate。](https://flow.microsoft.com/)</span><span class="sxs-lookup"><span data-stu-id="104fd-131">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="104fd-132">如果之前尚未创建流，请查看我们的教程开始使用脚本和Power Automate了解基础知识[](../../tutorials/excel-power-automate-manual.md)。</span><span class="sxs-lookup"><span data-stu-id="104fd-132">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

1. <span data-ttu-id="104fd-133">创建新的即时 **流**。</span><span class="sxs-lookup"><span data-stu-id="104fd-133">Create a new **Instant flow**.</span></span>

1. <span data-ttu-id="104fd-134">选择 **"手动触发选项中的** 流"，然后选择"创建 **"。**</span><span class="sxs-lookup"><span data-stu-id="104fd-134">Choose **Manually trigger a flow** from the options and select **Create**.</span></span>

1. <span data-ttu-id="104fd-135">该流需要调用 **"获取人员** "脚本，获取具有空状态字段的所有员工。</span><span class="sxs-lookup"><span data-stu-id="104fd-135">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="104fd-136">选择 **"新建步骤**"，然后选择 **"Excel Online (Business) "。**</span><span class="sxs-lookup"><span data-stu-id="104fd-136">Select **New step**, then select **Excel Online (Business)**.</span></span> <span data-ttu-id="104fd-137">在 **操作** 下，选择 **运行脚本**。</span><span class="sxs-lookup"><span data-stu-id="104fd-137">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="104fd-138">为流步骤提供以下条目：</span><span class="sxs-lookup"><span data-stu-id="104fd-138">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="104fd-139">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="104fd-139">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="104fd-140">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="104fd-140">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="104fd-141">\**文件\*\*\*：task-reminders.xlsx (浏览器选项选择)*</span><span class="sxs-lookup"><span data-stu-id="104fd-141">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="104fd-142">**脚本**：获取人员</span><span class="sxs-lookup"><span data-stu-id="104fd-142">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="显示Power Automate运行脚本流步骤的脚本流。":::

1. <span data-ttu-id="104fd-144">接下来，该流需要处理脚本返回的数组中的每个 Employee。</span><span class="sxs-lookup"><span data-stu-id="104fd-144">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="104fd-145">选择 **"新建步骤**"，然后选择"将自适应卡片 **Teams用户并等待响应**。</span><span class="sxs-lookup"><span data-stu-id="104fd-145">Select **New step**, then choose **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

1. <span data-ttu-id="104fd-146">对于 **"收件人**"字段，**添加** 来自动态内容的电子邮件 (选定内容将具有Excel徽标) 。</span><span class="sxs-lookup"><span data-stu-id="104fd-146">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="104fd-147">添加 **电子邮件** 会导致流步骤被应用到每个块 **包围** 。</span><span class="sxs-lookup"><span data-stu-id="104fd-147">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="104fd-148">这意味着数组将按以下方法进行Power Automate。</span><span class="sxs-lookup"><span data-stu-id="104fd-148">That means the array will be iterated over by Power Automate.</span></span>

1. <span data-ttu-id="104fd-149">发送自适应卡片需要将卡片的 JSON 作为消息 **提供**。</span><span class="sxs-lookup"><span data-stu-id="104fd-149">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="104fd-150">可以使用自适应卡片 [设计器创建自定义](https://adaptivecards.io/designer/) 卡片。</span><span class="sxs-lookup"><span data-stu-id="104fd-150">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="104fd-151">对于此示例，请使用以下 JSON。</span><span class="sxs-lookup"><span data-stu-id="104fd-151">For this sample, use the following JSON.</span></span>  

    ```json
    {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "size": "Medium",
          "weight": "Bolder",
          "text": "Update your Status Report"
        },
        {
          "type": "Image",
          "altText": "",
          "url": "https://i.imgur.com/f5RcuF3.png"
        },
        {
          "type": "TextBlock",
          "text": "This is a reminder to update your status report for this month's review. You can do so right here in this card, or by adding it directly to the spreadsheet.",
          "wrap": true
        },
        {
          "type": "Input.Text",
          "placeholder": "My status report for this month is...",
          "id": "response",
          "isMultiline": true
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "Submit",
          "id": "submit"
        }
      ]
    }
    ```

1. <span data-ttu-id="104fd-152">填写其余字段，如下所示：</span><span class="sxs-lookup"><span data-stu-id="104fd-152">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="104fd-153">**更新消息**：感谢您提交状态报告。</span><span class="sxs-lookup"><span data-stu-id="104fd-153">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="104fd-154">您的响应已成功添加到电子表格。</span><span class="sxs-lookup"><span data-stu-id="104fd-154">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="104fd-155">**应更新卡片**：是</span><span class="sxs-lookup"><span data-stu-id="104fd-155">**Should update card**: Yes</span></span>

1. <span data-ttu-id="104fd-156">在"**应用到每个** 块"中，在将自适应卡片Teams **用户并等待响应** 后，选择"**添加操作"。**</span><span class="sxs-lookup"><span data-stu-id="104fd-156">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, select **Add an action**.</span></span> <span data-ttu-id="104fd-157">选择 **Excel Online (Business) 。**</span><span class="sxs-lookup"><span data-stu-id="104fd-157">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="104fd-158">在 **操作** 下，选择 **运行脚本**。</span><span class="sxs-lookup"><span data-stu-id="104fd-158">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="104fd-159">为流步骤提供以下条目：</span><span class="sxs-lookup"><span data-stu-id="104fd-159">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="104fd-160">**位置**：OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="104fd-160">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="104fd-161">**文档库**：OneDrive</span><span class="sxs-lookup"><span data-stu-id="104fd-161">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="104fd-162">\**文件\*\*\*：task-reminders.xlsx (浏览器选项选择)*</span><span class="sxs-lookup"><span data-stu-id="104fd-162">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="104fd-163">**脚本**：保存状态</span><span class="sxs-lookup"><span data-stu-id="104fd-163">**Script**: Save Status</span></span>
    - <span data-ttu-id="104fd-164">**senderEmail：** email *(dynamic content from Excel)*</span><span class="sxs-lookup"><span data-stu-id="104fd-164">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="104fd-165">**statusReportResponse：** 响应 *(动态内容Teams)*</span><span class="sxs-lookup"><span data-stu-id="104fd-165">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="显示Power Automate应用到每个步骤的流。":::

1. <span data-ttu-id="104fd-167">保存流。</span><span class="sxs-lookup"><span data-stu-id="104fd-167">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="104fd-168">运行流</span><span class="sxs-lookup"><span data-stu-id="104fd-168">Running the flow</span></span>

<span data-ttu-id="104fd-169">若要测试流，请确保任何空状态的表行都使用绑定到 Teams 帐户的电子邮件地址 (在测试) 时，应该使用自己的) 。</span><span class="sxs-lookup"><span data-stu-id="104fd-169">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span> <span data-ttu-id="104fd-170">使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。</span><span class="sxs-lookup"><span data-stu-id="104fd-170">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

<span data-ttu-id="104fd-171">你应该从用户到用户Power Automate自适应Teams。</span><span class="sxs-lookup"><span data-stu-id="104fd-171">You should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="104fd-172">在卡片中填写状态字段后，流程将继续，并更新电子表格，并包含你提供的状态。</span><span class="sxs-lookup"><span data-stu-id="104fd-172">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="104fd-173">运行流之前</span><span class="sxs-lookup"><span data-stu-id="104fd-173">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="包含一个缺少状态条目的状态报告工作表。":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="104fd-175">接收自适应卡片</span><span class="sxs-lookup"><span data-stu-id="104fd-175">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="要求员工Teams更新的自适应卡片。":::

### <a name="after-running-the-flow"></a><span data-ttu-id="104fd-177">运行流后</span><span class="sxs-lookup"><span data-stu-id="104fd-177">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="包含状态报告的工作表，现在填充了状态条目。":::
