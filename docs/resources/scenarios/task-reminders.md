---
title: Office 脚本示例方案：自动任务提醒
description: 使用 Power Automate 和自适应卡片在项目管理电子表格中自动执行任务提醒的示例。
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: a229a06e9f1f9118d57dadac8864bbc7eae7315b
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755152"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office 脚本示例方案：自动任务提醒

在此方案中，你将管理项目。 每月使用 Excel 工作表跟踪员工的状态。 你经常需要提醒用户填写其状态，因此你已决定自动执行该提醒过程。

你将创建 Power Automate 流，以向缺少状态字段的人发送消息，然后对电子表格应用他们的响应。 为此，您将开发一对脚本来处理工作簿处理。 第一个脚本获取具有空白状态的人的列表，第二个脚本将状态字符串添加到右侧行。 你还将使用 [Teams 自适应卡片](/microsoftteams/platform/task-modules-and-cards/what-are-cards) 让员工直接从通知中输入其状态。

## <a name="scripting-skills-covered"></a>涵盖的脚本编写技能

- 在 Power Automate 中创建流
- 将数据传递到脚本
- 从脚本返回数据
- Teams 自适应卡片
- 表格

## <a name="prerequisites"></a>先决条件

此方案使用[Power Automate](https://flow.microsoft.com)和 Microsoft [Teams。](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software) 你将需要与用于开发 Office 脚本的帐户关联的两者。 若要免费访问 Microsoft 开发人员订阅以了解这些应用程序并使用这些应用程序，请考虑加入 [Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)。

## <a name="setup-instructions"></a>设置说明

1. 将 <a href="task-reminders.xlsx">task-reminders.xlsx</a> 下载到 OneDrive。

2. 在 Excel 网页中打开工作簿。

3. 在"**自动化"选项卡** 下，打开 **"所有脚本"。**

4. 首先，我们需要一个脚本，用于获取电子表格中缺少状态报告的所有员工。 在" **代码编辑器"** 任务窗格中，按 **"新建脚本** "，然后将以下脚本粘贴到编辑器中。

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

5. 保存名称为"获取人员" **的脚本**。

6. 接下来，我们需要第二个脚本处理状态报告卡，将新信息放入电子表格中。 在" **代码编辑器"** 任务窗格中，按 **"新建脚本** "，然后将以下脚本粘贴到编辑器中。

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

7. 使用名称保存状态 **保存脚本**。

8. 现在，我们需要创建流。 打开 [Power Automate](https://flow.microsoft.com/)。

    > [!TIP]
    > 如果之前尚未创建流，请查看我们的教程开始使用 [Power Automate](../../tutorials/excel-power-automate-manual.md) 脚本了解基础知识。

9. 创建新的即时 **流**。

10. 从 **选项中选择"手动触发** 流"，然后按"创建 **"。**

11. 该流需要调用 **"获取人员** "脚本，获取具有空状态字段的所有员工。 按 **"新建步骤**"，然后选择 **"Excel Online (Business) "。** 在 "**操作**"下，选择 **运行脚本（预览版）**。 为流步骤提供以下条目：

    - **位置**：OneDrive for Business
    - **文档库**：OneDrive
    - **文件***：task-reminders.xlsx (浏览器选项选择)*
    - **脚本**：获取人员

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Power Automate 流显示第一个运行脚本流步骤。":::

12. 接下来，该流需要处理脚本返回的数组中的每个 Employee。 按 **"新建步骤** "，然后选择"向 Teams 用户发布自适应卡片 **"并等待响应**。

13. 对于 **"收件人**"字段，添加来自动态内容的电子邮件 (所选内容将具有 Excel 徽标，并) 。 添加 **电子邮件** 会导致流步骤被应用到每个块 **包围** 。 这意味着该数组将由 Power Automate 进行重复。

14. 发送自适应卡片需要将卡片的 JSON 作为消息 **提供**。 可以使用自适应卡片 [设计器创建自定义](https://adaptivecards.io/designer/) 卡片。 对于此示例，请使用以下 JSON。  

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

15. 填写其余字段，如下所示：

    - **更新消息**：感谢您提交状态报告。 您的响应已成功添加到电子表格。
    - **应更新卡片**：是

16. 在 **"应用到每个** 块"中，在将自适应卡片发布给 Teams 用户并 **等待响应** 后，按 **"添加操作"。** 选择 **Excel Online (Business)**。 在 "**操作**"下，选择 **运行脚本（预览版）**。 为流步骤提供以下条目：

    - **位置**：OneDrive for Business
    - **文档库**：OneDrive
    - **文件***：task-reminders.xlsx (浏览器选项选择)*
    - **脚本**：保存状态
    - **senderEmail：** email *(dynamic content from Excel)*
    - **statusReportResponse：** 来自 *Teams (动态内容的响应)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="显示适用于每个步骤的 Power Automate 流。":::

17. 保存流。

## <a name="running-the-flow"></a>运行流

若要测试流，请确保任何空状态的表行使用绑定到 Teams 帐户的电子邮件地址 (在测试) 时，你可能应该使用自己的) 。

可以从流设计器 **中选择"测试** "，也可以从"我的流"页 **运行** 流。 启动流程并接受使用所需连接后，你应该通过 Teams 从 Power Automate 收到自适应卡片。 在卡片中填写状态字段后，流程将继续，并更新电子表格，并包含你提供的状态。

### <a name="before-running-the-flow"></a>运行流之前

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="包含一个缺少状态条目的状态报告工作表。":::

### <a name="receiving-the-adaptive-card"></a>接收自适应卡片

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Teams 中的自适应卡片，要求员工提供状态更新。":::

### <a name="after-running-the-flow"></a>运行流后

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="包含状态报告的工作表，现在填充了状态条目。":::
