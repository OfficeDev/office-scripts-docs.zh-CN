---
title: Office脚本示例方案：自动任务提醒
description: 在项目管理电子表格中使用Power Automate和自适应卡自动执行任务提醒的示例。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 08f3713210e83162f86d38bc8eb33d76bf8a7288
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088111"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office脚本示例方案：自动任务提醒

在此方案中，你正在管理项目。 使用Excel工作表每月跟踪员工的状态。 你通常需要提醒用户填写其状态，因此你决定自动执行该提醒过程。

你将创建一个Power Automate流，向缺少状态字段的人员发送消息，并将他们的响应应用到电子表格。 为此，你将开发一配对脚本来处理工作簿的使用。 第一个脚本获取具有空白状态的人员列表，第二个脚本将状态字符串添加到右行。 你还将使用[Teams自适应卡片](/microsoftteams/platform/task-modules-and-cards/what-are-cards)让员工直接从通知中输入其状态。

## <a name="scripting-skills-covered"></a>所涵盖的脚本技能

- 在Power Automate中创建流
- 将数据传递给脚本
- 从脚本返回数据
- Teams自适应卡片
- 表格

## <a name="prerequisites"></a>先决条件

此方案使用[Power Automate](https://flow.microsoft.com)和[Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)。 需要与用于开发Office脚本的帐户相关联。 若要免费访问 Microsoft 开发人员订阅以了解和使用这些应用程序，请考虑加入[Microsoft 365开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)。

## <a name="setup-instructions"></a>设置说明

1. <a href="task-reminders.xlsx"> 将task-reminders.xlsx</a>下载到OneDrive。

1. 在Excel web 版中打开工作簿。

1. 首先，我们需要一个脚本来获取电子表格中缺少状态报告的所有员工。 在 **“自动执行”** 选项卡下，选择 **“新建脚本** ”并将以下脚本粘贴到编辑器中。

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

1. 保存名为 **“获取人员**”的脚本。

1. 接下来，我们需要第二个脚本来处理状态报表卡，并将新信息放入电子表格中。 在“代码编辑器”任务窗格中，选择 **“新建脚本** ”并将以下脚本粘贴到编辑器中。

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

1. 保存名为 **“保存状态**”的脚本。

1. 现在，我们需要创建流。 打开[Power Automate](https://flow.microsoft.com/)。

    > [!TIP]
    > 如果之前尚未创建流，请查看本教程["开始"菜单使用带Power Automate的脚本](../../tutorials/excel-power-automate-manual.md)来了解基础知识。

1. 创建新的 **即时流**。

1. 从选项 **中选择“手动触发流** ”，然后选择 **“创建**”。

1. 该流需要调用 **Get People** 脚本，以获取具有空状态字段的所有员工。 选择 **“新建步骤**”，然后选择 **“联机 (业务) Excel**。 在 **操作** 下，选择 **运行脚本**。 为流步骤提供以下条目：

    - **位置**：OneDrive for Business
    - **文档库**：OneDrive
    - **文件**：通过 *文件浏览器) 选择task-reminders.xlsx (*
    - **脚本**：获取人员

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="显示第一个运行脚本流步骤的Power Automate流。":::

1. 接下来，流需要处理脚本返回的数组中的每个 Employee。 选择 **“新建”步骤**，然后选择“**向Teams用户发布自适应卡片并等待响应**。

1. 对于 **“收件人”** 字段，添加来自动态内容 **的电子邮件** (所选内容) 包含Excel徽标。 添加 **电子邮件** 会导致流步骤被 **应用到每个** 块所包围。 这意味着数组将通过Power Automate进行迭代。

1. 发送自适应卡片需要将卡片的 [JSON](https://www.w3schools.com/whatis/whatis_json.asp) 作为 **消息** 提供。 可以使用 [自适应卡片设计器](https://adaptivecards.io/designer/) 创建自定义卡片。 对于此示例，请使用以下 JSON。  

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

1. 按如下所示填写剩余字段：

    - **更新消息**：感谢你提交状态报告。 响应已成功添加到电子表格中。
    - **应更新卡** 片：是

1. 在 **“应用到每个** 块”中，在向 **Teams用户发布自适应卡片并等待响应** 后，选择 **“添加操作**”。 选择 **Excel联机 (业务)**。 在 **操作** 下，选择 **运行脚本**。 为流步骤提供以下条目：

    - **位置**：OneDrive for Business
    - **文档库**：OneDrive
    - **文件**：通过 *文件浏览器) 选择task-reminders.xlsx (*
    - **脚本**：保存状态
    - **senderEmail**：*从Excel) 发送 (动态内容* 的电子邮件
    - **statusReportResponse**：响应 *(来自Teams) 的动态内容*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="显示应用到每个步骤的Power Automate流。":::

1. 保存流。

## <a name="running-the-flow"></a>运行流

若要测试流，请确保任何具有空白状态的表行都使用绑定到Teams帐户的电子邮件地址 (在测试) 时，您可能应使用自己的电子邮件地址。 使用流编辑器页上的 **“测试** ”按钮，或通过“ **我的流** ”选项卡运行流。出现提示时，请务必允许访问。

应从Power Automate到Teams收到自适应卡片。 在卡片中填写状态字段后，流将继续使用所提供的状态更新电子表格。

### <a name="before-running-the-flow"></a>运行流之前

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="一个包含一个缺失状态条目的状态报告的工作表。":::

### <a name="receiving-the-adaptive-card"></a>接收自适应卡片

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Teams中向员工询问状态更新的自适应卡片。":::

### <a name="after-running-the-flow"></a>运行流后

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="包含状态报表的工作表，其中包含现在已填充的状态条目。":::
