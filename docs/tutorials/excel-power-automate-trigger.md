---
title: 自动通过 Power Automate 运行脚本
description: 本教程介绍了如何使用自动外部触发器（通过 Outlook 接收邮件）在 web 上运行 Excel 的 Office 脚本。
ms.date: 06/29/2020
localization_priority: Priority
ms.openlocfilehash: a750197d6b5ae770ad7d2e17b3ee00dc65ee8875
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043417"
---
# <a name="automatically-run-scripts-with-power-automate-preview"></a>自动运行带电自动化的脚本（预览）

本教程向您介绍如何通过自动[功能自动执行](https://flow.microsoft.com)工作流，在 web 上使用适用于 Excel 的 Office 脚本。 您的脚本将在您每次收到电子邮件时自动运行，并在 Excel 工作簿中记录来自电子邮件的信息。

## <a name="prerequisites"></a>先决条件

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> 本教程假定您已完成[使用 Power 自动化教程的 web 上的 Excel 中运行 Office 脚本](excel-power-automate-manual.md)。

## <a name="prepare-the-workbook"></a>准备工作簿

Power 自动执行无法使用[相对引用](../develop/power-automate-integration.md#avoid-using-relative-references) `Workbook.getActiveWorksheet` ，如访问工作簿组件。 因此，我们需要具有一致的名称的工作簿和工作表，以使电源自动参考。

1. 创建一个名为**MyWorkbook**的新工作簿。

2. 转到 "**自动**" 选项卡，然后选择 "**代码编辑器**"。

3. 选择 "**新建脚本**"。

4. 将现有代码替换为以下脚本，然后按 "**运行**"。 这将使用一致的工作表、表和数据透视表名称设置工作簿。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("SubjectPivot");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script-for-your-automated-workflow"></a>为自动工作流创建 Office 脚本

我们来创建一个脚本，用于记录来自电子邮件的信息。 我们想知道我们在一周中的哪几天收到最多的邮件，以及有多少唯一发件人发送该邮件。 我们的工作簿包含一个包含**日期**、**星期几**、**电子邮件地址**和**主题**列的表格。 我们的工作表还有一个透视表，该数据透视表在**星期**和**电子邮件地址**（这些行是行分层）中进行透视。 "唯一"**主题**的计数是要显示的聚合信息（数据层次结构）。 我们将在更新电子邮件表后，脚本刷新数据透视表。

1. 在**代码编辑器**中，选择 "**新建脚本**"。

2. 我们稍后将在本教程中创建的流程将发送有关收到的每封电子邮件的脚本信息。 脚本需要通过函数中的参数接受该输入 `main` 。 将默认脚本替换为以下脚本：

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. 该脚本需要访问工作簿的表和数据透视表。 将下面的代码添加到脚本正文中的开头 `{` ：

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. `dateReceived`参数的类型为 `string` 。 让我们将它转换为[ `Date` 对象](../develop/javascript-objects.md#date)，以便我们可以轻松获取一周中的一天。 执行此操作后，我们需要将日的数字值映射到更易于阅读的版本。 将以下代码添加到脚本末尾，在结束之前 `}` ：

    ```TypeScript
    // Parse the received date string.
    let date = new Date(dateReceived);

    // Convert number representing the day of the week into the name of the day.
    let dayText : string;
    switch (date.getDay()) {
      case 0:
        dayText = "Sunday";
        break;
      case 1:
        dayText = "Monday";
        break;
      case 2:
        dayText = "Tuesday";
        break;
      case 3:
        dayText = "Wednesday";
        break;
      case 4:
        dayText = "Thursday";
        break;
      case 5:
        dayText = "Friday";
        break;
      default:
        dayText = "Saturday";
        break;
    }
    ```

5. 该 `subject` 字符串可能包含 "RE：" 答复标记。 让我们从字符串中删除该对象，以便同一线程中的电子邮件具有相同的表格主题。 将以下代码添加到脚本末尾，在结束之前 `}` ：

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. 现在，已经根据需要对电子邮件数据进行了格式化，我们将行添加到电子邮件表中。 将以下代码添加到脚本末尾，在结束之前 `}` ：

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. 最后，让我们确保刷新数据透视表。 将以下代码添加到脚本末尾，在结束之前 `}` ：

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. 重命名脚本**记录电子邮件**，然后按 "**保存脚本**"。

现在，你的脚本已准备就绪，可以自动执行工作流。 它应类似于下面的脚本：

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Pivot");
  let pivotTable = pivotTableWorksheet.getPivotTable("SubjectPivot");

  // Parse the received date string.
  let date = new Date(dateReceived);

  // Convert number representing the day of the week into the name of the day.
  let dayText: string;
  switch (date.getDay()) {
    case 0:
      dayText = "Sunday";
      break;
    case 1:
      dayText = "Monday";
      break;
    case 2:
      dayText = "Tuesday";
      break;
    case 3:
      dayText = "Wednesday";
      break;
    case 4:
      dayText = "Thursday";
      break;
    case 5:
      dayText = "Friday";
      break;
    default:
      dayText = "Saturday";
      break;
  }

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayText, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a>使用 Power 自动化创建自动工作流

1. 登录到[Power 自动预览网站](https://flow.microsoft.com)。

2. 在屏幕左侧显示的菜单中，按 "**创建**"。 这将向你显示创建新工作流的方式列表。

    !["增强电源" 中的 "创建" 按钮。](../images/power-automate-tutorial-1.png)

3. 在 "**从空白开始**" 部分，选择 "**自动流**"。 这将创建由事件（如接收电子邮件）触发的工作流。

    ![自动执行电源中的自动流选项。](../images/power-automate-params-tutorial-1.png)

4. 在出现的对话框窗口中，在 "**流名称**" 文本框中输入流的名称。 然后，从 "**选择您的流的触发**" 下的选项列表中选择 "**新电子邮件到达时**"。 您可能需要使用 "搜索" 框搜索选项。 最后，按 "**创建**"。

    ![在 "电源自动执行" 部分的 "构建自动流" 窗口显示 "新电子邮件到达" 选项。](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > 本教程使用 Outlook。 你可以改用首选的电子邮件服务，但某些选项可能会有所不同。

5. 按 "**新建步骤**"。

6. 选择 "**标准**" 选项卡，然后选择 " **Excel Online （企业）**"。

    ![Excel Online （业务）的 "电源自动" 选项。](../images/power-automate-tutorial-4.png)

7. 在 "**操作**" 下，选择 "**运行脚本（预览）**"。

    !["运行脚本（预览）" 的 "电源自动操作" 选项。](../images/power-automate-tutorial-5.png)

8. 为 "**运行脚本**" 连接器指定以下设置：

    - **位置**： OneDrive for business
    - **文档库**： OneDrive
    - **文件**： MyWorkbook.xlsx
    - **脚本**：记录电子邮件
    - **发件人**： from *（来自 Outlook 的动态内容）*
    - **dateReceived**：接收时间 *（来自 Outlook 的动态内容）*
    - **subject**： subject *（来自 Outlook 的动态内容）*

    *请注意，只有在选择脚本后，才会显示脚本的参数。*

    !["运行脚本（预览）" 的 "电源自动操作" 选项。](../images/power-automate-params-tutorial-3.png)

9. 按 "**保存**"。

您的流现已启用。 它将在您每次通过 Outlook 收到电子邮件时自动运行您的脚本。

## <a name="manage-the-script-in-power-automate"></a>管理自动电源中的脚本

1. 从 "主电自动" 页面中，选择 "**我的流**"。

    !["电源自动" 中的 "我的流" 按钮。](../images/power-automate-tutorial-7.png)

2. 选择您的流。 你可以在此处查看运行历史记录。 您可以刷新页面或按 "刷新**所有运行**" 按钮以更新历史记录。 在收到电子邮件后，流将立即触发。 通过发送自己的邮件测试流。

当触发流并成功运行脚本时，您应该会看到工作簿的表和数据透视表更新。

![流运行几次后的电子邮件表。](../images/power-automate-params-tutorial-4.png)

![在流运行几次之后的数据透视表。](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a>后续步骤

访问 "[使用 power 自动运行 Office 脚本](../develop/power-automate-integration.md)"，以详细了解如何通过 power 自动化连接 office 脚本。

您还可以查看[自动任务提醒示例方案](../resources/scenarios/task-reminders.md)，以了解如何使用工作组自适应卡片将 Office 脚本和电力自动化组合在一起。
