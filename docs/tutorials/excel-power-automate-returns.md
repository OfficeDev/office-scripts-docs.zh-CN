---
title: 从脚本返回数据到自动运行 Power Automated 流
description: 本教程演示了如何通过 Power Automate 运行适用于 Excel 网页版的 Office 脚本来发送提醒电子邮件。
ms.date: 06/29/2021
localization_priority: Priority
ms.openlocfilehash: 5c021a4d146fc2c0f13c7d7ab560107a1bf5c335
ms.sourcegitcommit: 6654aeae8a3ee2af84b4d4c4d8ff45b360a303eb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2021
ms.locfileid: "58862220"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow"></a>从脚本返回数据到自动运行 Power Automated 流

本教程将教你如何从适用于 Excel 网页版的 Office 脚本中将信息作为自动 [Power Automate](https://flow.microsoft.com) 工作流的一部分返回。 将创建一个脚本，它可以查看时间表并与流一起发送提醒电子邮件。 此流程将按常规计划运行，代表你提供这些提醒。

> [!TIP]
> 如果你不熟悉 Office 脚本，建议先查看[在 Excel 网页版中录制、编辑和创建 Office 脚本](excel-tutorial.md)教程。
>
> 如果你没有使用过 Power Automate，建议你从[手动 Power Automated 流中调用脚本](excel-power-automate-manual.md)和[在自动运行 Power Automated 流中将数据传递到脚本](excel-power-automate-trigger.md)教程开始。
>
> [Office 脚本使用 TypeScript](../overview/code-editor-environment.md)，本教程面向在 JavaScript 或 TypeScript 方面具备初级到中级知识的人员。 如果你不熟悉 JavaScript，建议从 [Mozilla JavaScript 教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)入手。

## <a name="prerequisites"></a>先决条件

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>准备工作簿

1. 随时下载工作簿 - <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> 到 OneDrive。

1. 在 Excel 网页版中打开 **on-call-rotation.xlsx**。

1. 在表中添加行，其中包含姓名、电子邮件地址以及与当前日期重叠的开始和结束日期。

    > [!IMPORTANT]
    > 要编写的脚本使用表中第一个匹配的条目，因此请确保你的名称位于当前周的任何行的上方。

    :::image type="content" source="../images/power-automate-return-tutorial-1.png" alt-text="包含呼叫旋转表格数据的工作表。":::

## <a name="create-an-office-script"></a>创建 Office 脚本

1. 转到“**自动**”选项卡，然后选择“**所有脚本**”。

1. 选择“**新建脚本**”。

1. 将脚本命名为“**获取待命人员**”。

1. 现在应该有一个空脚本。 我们希望使用脚本从电子表格中获取电子邮件地址。 更改 `main` 以返回字符串，如下所示：

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. 接下来，需要从表中获取所有数据。 这样就可以通过脚本查看每一行。 在 `main` 函数中添加以下代码。

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. 表中的日期使用 [Excel 的日期序列号](https://support.microsoft.com/office/e7fe7167-48a9-4b96-bb53-5612a800b487)存储。 需要将这些日期转换为 JavaScript 日期以便进行比较。 将在脚本中添加帮助程序函数。 在 `main` 函数外添加以下代码：

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. 现在，我们需要弄清楚谁在待命。 他们的行将具有围绕当前日期的开始和结束日期。 我们将编写脚本，假设一次只有一个人待命。 脚本可以返回数组来处理多个值，但现在我们将返回第一个匹配的电子邮件地址。 将以下代码添加到`main` 函数末尾。

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

1. 最后的脚本应该如下所示：

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

## <a name="create-an-automated-workflow-with-power-automate"></a>使用 Power Automate 功能创建自动工作流

1. 登录 [Power Automate 网站](https://flow.microsoft.com)。

1. 在屏幕左侧显示的菜单中，选择“**创建**”。 这将带你进入创建新工作流的方式列表。

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text=" Power Automate 中的&quot;创建&quot;按钮。":::

1. 在“**从空白开始**”部分下，选择“**计划云流**”。

    :::image type="content" source="../images/power-automate-return-tutorial-2.png" alt-text=" Power Automate 中的&quot;计划云流&quot;按钮。":::

1. 现在需要为这个流程设置时间表。 从 2021 年上半年开始，电子表格在每周一都有一个新的待命任务。 把流设置为星期一早上的首个运行的项。 使用以下选项将流配置为每周星期一运行。

    - **流名称**：通知待命人
    - **开始时间**：1/4/21 凌晨 1:00
    - **重复间隔**：1 周
    - **这些日期**：星期一

    :::image type="content" source="../images/power-automate-return-tutorial-3.png" alt-text="Power Automate &quot;生成计划云流&quot;对话框显示选项。选项包括流名称、开始时间、重复时间以及一周中的哪一天运行流。":::

1. 选择“**创建**”。

1. 选择“**新建步骤**”。

1. 选择 "**标准**" 选项卡，然后选择 "**Excel Online （企业）**"。

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text=" Power Automate 中的 Excel Online (商业版)选项。":::

1. 在 **操作** 下，选择 **运行脚本**。

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text=" Power Automate 中的运行脚本操作选项。":::

1. 接下来，选择要在流步骤中使用的工作簿和脚本。 使用 **on-call-rotation.xlsx** 在 OneDrive 中创建的工作簿。 为 **运行脚本** 连接器指定以下设置：

    - **位置**：OneDrive for Business
    - **文档库**：OneDrive
    - **文件**: on-call-rotation.xlsx *（通过文件浏览器选择）*
    - **脚本**：获取待命人员

    :::image type="content" source="../images/power-automate-return-tutorial-4.png" alt-text="用于运行脚本的 Power Automate 连接器设置。":::

1. 选择“**新建步骤**”。

1. 我们将通过发送提醒邮件来结束流。 使用连接器的搜索栏选择“**发送电子邮件 (V2)**”。 使用“**新增动态内容**”控件添加脚本返回的电子邮件地址。 这将被标记为 **结果**，旁边有 Excel 图标。 可以提供你想要的任何主题和正文。

    :::image type="content" source="../images/power-automate-return-tutorial-5.png" alt-text="用于发送电子邮件的 Power Automate Outlook 连接器设置。选项包括要发送的文件、电子邮件主题、电子邮件正文以及高级选项。":::

    > [!NOTE]
    > 此教程使用 Outlook。可以改用你喜欢的电子邮件服务，但某些选项可能不同。

1. 选择“**保存**”。

## <a name="test-the-script-in-power-automate"></a>在 Power Automate 功能中测试脚本

你的流将在每周一早上运行。 现在可以通过选择屏幕右上角的“**测试**”按钮来测试脚本。 选择“**手动**”，然后选择“**运行测试**”来立即运行流并测试行为。 可能需要向 Excel 和 Outlook 授予权限才能继续。

:::image type="content" source="../images/power-automate-return-tutorial-6.png" alt-text="&quot;Power Automate 测试&quot;按钮。":::

> [!TIP]
> 如果流无法发送电子邮件，请在电子表格中仔细检查是否在表格顶部列出了当前日期范围的有效电子邮件。

## <a name="next-steps"></a>后续步骤

访问[使用 Power Automate 运行 Office 脚本](../develop/power-automate-integration.md)，以了解有关将 Office Script 与 Power Automate 连接的更多信息。

你还可以查看[自动任务提醒示例场景](../resources/scenarios/task-reminders.md)，以了解如何将 Office 脚本和 Power Automate 与 Team Adaptive Cards 结合使用。
