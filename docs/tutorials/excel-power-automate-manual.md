---
title: 开始使用自动加电的脚本
description: 有关使用手动触发器将 Power 自动化与 Office 脚本集成的教程。
ms.date: 06/09/2020
localization_priority: Priority
ms.openlocfilehash: 37c2d9ae4c5456a1355362c70695fc61c236a725
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878740"
---
# <a name="start-using-scripts-with-power-automate-preview"></a>开始使用启用了 Power 自动化的脚本（预览）

本教程向您介绍如何通过[Power 自动化](https://flow.microsoft.com)在 web 上运行适用于 Excel 的 Office 脚本。

## <a name="prerequisites"></a>先决条件

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> 本教程假定您已在 web 教程中完成了[Excel 中的记录、编辑和创建 Office 脚本](excel-tutorial.md)。

## <a name="prepare-the-workbook"></a>准备工作簿

Power 自动执行无法使用相对引用 `Workbook.getActiveWorksheet` ，如访问工作簿组件。 因此，我们需要具有一致的名称的工作簿和工作表，电源自动化可以参考。

1. 创建一个名为**MyWorkbook**的新工作簿。

2. 在**MyWorkbook**工作簿中，创建一个名为 " **TutorialWorksheet**" 的工作表。

## <a name="create-an-office-script"></a>创建 Office 脚本

1. 转到 "**自动**" 选项卡，然后选择 "**代码编辑器**"。

2. 选择 "**新建脚本**"。

3. 将默认脚本替换为以下脚本。 此脚本将当前日期和时间添加到**TutorialWorksheet**工作表的前两个单元格中。

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

4. 重命名脚本以**设置日期和时间**。 若要更改此名称，请按脚本名称。

5. 通过按 "**保存脚本**" 保存该脚本。

## <a name="create-an-automated-workflow-with-power-automate"></a>使用 Power 自动化创建自动工作流

1. 登录到[Power 自动预览网站](https://flow.microsoft.com)。

2. 在屏幕左侧显示的菜单中，按 "**创建**"。 这将向你显示创建新工作流的方式列表。

    !["增强电源" 中的 "创建" 按钮。](../images/power-automate-tutorial-1.png)

3. 在 "**从空白开始**" 部分中，选择 "**即时流**"。 这将创建手动激活的工作流。

    ![用于创建新工作流的 "即时流" 选项。](../images/power-automate-tutorial-2.png)

4. 在显示的对话框窗口中，在 "**流名称**" 文本框中输入流的名称，从 "**选择如何触发流**" 下的选项列表中选择 "**手动触发流**"，然后按 "**创建**"。

    ![用于创建新的即时流的手动触发器选项。](../images/power-automate-tutorial-3.png)

5. 按 "**新建步骤**"。

6. 选择 "**标准**" 选项卡，然后选择 " **Excel Online （企业）**"。

    ![Excel Online （业务）的 "电源自动" 选项。](../images/power-automate-tutorial-4.png)

7. 在 "**操作**" 下，选择 "**运行脚本（预览）**"。

    !["运行脚本（预览）" 的 "电源自动操作" 选项。](../images/power-automate-tutorial-5.png)

8. 为 "**运行脚本**" 连接器指定以下设置：

    - **位置**： OneDrive for business
    - **文档库**： OneDrive
    - **文件**： MyWorkbook.xlsx
    - **脚本**：设置日期和时间

    ![用于在 Power 自动化中运行脚本的连接器设置。](../images/power-automate-tutorial-6.png)

9. 按 "**保存**"。

您的流程现已准备好通过电源自动运行。 您可以使用流编辑器中的 "**测试**" 按钮对其进行测试，也可以按照其余的教程步骤运行流集合中的流。

## <a name="run-the-script-through-power-automate"></a>通过 Power 自动运行脚本

1. 从 "主电自动" 页面中，选择 "**我的流**"。

    !["电源自动" 中的 "我的流" 按钮。](../images/power-automate-tutorial-7.png)

2. 从 "**我的流量**" 选项卡中显示的流列表中选择 **"我的教程流**"。这将显示之前创建的流的详细信息。

3. 按 "**运行**"。

    !["电源自动运行" 按钮。](../images/power-automate-tutorial-8.png)

4. 将显示用于运行流的任务窗格。 如果系统询问您是否**登录**到 Excel Online，请按 "**继续**"。

5. 按 "**运行流**"。 这将运行流，这将运行相关的 Office 脚本。

6. 按 "**完成**"。 您应该会看到 "**运行**" 部分进行了相应的更新。

7. 刷新页面以查看电源自动执行的结果。 如果成功，请转到工作簿以查看更新的单元格。 如果失败，请验证流设置并再次运行它。

    ![自动关闭显示成功流运行的输出。](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a>后续步骤

完成 "[使用 Power 自动运行脚本](excel-power-automate-trigger.md)" 教程。 它向您介绍如何将数据从工作流服务传递到您的 Office 脚本。
