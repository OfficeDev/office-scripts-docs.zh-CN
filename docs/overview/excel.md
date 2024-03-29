---
title: Excel 中的 Office 脚本
description: Office 脚本中的操作录制器和代码编辑器简介。
ms.topic: overview
ms.date: 10/05/2022
ms.localizationpriority: high
ms.openlocfilehash: 02a45e5aae468cff2c61e18b35c54ba656d0484b
ms.sourcegitcommit: 64d506257bee282fb01aedbf4d090781b06e4900
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/07/2022
ms.locfileid: "68495472"
---
# <a name="office-scripts-in-excel"></a>Excel 中的 Office 脚本

Excel 中的 Office 脚本使你能够自动执行日常任务。 在 Excel 网页版中，可以使用操作记录器记录操作。 这会创建一个可随时再次运行的 TypeScript 语言脚本。 此外，你还可以使用代码编辑器创建和编辑脚本。 然后，可在组织中共享你的脚本，以便同事也可实现其工作流的自动化。

本文档系列将指导你如何使用这些工具。 我们将向你介绍操作录制器，让你了解如何录制频繁的 Excel 操作。 你还将学习如何使用代码编辑器创建或更新自己的脚本。

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a>Requirements

若要使用 Office 脚本，需要以下内容。

1. [Excel web 版](https://www.office.com/launch/excel) (Windows 上的 Excel 只能将 Office 脚本与[脚本按钮](../develop/script-buttons.md)配合使用) 。

    > [!TIP]
    > Office 脚本现可在 Windows 版 Office 版和 Mac 版 Office [预览体验成员](https://insider.office.com/)版中使用。

1. OneDrive for Business。
1. 可访问 Microsoft 365 Office 桌面应用的任何商业版或教育版 Microsoft 365 许可证，例如：

    - Office 365 商业版
    - Office 365 商业高级版
    - Office 365 专业增强版
    - Office 365 专业增强版（设备）
    - Office 365 企业版 E3
    - Office 365 企业版 E5
    - Office 365 A3
    - Office 365 A5

> [!NOTE]
> 如果符合这些要求，但仍不能看到 **Automate** 选项卡，你的管理员可能已禁用此功能，或者环境存在其他问题。 请按照 [Automate 选项卡未出现或 Office 脚本不可用 ](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable) 下的步骤开始使用 Office 脚本。

## <a name="when-to-use-office-scripts"></a>何时使用 Office 脚本

你可以使用脚本录制和重播不同工作簿和工作表上的 Excel 操作。 如果你发现自己正在重复执行相同的操作，则可以将所有工作转变为易于运行的 Office 脚本。 通过 Excel 中的一个按钮运行脚本，或将其与 Power Automate 结合使用，简化整个工作流程。

例如，假如你在 Excel 中打开一个会计网站的 .csv 文件，以此开始一天的工作。 你需要花几分钟删除不必要的列，设置表格格式，添加公式和在新工作表中创建一个数据透视表。 你可以使用操作录制器录制这些每天重复的操作。 录制之后，运行脚本即可处理整个 .csv 转换。 这样不仅可以消除忘记步骤的风险，而且还能够与他们共享流程，无需为他们提供任何指导。 通过 Office 脚本，可以自动化执行常见任务，提高你和工作区的效率和生产力。

## <a name="action-recorder"></a>操作录制器

:::image type="content" source="../images/action-recorder-intro.png" alt-text="操作记录器记录的操作列表。":::

操作录制器可以录制你在 Excel 中进行的操作，并将它们转换为脚本。 运行操作录制器之后，你可以在编辑单元格、更改格式和创建表格时捕获 Excel 操作。 可以在其他工作表和工作簿上运行生成的脚本，以重复创建原始操作。

## <a name="code-editor"></a>代码编辑器

:::image type="content" source="../images/code-editor-intro.png" alt-text="代码编辑器显示本教程中使用的脚本代码。":::

使用操作录制器录制的所有脚本均可通过代码编辑器编辑。 这使你能够调整和自定义脚本，以更好地满足你的准确需求。 此外，你还可以添加不能直接通过 Excel UI 访问的逻辑和功能，例如条件语句 (if/else) 和循环。

> [!TIP]
> 操作记录器有一个 **复制为代码** 按钮，用于将操作记录到脚本代码中，而无需保存整个脚本。
>
> :::image type="content" source="../images/action-recorder-copy-code.png" alt-text="操作记录器任务窗格，其中突出显示了“复制为代码”按钮。":::

我们的[教程](../tutorials/excel-tutorial.md)提供了一种引导式和结构化的方式来了解 Office 脚本的功能。 完成本教程之后，请阅读 [Excel 网页版中 Office 脚本的编写脚本基础](../develop/scripting-fundamentals.md)，以了解有关代码编辑器以及如何编写和编辑你自己的脚本的详细信息。 有关代码编辑器以及如何解读脚本代码的其他信息，请阅读 [Office 脚本代码编辑器环境](code-editor-environment.md)。

## <a name="share-office-scripts"></a>共享Office脚本

Office 脚本可与 Excel 工作簿的其他用户共享。 当共享了共享工作簿中的脚本时，有权访问该工作簿的每个人都可以查看和运行该脚本。 有关共享和取消共享脚本的详细信息，请参阅[在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)。

:::image type="content" source="../images/script-sharing.png" alt-text="显示“在此工作簿中与其他人共享”选项的脚本详细信息页面。":::

添加用于运行脚本的按钮，以帮助同事发现有价值的解决方案，并让他们在 Excel 桌面版中运行脚本。 在 Run [Office Scripts with buttons 中了解有关脚本按钮的更多信息](../develop/script-buttons.md)。

:::image type="content" source="../images/add-button.png" alt-text="单击时运行脚本的工作表中的一个按钮。":::

> [!NOTE]
> 请参阅 [ Office 脚本存储和所有权 ](script-storage.md) 了解关于如何在 OneDrive 中存储脚本的详细信息。

## <a name="connect-office-scripts-to-power-automate"></a>将 Office 脚本连接到 Power Automate

[Power Automate](https://flow.microsoft.com/) 是一种可帮助你在多个应用和服务之间创建自动化工作流的服务。 Office 脚本可以在这些工作流中使用，以便你在工作簿之外控制脚本。 你可以按计划运行脚本，在回复电子邮件时触发它们，等等。 若要了解有关连接这些自动化服务的基础知识，请访问[使用 Power Automate 运行 Office 脚本](../tutorials/excel-power-automate-manual.md)教程。

## <a name="next-steps"></a>后续步骤

完成 [Excel 网页版上的 Office 脚本教程](../tutorials/excel-tutorial.md)，以了解如何创建你的第一个脚本。

## <a name="see-also"></a>另请参阅

- [Excel 网页版中 Office 脚本的脚本基础知识](../develop/scripting-fundamentals.md)
- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [M365 中的 Office 脚本设置](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Excel 中的 Office 脚本简介](https://support.microsoft.com/office/9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office 脚本开发中心](https://developer.microsoft.com/office-scripts)
