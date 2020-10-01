---
title: Excel 网页版中的 Office 脚本
description: Office 脚本中的操作录制器和代码编辑器简介。
ms.date: 09/29/2020
localization_priority: Priority
ms.openlocfilehash: 965e28be285d59d79d46fe005ab16f29b271041f
ms.sourcegitcommit: ce72354381561dc167ea0092efd915642a9161b3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/30/2020
ms.locfileid: "48319670"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a>Excel 网页版中的 Office 脚本（预览版）

Excel 网页版中的 Office 脚本可以让您可以自动化日常任务。 你可以使用操作录制器录制 Excel 操作，这会创建一个脚本。 此外，你还可以使用代码编辑器创建和编辑脚本。 然后，可在组织中共享你的脚本，以便同事也可实现其工作流的自动化。

本文档系列将指导你如何使用这些工具。 我们将向你介绍操作录制器，让你了解如何录制频繁的 Excel 操作。 你还将学习如何使用代码编辑器创建或更新自己的脚本。

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a>Requirements

[!INCLUDE [Preview note](../includes/preview-note.md)]

若要使用 Office 脚本，需要以下内容。

1. [Excel 网页版](https://www.office.com/launch/excel)（不支持桌面等其他平台）。
1. OneDrive for Business。
1. [管理员已启用](/microsoft-365/admin/manage/manage-office-scripts-settings) Office 脚本。
1. 可访问 Microsoft 365 Office 桌面应用的任何商业版或教育版 Microsoft 365 许可证，例如：

    - Office 365 商业版
    - Office 365 商业高级版
    - Office 365 专业增强版
    - Office 365 专业增强版（设备）
    - Office 365 企业版 E3
    - Office 365 企业版 E5
    - Office 365 A3
    - Office 365 A5

## <a name="when-to-use-office-scripts"></a>何时使用 Office 脚本

你可以使用脚本录制和重播不同工作簿和工作表上的 Excel 操作。 如果你发现自己正在重复执行相同的操作，则可以将所有工作转变为易于运行的 Office 脚本。 通过 Excel 中的一个按钮运行脚本，或将其与 Power Automate 结合使用，简化整个工作流程。

例如，假如你在 Excel 中打开一个会计网站的 .csv 文件，以此开始一天的工作。 你需要花几分钟删除不必要的列，设置表格格式，添加公式和在新工作表中创建一个数据透视表。 你可以使用操作录制器录制这些每天重复的操作。 录制之后，运行脚本即可处理整个 .csv 转换。 这样不仅可以消除忘记步骤的风险，而且还能够与他们共享流程，无需为他们提供任何指导。 Office 脚本可以自动化常见任务，使你和你的工作空间可以更有效率、更加高效。

## <a name="action-recorder"></a>操作录制器

![录制若干操作之后的操作录制器。](../images/action-recorder-intro.png)

操作录制器可以录制你在 Excel 中进行的操作，并将它们转换为脚本。 运行操作录制器之后，你可以在编辑单元格、更改格式和创建表格时捕获 Excel 操作。 可以在其他工作表和工作簿上运行生成的脚本，以重复创建原始操作。

## <a name="code-editor"></a>代码编辑器

![显示以上脚本的脚本代码的代码编辑器。](../images/code-editor-intro.png)

使用操作录制器录制的所有脚本均可通过代码编辑器编辑。 这使你能够调整和自定义脚本，以更好地满足你的准确需求。 此外，你还可以添加不能直接通过 Excel UI 访问的逻辑和功能，例如条件语句 (if/else) 和循环。

一种简单的开始学习 Office 脚本方式就是在 Excel 网页版上录制脚本，然后查看生成的代码。 另一种选择是按照我们的[教程](../tutorials/excel-tutorial.md)进行，以更具指导性的结构化方式进行学习。

## <a name="sharing-scripts"></a>共享脚本

![显示“在此工作簿中与其他人共享”选项的脚本“详细信息”页面。](../images/script-sharing.png)

Office 脚本可与 Excel 工作簿的其他用户共享。 在工作簿中与其他人共享脚本时，该脚本将附加到工作簿中。 你的脚本存储在你的 OneDrive 中，当你共享一个脚本时，你将在打开的工作簿中创建指向该脚本的链接。

有关共享和取消共享脚本的详细信息，请参阅[在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)一文。

> [!NOTE]
> 由于 Office 脚本存储在用户的 OneDrive 上，因此它们遵循相同的保留和删除策略。 若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。

## <a name="connecting-office-scripts-to-power-automate"></a>将 Office 脚本连接到 Power Automate

[Power Automate](https://flow.microsoft.com/) 是一种可帮助你在多个应用和服务之间创建自动化工作流的服务。 Office 脚本可以在这些工作流中使用，以便你在工作簿之外控制脚本。 你可以按计划运行脚本，在回复电子邮件时触发它们，等等。 若要了解有关连接这些自动化服务的基础知识，请访问[使用 Power Automate 在 Excel 网页版中运行 Office 脚本](../tutorials/excel-power-automate-manual.md)教程。

## <a name="next-steps"></a>后续步骤

完成 [Excel 网页版上的 Office 脚本教程](../tutorials/excel-tutorial.md)，以了解如何创建你的第一个 Office 脚本。

## <a name="see-also"></a>另请参阅

- [Excel 网页版中 Office 脚本的脚本基础知识](../develop/scripting-fundamentals.md)
- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [M365 中的 Office 脚本设置](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Excel 中的 Office 脚本简介 (support.office.com)](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US)
