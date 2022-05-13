---
title: Office脚本和 VBA 宏之间的差异
description: Office脚本和Excel VBA 宏之间的行为和 API 差异。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60e4fba6e63967302066f544b76fb20a8c8630a6
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393612"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Office脚本和 VBA 宏之间的差异

Office脚本和 VBA 宏有很多共同点。 它们都允许用户通过易于使用的操作录制器自动执行解决方案，并允许编辑这些录制。 这两个框架都旨在使那些可能不认为自己是程序员的人能够在Excel中创建小型项目。

根本区别在于，VBA 宏是为桌面解决方案开发的，Office脚本专为基于云的安全解决方案而设计。 目前，仅Excel web 版支持Office脚本。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="一个四象限关系图，显示不同Office扩展性解决方案的重点区域。Office脚本和 VBA 宏都旨在帮助最终用户创建解决方案，但Office脚本是为 Web 和协作 (构建的，而 VBA 则用于桌面) 。":::

本文介绍 VBA 宏 (与 VBA 在常规) 和Office脚本中的主要区别。 由于Office脚本仅适用于Excel，因此这是此处讨论的唯一主机。

## <a name="platform-and-ecosystem"></a>平台和生态系统

Windows和 Mac 上的Excel支持 VBA。 Excel web 版支持Office脚本。

这两个解决方案是为各自的平台设计的。 VBA 可以与用户的桌面交互，以连接类似的技术，如 COM 和 OLE。 但是，VBA 无法通过方便的方式调用 Internet。 Office脚本使用适用于 JavaScript 的通用运行时。 这提供一致的行为和辅助功能，而不考虑用于运行脚本的计算机。 他们还可以调用其他 Web 服务。

### <a name="script-support-for-excel-on-windows"></a>Windows上Excel的脚本支持

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="security"></a>安全性

VBA 宏的安全间隙与Excel相同。 这使他们能够完全访问桌面。 Office脚本只能访问工作簿，而不能访问托管工作簿的计算机。 此外，无法与脚本共享 JavaScript 身份验证令牌。 这意味着脚本既没有已登录用户的令牌，也没有用于登录到外部服务的任何 API 功能，因此无法使用现有令牌代表用户进行外部调用。

管理员有三个 VBA 宏选项：允许租户上的所有宏、不允许租户上任何宏，或者仅允许具有已签名证书的宏。 这种缺乏粒度使得很难孤立一个坏演员。 目前，对于整个租户、整个租户或租户中的一组用户，Office脚本可以关闭。 管理员还可以控制谁可以与他人共享脚本，以及谁可以在Power Automate中使用脚本。

## <a name="coverage"></a>覆盖

目前，VBA 提供更全面的Excel功能，尤其是桌面客户端上提供的功能。 Office脚本几乎涵盖了Excel web 版的所有方案。 此外，随着新功能在 Web 上首次推出，Office脚本将支持操作记录器和 JavaScript API。

Office脚本不支持Excel级[事件](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)。 只有当用户手动启动脚本或Power Automate流调用脚本时，脚本才会运行。

## <a name="power-automate"></a>Power Automate

Office脚本可以通过Power Automate运行。 工作簿可以通过计划流或事件驱动流进行更新，使你无需打开Excel即可自动执行工作流。 这意味着，只要工作簿存储在OneDrive (中且可供Power Automate) 访问，流就可以运行脚本，而不管你和你的组织是使用Excel桌面、Mac 还是 Web 客户端。

VBA 没有Power Automate连接器。 所有支持的 VBA 方案都涉及到用户参与宏的执行。

尝试[手动Power Automate流教程中的呼叫脚本](../tutorials/excel-power-automate-manual.md)，开始了解Power Automate。 还可以查看[自动任务提醒](scenarios/task-reminders.md)示例，查看Office脚本通过实际方案中的Power Automate连接到Teams。

## <a name="see-also"></a>另请参阅

- [Office Excel中的脚本](../overview/excel.md)
- [使用Power Automate运行Office脚本](../develop/power-automate-integration.md)
- [Office 脚本与 Office 加载项之间的差异](add-ins-differences.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [Excel VBA 参考](/office/vba/api/overview/excel)
