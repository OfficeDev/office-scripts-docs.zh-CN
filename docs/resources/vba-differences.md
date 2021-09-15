---
title: 脚本Office VBA 宏之间的差异
description: 脚本和 VBA Office之间的行为和 API Excel差异。
ms.date: 05/21/2021
ms.localizationpriority: medium
ms.openlocfilehash: 5038f8c0195cb84a2b77065d6b4c6a53e813f6a4
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/15/2021
ms.locfileid: "59327880"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>脚本Office VBA 宏之间的差异

Office脚本和 VBA 宏有很多共同之处。 它们都允许用户通过易于使用的操作录制器自动处理解决方案，并允许编辑这些录制。 这两个框架旨在让可能不将自己认为是程序员的人在 Excel。
基本区别在于，VBA 宏是为桌面解决方案而开发的，Office脚本专为安全的基于云的解决方案设计。 目前，Office脚本仅在 Excel web 版 中受支持。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="四象限图，显示不同扩展性解决方案Office区域。Office脚本和 VBA 宏旨在帮助最终用户创建解决方案，但 Office 脚本是为 Web 和协作 (而 VBA 用于桌面) 。":::

本文介绍 VBA 宏与 (脚本和 VBA 之间的主要) Office。 由于Office脚本仅适用于Excel，因此这是此处讨论的唯一主机。

## <a name="platform-and-ecosystem"></a>平台和生态系统

VBA 专为桌面设计，Office脚本专为 Web 设计。 VBA 可以与用户的桌面进行交互，以使用类似技术（如 COM 和 OLE）进行连接。 但是，VBA 无法方便地调用 Internet。

Office脚本使用 JavaScript 的通用运行时。 这将提供一致的行为和辅助功能，而不考虑用于运行脚本的机器。 他们还可以调用其他 Web 服务。

## <a name="security"></a>安全性

VBA 宏与宏具有相同的安全Excel。 这样，他们可以访问你的桌面。 Office脚本只能访问工作簿，而无法访问托管工作簿的机器。 此外，无法与脚本共享 JavaScript 身份验证令牌。 这意味着脚本既不具有已登录用户的令牌，也没有用于登录到外部服务的任何 API 功能，因此它们无法使用现有令牌代表用户进行外部调用。

管理员有三个 VBA 宏选项：允许租户上的所有宏、不允许在租户上运行宏或只允许使用签名证书的宏。 这种缺少粒度使隔离单个错误参与者变得困难。 目前Office整个租户、整个租户或租户中的一组用户关闭脚本。 管理员还可以控制谁可以与其他人共享脚本，以及谁可以在 Power Automate。

## <a name="coverage"></a>覆盖范围

目前，VBA 更全面涵盖Excel功能，尤其是桌面客户端上提供的功能。 Office脚本几乎涵盖所有用于Excel web 版。 此外，随着新功能在 Web 上首次推出，Office脚本将同时支持操作录制器和 JavaScript API。

Office脚本不支持Excel级[事件](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)。 脚本仅在用户手动启动脚本或脚本流调用脚本Power Automate运行。

## <a name="power-automate"></a>Power Automate

Office脚本可以运行在Power Automate。 工作簿可以通过计划流或事件驱动的流进行更新，使工作流自动化，甚至无需打开Excel。 这意味着，只要工作簿存储在 OneDrive (中并且可供 Power Automate) 访问，无论您的组织是使用 Excel 桌面、Mac 还是 Web 客户端，流都可以运行脚本。

VBA 没有Power Automate连接器。 所有支持的 VBA 方案都涉及用户参与宏的执行。

尝试[从手动调用流教程Power Automate调用](../tutorials/excel-power-automate-manual.md)脚本，以开始了解Power Automate。 还可以查看自动任务[提醒示例，](scenarios/task-reminders.md)以查看Office方案中Teams Power Automate脚本。

## <a name="see-also"></a>另请参阅

- [Excel web 版中的 Office 脚本](../overview/excel.md)
- [使用 Office 运行脚本Power Automate](../develop/power-automate-integration.md)
- [Office 脚本与 Office 加载项之间的差异](add-ins-differences.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [Excel VBA 参考](/office/vba/api/overview/excel)
