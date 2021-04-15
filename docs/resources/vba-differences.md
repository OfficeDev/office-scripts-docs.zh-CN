---
title: Office 脚本和 VBA 宏之间的差异
description: Office 脚本和 Excel VBA 宏之间的行为和 API 差异。
ms.date: 12/14/2020
localization_priority: Normal
ms.openlocfilehash: a56409a5de3eb07876faa88bfbfe78eeca59f70f
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755019"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Office 脚本和 VBA 宏之间的差异

Office 脚本和 VBA 宏有很多共同之处。 它们都允许用户通过易于使用的操作录制器自动处理解决方案，并允许编辑这些录制。 这两个框架旨在让不将自己认为是程序员的人在 Excel 中创建小型程序。
基本区别在于，VBA 宏是为桌面解决方案开发的，而 Office 脚本的设计以跨平台支持和安全性作为指导原则。 目前，Office 脚本仅在 Excel 网页中受支持。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="一个四象限图表，显示不同 Office 扩展性解决方案的重点区域。Office 脚本和 VBA 宏旨在帮助最终用户创建解决方案，但 Office 脚本是为 Web 和协作 (而 VBA 用于桌面) 。":::

本文介绍 VBA 宏与常规 (和 Office 脚本) VBA 之间的主要区别。 由于 Office 脚本仅适用于 Excel，因此这是此处讨论的唯一主机。

## <a name="platform-and-ecosystem"></a>平台和生态系统

VBA 专为桌面设计，Office 脚本专为 Web 设计。 VBA 可以与用户的桌面进行交互，以使用类似技术（如 COM 和 OLE）进行连接。 但是，VBA 无法方便地调用 Internet。

Office 脚本使用 JavaScript 的通用运行时。 这将提供一致的行为和辅助功能，而不考虑用于运行脚本的机器。 他们还可以调用其他 Web 服务。

## <a name="security"></a>安全性

VBA 宏具有与 Excel 相同的安全许可。 这样，他们可以访问你的桌面。 Office 脚本只能访问工作簿，而无法访问托管工作簿的机器。 此外，无法与脚本共享 JavaScript 身份验证令牌。 这意味着脚本既不具有已登录用户的令牌，也没有用于登录到外部服务的任何 API 功能，因此它们无法使用现有令牌代表用户进行外部调用。

管理员有三个 VBA 宏选项：允许租户上的所有宏、不允许在租户上运行宏或只允许使用签名证书的宏。 这种缺少粒度会使隔离单个错误参与者变得困难。 目前，租户的 Office 脚本处于打开或关闭状态。 但是，我们正在努力使管理员能够更加控制单个脚本和脚本创建者。

## <a name="coverage"></a>覆盖范围

目前，VBA 更全面涵盖 Excel 功能，尤其是桌面客户端上提供的功能。 Office 脚本几乎涵盖 Excel 网页应用的所有方案。 此外，随着新功能在 Web 上首次推出，Office 脚本将同时支持操作录制器和 JavaScript API。

Office 脚本不支持 Excel 级 [事件](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)。 脚本仅在用户手动启动或 Power Automate 流调用脚本时运行。

## <a name="power-automate"></a>Power Automate

Office 脚本可以通过 Power Automate 运行。 工作簿可以通过计划流或事件驱动的流进行更新，让你无需打开 Excel 即可自动执行工作流。 这意味着，只要工作簿存储在 OneDrive (且可供 Power Automate) 访问，无论您和组织是使用 Excel 桌面版、Mac 还是 Web 客户端，流都可以运行脚本。

VBA 没有 Power Automate 连接器。 所有支持的 VBA 方案都涉及用户参与宏的执行。

## <a name="see-also"></a>另请参阅

- [Excel web 版中的 Office 脚本](../overview/excel.md)
- [Office 脚本与 Office 加载项之间的差异](add-ins-differences.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [Excel VBA 参考](/office/vba/api/overview/excel)
