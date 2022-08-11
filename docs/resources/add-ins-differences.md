---
title: Office 脚本与 Office 加载项之间的差异
description: Office 脚本和 Office 加载项之间的行为和 API 差异。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: a3df4daf04f963598d2cb31f82dd2c1c9923fdc8
ms.sourcegitcommit: 33fe0f6807daefb16b148fd73c863de101f47cea
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/08/2022
ms.locfileid: "67281908"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office 脚本与 Office 加载项之间的差异

了解 Office 脚本和 Office 加载项之间的差异，以了解何时使用每个脚本。 Office 脚本旨在由任何希望改进其工作流的人员快速制作。 Office 外接程序与 Office UI 集成，通过功能区按钮和任务窗格获得更交互式体验。 Office 加载项还可以通过提供自定义函数来扩展内置 Excel 函数。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="显示不同 Office 扩展性解决方案的焦点区域的四象限关系图。Office 脚本和 Office Web 加载项都侧重于 Web 和协作，但 Office 脚本迎合最终用户 (而 Office Web 加载项则面向专业开发人员) 。":::

Office 脚本以手动按钮或 [Power Automate](https://flow.microsoft.com/) 中的步骤运行到完成，而 Office 加载项则根据配置方式继续运行。 例如，可以将 Office 加载项配置为在任务窗格关闭时继续运行。 这意味着 Office 加载项在会话期间保持状态，而 Office 脚本不会在运行之间保持内部状态。 如果要构建的解决方案需要维护状态，则应访问 [Office 加载项文档](/office/dev/add-ins) ，了解有关 Office 外接程序的详细信息。

本文的其余部分介绍 Office 加载项和 Office 脚本之间的主要区别。

## <a name="platform-support"></a>平台支持

Office 加载项是跨平台的。 它们跨 Windows 桌面、Mac、iOS 和 Web 平台工作，并在每个平台上提供相同的体验。 单个 API 的文档中会记下此项的任何异常。

Office 脚本目前仅受Excel web 版支持。 所有录制、编辑和脚本管理都是在 Web 平台上完成的。

### <a name="script-support-for-excel-on-windows"></a>Windows 上 Excel 的脚本支持

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>API

虽然 Office 外接程序的 Office JavaScript API 和 Office 脚本 API 共享一些功能，但它们是不同的平台。 Office 脚本 API 是 Excel JavaScript API 模型的优化同步子集。 主要区别是 `load`/`sync` 使用外接程序的范例。此外，外接程序还为事件提供 API，在 Excel 之外提供一组更广泛的功能（称为通用 API）。

### <a name="events"></a>活动

Office 脚本不支持工作簿级别 [的事件](/office/dev/add-ins/excel/excel-add-ins-events)。 脚本由用户选择脚本的 **“运行** ”按钮或通过 Power Automate 触发。 每个脚本在单 `main` 个函数中运行代码，然后结束。

### <a name="common-apis"></a>通用 API

Office 脚本不能使用 [通用 API](/javascript/api/office)。 如果需要身份验证、对话窗口或仅受公用 API 支持的其他功能，则可能需要创建 Office 加载项而不是 Office 脚本。

## <a name="see-also"></a>另请参阅

- [Excel 中的 Office 脚本](../overview/excel.md)
- [Office 脚本和 VBA 宏之间的差异](vba-differences.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [生成 Excel 任务窗格加载项](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
