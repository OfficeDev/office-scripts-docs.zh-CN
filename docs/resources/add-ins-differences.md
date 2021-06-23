---
title: Office 脚本与 Office 加载项之间的差异
description: 脚本和加载项Office API 的行为Office API 差异。
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: 17d66e37c7bf2b1263c0232bb0afb3ee4d29aa36
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074562"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office 脚本与 Office 加载项之间的差异

了解Office脚本Office外接程序之间的差异，以了解何时使用每个脚本和外接程序。 Office脚本旨在让任何希望改进其工作流的人快速创建脚本。 Office外接程序与 Office UI 集成，通过功能区按钮和任务窗格实现更具交互性的体验。 Office加载项还可以通过提供自定义函数来扩展内置Excel函数。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="四象限图，显示不同扩展性解决方案Office区域。Office 脚本和 Office Web 外接程序都侧重于 Web 和协作，但 Office 脚本适合最终用户 (而 Office Web 外接程序面向专业开发人员) 。":::

Office脚本通过手动按下按钮或作为 Power Automate 中的步骤运行以[](https://flow.microsoft.com/)完成，Office外接程序将继续运行，具体取决于其配置方式。 例如，可以将加载项Office，即使任务窗格关闭，也可以继续运行。 这意味着Office加载项在会话期间保持状态，而Office脚本不会在运行之间保持内部状态。 如果您要构建的解决方案需要保持状态，则应该访问 Office[外接程序](/office/dev/add-ins)文档，以了解有关Office外接程序的信息。

本文的其余部分介绍加载项和脚本Office之间的主要Office区别。

## <a name="platform-support"></a>平台支持

Office外接程序是跨平台的。 它们跨桌面Windows、Mac、iOS 和 Web 平台运行，并在每个平台上提供相同的体验。 有关此情况的任何例外情况都记录在单个 API 的文档中。

Office脚本当前仅受 Excel web 版。 所有录制、编辑和运行均在 Web 平台上完成。

## <a name="apis"></a>API

尽管Office加载项Office JavaScript API 和 Office 脚本 API 共享一些功能，但两者是不同的平台。 Office脚本 API 是 JavaScript API 模型的优化Excel子集。 主要区别是范例 `load` / `sync` 与加载项的用法。此外，加载项还提供事件 API 以及 Excel 之外的一组更广泛的功能，称为通用 API。

### <a name="events"></a>活动

Office脚本不支持工作簿级[事件](/office/dev/add-ins/excel/excel-add-ins-events)。 脚本由用户按脚本的 **"运行**"按钮触发，或Power Automate。 每个脚本在一个方法中运行 `main` 代码，然后结束。

### <a name="common-apis"></a>通用 API

Office脚本不能使用[通用 API。](/javascript/api/office) 如果你需要身份验证、对话框窗口或其他仅受通用 API 支持的功能，你可能需要创建一个 Office 外接程序，而不是一个 Office 脚本。

## <a name="see-also"></a>另请参阅

- [Excel web 版中的 Office 脚本](../overview/excel.md)
- [脚本Office VBA 宏之间的差异](vba-differences.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [生成 Excel 任务窗格加载项](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
