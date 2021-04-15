---
title: Office 脚本与 Office 加载项之间的差异
description: Office 脚本和 Office 外接程序之间的行为和 API 差异。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 96af98ca9f247406c5cc916f38892c318d33c560
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755096"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office 脚本与 Office 加载项之间的差异

Office 外接程序和 Office 脚本有很多共同之处。 它们均提供对 Excel 工作簿的 JavaScript API 的自动化控制。 但是，Office 脚本 API 是 Office JavaScript API 的专用同步版本。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="一个四象限图表，显示不同 Office 扩展性解决方案的重点区域。Office 脚本和 Office Web 外接程序均侧重于 Web 和协作，但 Office 脚本适合最终用户 (而 Office Web 外接程序面向专业开发人员) 。":::

Office 脚本通过手动按下按钮或作为 [Power Automate](https://flow.microsoft.com/)中的一个步骤运行以完成，而 Office 外接程序在任务窗格打开时仍然存在。 这意味着加载项可以在会话期间保持状态，而 Office 脚本不会在两次运行之间保持内部状态。 如果发现 Excel 扩展需要超过脚本平台的功能，请访问 [Office](/office/dev/add-ins) 加载项文档，详细了解 Office 加载项。

本文的其余部分将介绍 Office 外接程序和 Office 脚本之间的主要差异。

## <a name="platform-support"></a>平台支持

Office 外接程序是跨平台的。 它们跨 Windows 桌面、Mac、iOS 和 Web 平台运行，并且在每个平台上提供相同的体验。 有关此情况的任何例外情况都记录在单个 API 的文档中。

Office 脚本当前仅受 Excel 网页版本支持。 所有录制、编辑和运行均在 Web 平台上完成。

## <a name="apis"></a>API

没有适用于 Office 外接程序的 Office JavaScript API 的同步版本。标准 Office 脚本 API 对于平台是唯一的，并且具有大量优化和更改，以避免使用 `load` / `sync` 范例。

某些 [Excel JavaScript API](/javascript/api/excel?view=excel-js-preview&preserve-view=true) 与 Office [脚本异步 API 兼容](../develop/excel-async-model.md)。 一些示例和外接程序代码块可以移植到 `Excel.run` 转换最少的块。 虽然这两个平台共享功能，但存在一些差异。 Office 外接程序具有的两个主要 API 集，但 Office 脚本不是事件和通用 API。

### <a name="events"></a>活动

Office 脚本不支持 [事件](/office/dev/add-ins/excel/excel-add-ins-events)。 每个脚本在一个方法中运行 `main` 代码，然后结束。 它不会在触发事件时重新激活，因此无法注册事件。

### <a name="common-apis"></a>通用 API

Office 脚本不能使用[通用 API。](/javascript/api/office) 如果您需要身份验证、对话框窗口或其他仅受通用 API 支持的功能，您可能需要创建 Office 外接程序而不是 Office 脚本。

## <a name="see-also"></a>另请参阅

- [Excel web 版中的 Office 脚本](../overview/excel.md)
- [Office 脚本和 VBA 宏之间的差异](vba-differences.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [生成 Excel 任务窗格加载项](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
