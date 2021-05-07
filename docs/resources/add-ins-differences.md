---
title: Office 脚本与 Office 加载项之间的差异
description: 脚本和加载项Office API 的行为Office API 差异。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 45993d08d85cfceb299216dddbe2e7da9fd2e404
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232632"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office 脚本与 Office 加载项之间的差异

Office加载项和Office脚本有很多共同之处。 它们都提供对 JavaScript API Excel工作簿的自动化控制。 但是，Office脚本 API 是 JavaScript API 的专用Office同步版本。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="四象限图，显示不同扩展性解决方案Office区域。Office 脚本和 Office Web 外接程序均侧重于 Web 和协作，但 Office 脚本适合最终用户 (而 Office Web 外接程序面向专业开发人员) ":::

Office脚本通过手动按下按钮或作为 Power Automate 中的步骤运行以[](https://flow.microsoft.com/)完成，Office任务窗格打开时，外接程序将保持运行状态。 这意味着加载项可以在会话期间保持状态，而Office脚本不会在两次运行之间保持内部状态。 如果您发现您的 Excel 扩展需要超过脚本平台的功能，请访问[Office 外接程序](/office/dev/add-ins)文档以了解有关 Office 外接程序的信息。

本文的其余部分介绍加载项和脚本Office之间的主要Office区别。

## <a name="platform-support"></a>平台支持

Office外接程序是跨平台的。 它们跨桌面Windows、Mac、iOS 和 Web 平台运行，并在每个平台上提供相同的体验。 有关此情况的任何例外情况都记录在单个 API 的文档中。

Office脚本当前仅受 Excel web 版。 所有录制、编辑和运行均在 Web 平台上完成。

## <a name="apis"></a>API

没有适用于外接程序的 Office JavaScript API Office版本。标准Office脚本 API 对于平台是唯一的，并且具有许多优化和更改以避免使用 `load` / `sync` 范例。

一些[Excel JavaScript API](/javascript/api/excel?view=excel-js-preview&preserve-view=true)与 Office[脚本异步 API 兼容](../develop/excel-async-model.md)。 一些示例和外接程序代码块可以移植到 `Excel.run` 转换最少的块。 虽然这两个平台共享功能，但存在一些差异。 加载项具有但Office脚本的两个主要 API 集Office事件和通用 API。

### <a name="events"></a>事件

Office脚本不支持[事件](/office/dev/add-ins/excel/excel-add-ins-events)。 每个脚本在一个方法中运行 `main` 代码，然后结束。 它不会在触发事件时重新激活，因此无法注册事件。

### <a name="common-apis"></a>通用 API

Office脚本不能使用[通用 API。](/javascript/api/office) 如果你需要身份验证、对话框窗口或其他仅受通用 API 支持的功能，你可能需要创建一个 Office 外接程序，而不是一个 Office 脚本。

## <a name="see-also"></a>另请参阅

- [Excel web 版中的 Office 脚本](../overview/excel.md)
- [脚本Office VBA 宏之间的差异](vba-differences.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [生成 Excel 任务窗格加载项](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
