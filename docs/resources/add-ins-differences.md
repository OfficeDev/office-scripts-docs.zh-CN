---
title: Office 脚本与 Office 加载项之间的差异
description: Office 脚本与 Office 外接程序之间的行为和 API 差异。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: ddac6cc68874da34ae76c66a5c5b84ffa7a60eec
ms.sourcegitcommit: ce72354381561dc167ea0092efd915642a9161b3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/30/2020
ms.locfileid: "48319649"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office 脚本与 Office 加载项之间的差异

Office 外接程序和 Office 脚本具有很多共同之处。 它们都提供对 Excel 工作簿的自动控制（JavaScript API）。 但是，Office 脚本 Api 是 Office JavaScript API 的专用的同步版本。

![显示不同 Office 扩展性解决方案的焦点区域的四象限图。 Office 脚本和 Office Web 外接程序都集中在 Web 和协作上，但 office 脚本适用于最终用户 (而 Office Web 外接程序目标专业开发人员) 。 ) ](../images/office-programmability-diagram.png)

Office 脚本运行到完成时需要按手动按钮按下或以 " [自动](https://flow.microsoft.com/)运行" 的步骤，而 office 加载项在其任务窗格处于打开状态时保持不变。 这意味着外接程序可以在会话期间维护状态，而 Office 脚本不会在两个运行之间保持内部状态。 如果发现您的 Excel 扩展需要超过脚本平台的功能，请访问 [Office 外接程序文档](/office/dev/add-ins) 以了解有关 Office 外接程序的详细信息。

本文的其余部分将介绍 Office 外接程序和 Office 脚本之间的主要差异。

## <a name="platform-support"></a>平台支持

Office 外接程序是跨平台的。 它们在 Windows 桌面、Mac、iOS 和 web 平台上工作，并在每个平台上提供相同的体验。 每个 API 的文档中注明了此错误的任何例外。

Office 脚本目前仅对 web 上的 Excel 受支持。 所有录制、编辑和运行都是在 web 平台上完成的。

## <a name="apis"></a>API

Office 外接程序没有 Office JavaScript Api 的同步版本。标准 Office 脚本 api 对平台是唯一的，并进行了大量优化和变更，以避免使用 `load` / `sync` 范例。

某些 [Excel JavaScript api](/javascript/api/excel?view=excel-js-preview&preserve-view=true) 与 [Office 脚本异步 api](../develop/excel-async-model.md)兼容。 某些示例和外接代码块可以 `Excel.run` 通过最少的转换移植到块。 虽然这两个平台共享功能，但有一些缺口。 Office 外接程序设置了两个主要 API，但 Office 脚本不是事件和常见 Api。

### <a name="events"></a>活动

Office 脚本不支持 [事件](/office/dev/add-ins/excel/excel-add-ins-events)。 每个脚本在一个方法中运行代码 `main` ，然后结束。 触发事件时不会重新激活，因此无法注册事件。

### <a name="common-apis"></a>通用 API

Office 脚本无法使用 [通用 api](/javascript/api/office)。 如果需要身份验证、对话窗口或其他仅受常见 Api 支持的功能，则您可能需要创建 Office 加载项，而不是 Office 脚本。

## <a name="see-also"></a>另请参阅

- [Excel 网页版中的 Office 脚本](../overview/excel.md)
- [Office 脚本和 VBA 宏之间的区别](vba-differences.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [生成 Excel 任务窗格加载项](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
