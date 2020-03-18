---
title: Office 脚本与 Office 外接程序之间的差异
description: Office 脚本与 Office 外接程序之间的行为和 API 差异。
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700127"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office 脚本与 Office 外接程序之间的差异

Office 外接程序和 Office 脚本具有很多共同之处。 它们都通过 Office JavaScript API 的`Excel`命名空间提供对 Excel 工作簿的自动控制。 但是，Office 脚本的作用范围更有限。

Office 脚本运行到完成时只需要按手动按钮，而 Office 外接程序依赖于用户交互并在工作簿使用时保持不变。 如果发现您的 Excel 扩展需要超过脚本平台的功能，请访问[Office 外接程序文档](/office/dev/add-ins)以了解有关 Office 外接程序的详细信息。

本文的其余部分将介绍 Office 外接程序和 Office 脚本之间的主要差异。

## <a name="platform-support"></a>平台支持

Office 外接程序是跨平台的。 它们在 Windows 桌面、Mac、iOS 和 web 平台上工作，并在每个平台上提供相同的体验。 每个 API 的文档中注明了此错误的任何例外。

Office 脚本目前仅对 web 上的 Excel 受支持。 所有录制、编辑和运行都是在 web 平台上完成的。

## <a name="apis"></a>API

Office 脚本支持大多数 Excel JavaScript Api，这意味着这两个平台之间存在许多功能重叠。 有两个例外：事件和常见 Api。

### <a name="events"></a>活动

Office 脚本不支持[事件](/office/dev/add-ins/excel/excel-add-ins-events)。 每个脚本在一个`main`方法中运行代码，然后结束。 触发事件时不会重新激活，因此无法注册事件。

### <a name="common-apis"></a>通用 API

Office 脚本无法使用[通用 api](/javascript/api/office)。 如果需要身份验证、对话窗口或其他仅受常见 Api 支持的功能，则您可能需要创建 Office 加载项，而不是 Office 脚本。

## <a name="see-also"></a>另请参阅

- [Web 上的 Excel 中的 Office 脚本](../overview/excel.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [生成 Excel 任务窗格加载项](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)