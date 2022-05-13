---
title: Office 脚本与 Office 加载项之间的差异
description: Office脚本和Office加载项之间的行为和 API 差异。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: bd483f928e3e153b8a08537f6b333c3ea8d724dd
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393619"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Office 脚本与 Office 加载项之间的差异

了解Office脚本和Office加载项之间的差异，以了解何时使用每个脚本。 Office脚本设计为快速由任何希望改进其工作流的人制作。 Office外接程序与 Office UI 集成，通过功能区按钮和任务窗格获得更具交互性的体验。 Office外接程序还可以通过提供自定义函数来扩展内置Excel函数。

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="一个四象限关系图，显示不同Office扩展性解决方案的焦点区域。Office脚本和Office Web 加载项都侧重于 Web 和协作，但Office脚本迎合最终用户 (，而Office Web 加载项则面向专业开发人员) 。":::

Office脚本运行到完成时，按下手动按钮或作为[Power Automate](https://flow.microsoft.com/)中的一个步骤，而Office加载项则根据配置方式继续运行。 例如，可以将Office加载项配置为在任务窗格关闭时继续运行。 这意味着Office加载项在会话期间保持状态，而Office脚本不会在运行之间保持内部状态。 如果要构建的解决方案需要维护状态，则应访问[Office加载项文档](/office/dev/add-ins)，详细了解Office加载项。

本文的其余部分介绍Office加载项和Office脚本之间的主要区别。

## <a name="platform-support"></a>平台支持

Office加载项是跨平台的。 它们跨Windows桌面、Mac、iOS和 Web 平台工作，并在每个平台上提供相同的体验。 单个 API 的文档中会记下此项的任何异常。

Office脚本目前仅受Excel web 版支持。 所有录制、编辑和脚本管理都是在 Web 平台上完成的。

### <a name="script-support-for-excel-on-windows"></a>Windows上Excel的脚本支持

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>API

虽然Office加载项和Office脚本 API 的 Office JavaScript API 共享一些功能，但它们是不同的平台。 Office脚本 API 是 Excel JavaScript API 模型的优化同步子集。 主要区别是`load`/`sync`使用外接程序的范例。此外，外接程序还提供事件的 API 和Excel之外更广泛的功能集（称为通用 API）。

### <a name="events"></a>活动

Office脚本不支持工作簿级别[的事件](/office/dev/add-ins/excel/excel-add-ins-events)。 脚本由用户为脚本选择 **“运行**”按钮或通过Power Automate触发。 每个脚本在单 `main` 个方法中运行代码，然后结束。

### <a name="common-apis"></a>通用 API

Office脚本不能使用[通用 API](/javascript/api/office)。 如果需要身份验证、对话窗口或仅受通用 API 支持的其他功能，则可能需要创建Office加载项，而不是Office脚本。

## <a name="see-also"></a>另请参阅

- [Office Excel中的脚本](../overview/excel.md)
- [Office脚本和 VBA 宏之间的差异](vba-differences.md)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [生成 Excel 任务窗格加载项](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
