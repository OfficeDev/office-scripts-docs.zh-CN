---
title: 在流中Power Automate文件
description: 了解如何在流中使用宏文件或 xlsm Power Automate文件。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b232a1d31a7ff6e28016c5e28fd8a83c8d3f1859
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232653"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>如何在宏流中Power Automate文件

[Power Automate](https://flow.microsoft.com/)[流Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)连接器，以帮助将 Excel 文件与组织数据的其余部分以及应用程序（如 Teams、Outlook 和 SharePoint）连接。

但是，无法从"文件"下拉列表中选择宏 (请参阅以下屏幕截图中的示例) 。

:::image type="content" source="../images/no-xlsm.png" alt-text="The Power Automate Run script action showing no macro file selected.显示的错误为&quot;File&quot;是必需的":::

解决此问题的一个方法就是包括"获取文件元数据"操作 (OneDrive 或 SharePoint) 并使用"运行脚本"操作中的 ID 属性，如以下屏幕截图所示。

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="The Power Automate Run script action showing the macro file selected and no Run script error":::

> [!NOTE]
> 某些 XLSM (，尤其是具有 ActiveX/Form) 的 XLSM 在 Excel 连接器中可能不起作用。 请确保在部署解决方案之前进行测试。

## <a name="other-resources"></a>其他资源

[观看 Sudhi Ramamurthy 的 YouTube 视频](https://youtu.be/o-H9BbywJQQ)，了解如何在 Run Script 操作中使用 .xlsm 文件。
