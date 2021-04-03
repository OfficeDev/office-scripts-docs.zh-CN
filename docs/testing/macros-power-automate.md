---
title: 在 Power Automate 流中使用宏文件
description: 了解如何在 Power Automate 流中使用宏文件或 xlsm 文件。
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: ec1fe00eb9ddc382ae4bc02187de7a36c97288b1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571247"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>如何在 Power Automate 流中使用宏文件

[Power Automate 流](https://flow.microsoft.com/) 提供了 [Excel 连接器](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) ，可帮助将 Excel 文件与其余组织数据和应用（如 Teams、Outlook 和 SharePoint）连接。

但是，无法从"文件"下拉列表中选择宏 (请参阅以下屏幕截图中的示例) 。

![运行脚本操作中无 xlsm](../images/no-xlsm.png)

解决此问题的一个方法就是将"获取文件元数据"操作 (OneDrive 或 SharePoint) ，并使用"运行脚本"操作中的 ID 属性，如以下屏幕截图所示。

![运行脚本操作中的 xlsm](../images/xlsm-in-pa.png)

> [!NOTE]
> 某些 XLSM (，尤其是具有 ActiveX/Form) 的 XLSM 在 Excel 联机连接器中可能不起作用。 请确保在部署解决方案之前进行测试。

[![观看有关在运行脚本操作中使用 XLSM 的视频](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "有关在运行脚本操作中使用 XLSM 的视频")
