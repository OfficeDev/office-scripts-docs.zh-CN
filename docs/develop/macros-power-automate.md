---
title: 在流中Power Automate文件
description: 了解如何在流中使用宏文件或 xlsm Power Automate文件。
ms.date: 09/01/2021
localization_priority: Normal
ms.openlocfilehash: 204eb8315f90c0ab0e20a34ec64517fbfbf304b1
ms.sourcegitcommit: 9472e78eb186ceffdaaceb2718d5074ce55a0e74
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2021
ms.locfileid: "58866534"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>如何在流中Power Automate文件

Excel[中的 Excel Online (Business) ](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)连接器Power Automate仅适用于[](https://flow.microsoft.com/)Microsoft Excel Open XML 电子表格 (.xlsx) 文件。 文件浏览器将你的选择范围.xlsx连接器内的文件。 但是，如果使用文件元数据，宏文件与连接器的 **Run** 脚本操作兼容。

在流中，从连接器或连接器[OneDrive for Business获取SharePoint](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/)[操作。](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) Run **脚本** 操作接受此元数据作为有效文件。 运行脚本时，使用从获取文件 **元数据** 操作返回的 *ID* 动态内容作为"文件"参数。 以下屏幕截图显示了一个流，该流向 **Run** 脚本操作提供名为"Test Macro File.xlsm"的文件的元数据。

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="包含获取文件元数据操作（将宏文件的元数据传递到 Run 脚本操作）的流。":::

> [!WARNING]
> 某些 .xlsm 文件（尤其是具有 ActiveX 或 Form 控件的文件）可能无法在 Excel 连接器中工作。 请确保在部署解决方案之前进行测试。

## <a name="other-resources"></a>其他资源

[观看 Sudhi Ramamurthy 的 YouTube 视频](https://youtu.be/o-H9BbywJQQ)，了解如何在 Run Script 操作中使用 .xlsm 文件。
