---
title: 在流中使用启用Power Automate文件
description: 了解如何在流中使用启用宏的文件或 .xlsm Power Automate文件。
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f2ecefe9fb97d1c5514ddb52c3cbcd0596df426
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585742"
---
# <a name="how-to-use-macro-enabled-files-in-power-automate-flows"></a>如何在宏流中Power Automate文件

可以将 .xlsm 文件与Power Automate集成。 这使你可以开始将现有自动化解决方案转换为基于 Web 的格式。 请注意，.xslm 文件中包含的宏无法通过 Power Automate。 仅Office脚本。

[Power Automate](https://flow.microsoft.com/) [中的 Excel Online (Business) ](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) 连接器通常仅限于 Microsoft Excel Open XML 电子表格 (.xlsx) 格式的文件。 它的文件浏览器仅允许你选择.xlsx文件。 但是，如果使用文件元数据，则启用宏的文件与连接器的 **Run** 脚本操作兼容。

在流中，**从连接器或** 连接器OneDrive for Business"[](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/)[SharePoint"操作](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/)。 Run **脚本** 操作接受此元数据作为有效文件。 运行脚本时，使用从获取文件 **元数据** 操作返回的 *ID* 动态内容作为"文件"参数。 以下屏幕截图显示了一个流，该流向 Run 脚本操作提供名为"Test Macro File.xlsm" **的文件** 的元数据。

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="包含获取文件元数据操作（将宏文件的元数据传递到 Run 脚本操作）的流。":::

> [!WARNING]
> 某些 .xlsm 文件（尤其是具有 ActiveX 或 Form 控件的文件）可能无法在 Excel 连接器中工作。 请确保在部署解决方案之前进行测试。

## <a name="other-resources"></a>其他资源

[观看 Sudhi Ramamurthy 的 YouTube](https://youtu.be/o-H9BbywJQQ) 视频，了解如何在 Run Script 操作中使用 .xlsm 文件。
