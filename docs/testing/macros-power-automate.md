---
title: 在流中Power Automate文件
description: 了解如何在流中使用宏文件或 xlsm Power Automate文件。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 91e11424e4220a3e1f80cdd2711d05f219016147
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074639"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="054fd-103">如何在宏流中Power Automate文件</span><span class="sxs-lookup"><span data-stu-id="054fd-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="054fd-104">[Power Automate](https://flow.microsoft.com/)[流Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/)连接器，以帮助将 Excel 文件与组织数据的其余部分以及应用程序（如 Teams、Outlook 和 SharePoint）连接。</span><span class="sxs-lookup"><span data-stu-id="054fd-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="054fd-105">但是，无法从"文件"下拉列表中选择宏 (请参阅以下屏幕截图中的示例) 。</span><span class="sxs-lookup"><span data-stu-id="054fd-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="The Power Automate Run script action showing no macro file selected.显示的错误为&quot;File&quot;是必需的。":::

<span data-ttu-id="054fd-107">解决此问题的一个方法就是包括"获取文件元数据"操作 (OneDrive 或 SharePoint) 并使用"运行脚本"操作中的 ID 属性，如以下屏幕截图所示。</span><span class="sxs-lookup"><span data-stu-id="054fd-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="The Power Automate Run script action showing the macro file selected and no Run script error.":::

> [!NOTE]
> <span data-ttu-id="054fd-109">某些 XLSM (，尤其是具有 ActiveX/Form) 的 XLSM 在 Excel 连接器中可能不起作用。</span><span class="sxs-lookup"><span data-stu-id="054fd-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="054fd-110">请确保在部署解决方案之前进行测试。</span><span class="sxs-lookup"><span data-stu-id="054fd-110">Be sure to test before deploying your solution.</span></span>

## <a name="other-resources"></a><span data-ttu-id="054fd-111">其他资源</span><span class="sxs-lookup"><span data-stu-id="054fd-111">Other resources</span></span>

<span data-ttu-id="054fd-112">[观看 Sudhi Ramamurthy 的 YouTube 视频](https://youtu.be/o-H9BbywJQQ)，了解如何在 Run Script 操作中使用 .xlsm 文件。</span><span class="sxs-lookup"><span data-stu-id="054fd-112">[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).</span></span>
