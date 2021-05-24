---
title: Office脚本文件存储和所有权
description: 有关脚本Office和在所有者Microsoft OneDrive传输的信息。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545799"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office脚本文件存储和所有权

Office脚本作为 **.osts** 文件存储在你的Microsoft OneDrive。 它们独立于工作簿存储。 若要向其他人授予访问权限，[请与一个工作簿Excel脚本](excel.md#sharing-scripts)。 这意味着你要将脚本与文件链接，而不是附加它。 有权访问脚本Excel用户也能够查看、运行或制作脚本副本。

除非你共享脚本，否则其他人无法访问它们。 你的OneDrive设置控制所有脚本 **.osts** 文件的共享访问和权限，而不受Excel设置。 无法从本地磁盘或自定义云位置链接脚本。 Office脚本仅识别并运行脚本（如果它在 OneDrive文件夹中或与工作簿共享）。

## <a name="file-storage"></a>文件存储

You Office Scripts are stored in your OneDrive. **.osts** 文件位于 **/Documents/Office Scripts/** 文件夹中。 对这些 **.osts** 文件进行的任何编辑（如重命名或删除文件）都将反映在代码编辑器和脚本库中。

与工作簿之一共享的脚本将保留在脚本创建者的OneDrive。 在 OneDrive 中运行共享脚本时，不会将文件复制到任何本地或Excel。 代码 **编辑器的"创建** 副本"按钮会将脚本的单独副本保存在OneDrive。 对副本所做的更改不会影响原始脚本。

## <a name="file-ownership-and-retention"></a>文件所有权和保留

Office脚本存储在用户的 OneDrive。 它们遵循由用户指定的保留和删除Microsoft OneDrive。 若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。

在编辑过程中，文件会临时存储在浏览器中。 必须先保存脚本，然后再关闭Excel，以将其保存到OneDrive位置。 不要忘记在编辑后保存文件，否则这些编辑将仅在浏览器版本的文件中。

## <a name="see-also"></a>另请参阅

- [在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [M365 中的 Office 脚本设置](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [消除 Office 脚本的影响](../testing/undo.md)
