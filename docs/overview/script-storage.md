---
title: Office脚本文件存储和所有权
description: 有关脚本Office和在所有者Microsoft OneDrive传输的信息。
ms.date: 06/04/2021
ms.localizationpriority: medium
ms.openlocfilehash: 98c10ed8def417bef36d5a97eb5411648d49258e
ms.sourcegitcommit: 49f527a7f54aba00e843ad4a92385af59c1d7bfa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63352130"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office脚本文件存储和所有权

Office脚本作为 **.osts** 文件存储在你的Microsoft OneDrive。 它们独立于工作簿存储。 若要向其他人授予访问权限，[请与一个工作簿Excel脚本](excel.md#share-office-scripts)。 这意味着你要将脚本与文件链接，而不是附加它。 有权访问脚本Excel用户也能够查看、运行或制作脚本副本。

除非你共享脚本，否则其他人无法访问它们。 你的OneDrive设置控制所有脚本 **.osts** 文件的共享访问和权限，而不受Excel设置。 无法从本地磁盘或自定义云位置链接脚本。 Office脚本仅识别并运行脚本（如果它在 OneDrive文件夹中或与工作簿共享）。

## <a name="file-storage"></a>文件存储

You Office Scripts are stored in your OneDrive. **.osts** 文件位于 **/Documents/Office Scripts/** 文件夹中。 对这些 **.osts** 文件进行的任何编辑（如重命名或删除文件）都将反映在代码编辑器和脚本库中。

与工作簿之一共享的脚本将保留在脚本创建者的OneDrive。 在 Excel 中运行共享脚本时，它们不会复制到任何本地或OneDrive文件夹。 代码 **编辑器的"创建** 副本"按钮将脚本的单独副本保存在OneDrive。 对副本所做的更改不会影响原始脚本。

### <a name="restore-deleted-scripts"></a>还原已删除的脚本

在删除脚本时，Excel脚本将转到OneDrive回收站。 若要还原已删除的脚本，请按照还原已删除的文件或文件夹中[OneDrive。](https://support.microsoft.com/office/949ada80-0026-4db3-a953-c99083e6a84f) 还原 **.osts** 文件会返回到"所有脚本 **"** 列表。

已删除的脚本未与工作簿共享。 还原脚本时，它不会 **保留** 其脚本访问权限。 你将需要再次共享脚本。

还原的脚本仍像预期的那样与Power Automate一工作。 无需重新创建流连接器。

## <a name="file-ownership-and-retention"></a>文件所有权和保留

Office脚本存储在用户的 OneDrive。 它们遵循由用户指定的保留和删除Microsoft OneDrive。 若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。

在编辑过程中，文件会临时存储在浏览器中。 必须先保存脚本，然后再关闭Excel，以将其保存到OneDrive位置。 不要忘记在编辑后保存文件，否则这些编辑将仅在浏览器版本的文件中。

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>审核Office级别的脚本使用情况

发现哪些租户正在使用Office脚本审核日志合规中心中的脚本。 若要了解如何使用此工具，请访问安全与合规审核日志[搜索&搜索。](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log)

若要查找将脚本Office搜索工具，请添加`.osts`"文件"、**文件夹或网站** 字段。 这将搜索所有扩展名为 Office Scripts 的文件。 如果组织中的任何人已使用 Office 脚本功能，用户活动会显示在审核日志搜索结果中。

> [!NOTE]
> 当前未记录运行脚本。 仅记录创建、查看和修改操作。

## <a name="see-also"></a>另请参阅

- [在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [M365 中的 Office 脚本设置](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [消除 Office 脚本的影响](../testing/undo.md)
