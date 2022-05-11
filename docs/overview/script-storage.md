---
title: Office脚本文件存储和所有权
description: 有关Office脚本如何存储在Microsoft OneDrive中并在所有者之间传输的信息。
ms.date: 05/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5e2bc89db54ee5520c3b911ebd0f182777a78e2b
ms.sourcegitcommit: 8ae932e8b4e521fec8576ab16126eb9fe22a8dd7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2022
ms.locfileid: "65310755"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office脚本文件存储和所有权

Office脚本存储为Microsoft OneDrive中的 **.osts** 文件。 它们与工作簿分开存储。 若要授予其他人访问权限，请[使用Excel工作簿共享脚本](excel.md#share-office-scripts)。 这意味着你要将脚本与文件链接，而不是附加它。 无论谁有权访问Excel文件，也都可以查看、运行或复制脚本。

除非共享脚本，否则其他人无法访问它们。 OneDrive设置控制所有脚本 **.osts 文件的** 共享访问权限和权限，与任何Excel设置无关。 无法从本地磁盘或自定义云位置链接脚本。 Office脚本仅在OneDrive文件夹中或与工作簿共享时识别并运行脚本。

## <a name="file-storage"></a>文件存储

Office脚本存储在OneDrive中。 **.osts** 文件位于 **/Documents/Office Scripts/** 文件夹中。 对这些 **.osts** 文件所做的任何编辑（例如重命名或删除文件）都将反映在代码编辑器和脚本库中。

与其中一个工作簿共享的脚本仍保留在脚本创建者的OneDrive中。 在Excel中运行共享脚本时，它们不会复制到任何本地或OneDrive文件夹。 代码编辑器的“**创建复制**”按钮会在OneDrive中保存脚本的单独副本。 对副本的更改不会影响原始脚本。

### <a name="restore-deleted-scripts"></a>还原已删除的脚本

在Excel中删除脚本时，它会转到OneDrive回收站。 若要还原已删除的脚本，请按照[OneDrive中“还原已删除的文件或文件夹”中](https://support.microsoft.com/office/949ada80-0026-4db3-a953-c99083e6a84f)列出的步骤操作。 还原 **.osts** 文件会将其返回到 **“所有脚本”** 列表。

已删除的脚本与工作簿未共享。 还原脚本时， **它不会保留其** 脚本访问权限。 需要再次共享脚本。

还原的脚本仍按预期使用Power Automate流。 无需重新创建流连接器。

## <a name="file-ownership-and-retention"></a>文件所有权和保留

Office脚本存储在用户的OneDrive中。 它们遵循Microsoft OneDrive指定的保留和删除策略。 若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。

在编辑过程中，文件暂时存储在浏览器中。 在关闭Excel窗口之前，必须保存该脚本，以便将其保存到OneDrive位置。 不要忘记在编辑后保存文件，否则这些编辑将仅在浏览器的文件版本中。

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>在管理级别审核Office脚本使用情况

发现哪些租户在合规中心使用Office脚本和审核日志。 若要了解如何使用此工具，请访问[安全&合规中心中的审核日志。](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log)

若要查找在搜索工具中使用Office脚本的人员，请在 **“文件”、“文件夹”或“网站”** 字段中添加`.osts`。 这会搜索具有Office脚本文件扩展名的所有文件。 如果组织中有人使用了Office脚本功能，则用户活动会显示在审核日志搜索结果中。

## <a name="see-also"></a>另请参阅

- [在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [M365 中的 Office 脚本设置](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [消除 Office 脚本的影响](../testing/undo.md)
