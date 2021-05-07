---
title: Office脚本文件存储和所有权
description: 有关脚本Office和在所有者Microsoft OneDrive传输的信息。
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 47b732399c3068bea78b027e01324bbd73a83bc7
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232527"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office脚本文件存储和所有权

Office脚本作为 **.osts** 文件存储在你的Microsoft OneDrive。 这允许您的脚本存在于任何特定工作簿外部。 你的OneDrive设置控制所有脚本 **.osts** 文件的共享访问和权限;独立于任何Excel设置。

## <a name="file-storage"></a>文件存储

You Office Scripts are stored in your OneDrive. **.osts** 文件位于 **/Documents/Office Scripts/** 文件夹中。 对这些 **.osts** 文件进行的任何编辑（如重命名或删除文件）都将反映在代码编辑器和脚本库中。

与工作簿之一共享的脚本将保留在脚本创建者的OneDrive。 在 OneDrive 中运行共享脚本时，不会将文件复制到任何本地或Excel。 代码 **编辑器的"创建** 副本"按钮会将脚本的单独副本保存在OneDrive。 对副本所做的更改不会影响原始脚本。

### <a name="script-folders"></a>脚本文件夹

将文件夹添加到OneDrive有助于保持脚本的条理。 **/Documents/Office Scripts/** 下的任何文件夹都显示在代码编辑器的 **"我的** 脚本"部分下。 请注意，不能使用代码编辑器创建或删除这些文件夹。 同样，脚本也不能放置在文件夹中，也不能使用代码编辑器跨文件夹移动。

:::image type="content" source="../images/script-folders.png" alt-text="代码编辑器中的&quot;新建脚本&quot;对话框显示文件夹中包含的脚本，如任务窗格中显示":::

## <a name="file-ownership-and-retention"></a>文件所有权和保留

Office脚本存储在用户的 OneDrive。 它们遵循由用户指定的保留和删除Microsoft OneDrive。 若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。

## <a name="see-also"></a>另请参阅

- [在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [M365 中的 Office 脚本设置](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [消除 Office 脚本的影响](../testing/undo.md)
