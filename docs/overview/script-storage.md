---
title: Office 脚本文件存储和所有权
description: 有关 Office 脚本如何存储在 Microsoft OneDrive 中以及如何在所有者之间传输的信息。
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: bd868c1dbfd0b33d3cd9fc4ee774c654d86f9b07
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755103"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office 脚本文件存储和所有权

Office 脚本存储为 Microsoft OneDrive 中的 **.osts** 文件。 这允许您的脚本存在于任何特定工作簿外部。 OneDrive 设置控制所有脚本 **.osts** 文件的共享访问和权限;独立于任何 Excel 设置。

## <a name="file-storage"></a>文件存储

Office 脚本存储在 OneDrive 中。 **.osts** 文件位于 **/Documents/Office Scripts/** 文件夹中。 对这些 **.osts** 文件进行的任何编辑（如重命名或删除文件）都将反映在代码编辑器和脚本库中。

与其中一个工作簿共享的脚本保留在脚本创建者的 OneDrive 中。 在 Excel 中运行共享脚本时，它们不会复制到任何本地或 OneDrive 文件夹。 代码 **编辑器的"复制** "按钮在 OneDrive 中保存脚本的单独副本。 对副本所做的更改不会影响原始脚本。

### <a name="script-folders"></a>脚本文件夹

将文件夹添加到 OneDrive 有助于保持脚本组织。 **/Documents/Office Scripts/** 下的任何文件夹都显示在代码编辑器的 **"我的脚本**"部分下。 请注意，不能使用代码编辑器创建或删除这些文件夹。 同样，脚本也不能放置在文件夹中，也不能使用代码编辑器跨文件夹移动。

:::image type="content" source="../images/script-folders.png" alt-text="代码编辑器中的&quot;新建脚本&quot;对话框显示文件夹中包含的脚本，如任务窗格中显示。":::

## <a name="file-ownership-and-retention"></a>文件所有权和保留

Office 脚本存储在用户的 OneDrive 中。 它们遵循 Microsoft OneDrive 指定的保留和删除策略。 若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。

## <a name="see-also"></a>另请参阅

- [在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [M365 中的 Office 脚本设置](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [消除 Office 脚本的影响](../testing/undo.md)
