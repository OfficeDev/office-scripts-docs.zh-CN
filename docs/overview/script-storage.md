---
title: Office 脚本文件存储和所有权
description: 有关 Office 脚本如何存储在 Microsoft OneDrive 中并在所有者之间传输的信息。
ms.date: 08/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 573f65f299c29b4f481c9a2e23ebe7e36181706b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572505"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office 脚本文件存储和所有权

Office 脚本存储为 Microsoft OneDrive 或 SharePoint 文件夹中的 **.osts** 文件。 它们与工作簿分开存储。 若要向 SharePoint 网站外部的用户授予对脚本的访问权限，请 [使用 Excel 工作簿共享该脚本](excel.md#share-office-scripts)。 这意味着你要将脚本与文件链接，而不是附加它。 无论谁有权访问 Excel 文件，也都可以查看、运行或复制脚本。

仅当脚本位于 OneDrive 文件夹、Sharepoint 文件夹或与工作簿共享时，Excel 才会识别并运行该脚本。

## <a name="onedrive"></a>OneDrive

默认行为是 Office 脚本存储在 OneDrive 中。 **.osts** 文件位于 **/Documents/Office Scripts/** 文件夹中。 对这些 **.osts** 文件所做的任何编辑（例如重命名或删除文件）都将反映在代码编辑器和脚本库中。

与其中一个工作簿共享的脚本仍保留在脚本创建者的 OneDrive 中。 在 Excel 中运行共享脚本时，它们不会复制到任何本地或 OneDrive 文件夹。 代码编辑器的“ **创建复制** ”按钮将脚本的单独副本保存在 OneDrive 中。 对副本的更改不会影响原始脚本。

除非共享个人脚本，否则其他人无法访问它们。 OneDrive 设置控制所有脚本 **.osts** 文件的共享访问和权限，与任何 Excel 设置无关。 无法从本地磁盘或自定义云位置链接脚本。

## <a name="sharepoint"></a>SharePoint

保存到 SharePoint 网站的 Office 脚本归团队所有。 具有相应访问权限的组织成员可以从 SharePoint 运行和编辑脚本。 你还将看到这些脚本显示在 **“自动”** 选项卡的脚本库中。

若要从 SharePoint 加载脚本，请转到 **“所有脚本** ”，然后选择列表底部的 **“查看更多脚本** ”。 这会显示一个文件选取器，可在其中从你有权访问的任何 SharePoint 网站中选择 **.osts** 文件。 请注意，已打开的 SharePoint 脚本将显示在最近脚本列表中。

若要将脚本保存到 SharePoint，请转到 **“更多”选项 (...)** 菜单，然后选择 **“另存为**”。 这将打开文件选取器，可在其中选择 SharePoint 网站中的文件夹。 保存到新位置会在该位置创建脚本的副本。 原始版本仍位于 OneDrive 或其他 SharePoint 位置。

> [!IMPORTANT]
> 无法从 SharePoint 运行具有 [外部调](../develop/external-calls.md) 用的脚本。 你将收到一个错误，指出“目前不支持对保存到 SharePoint 网站的脚本进行网络访问调用”。

> [!IMPORTANT]
> Power Automate **目前不** 支持存储在 SharePoint 上的脚本。

## <a name="restore-deleted-scripts"></a>还原已删除的脚本

在 Excel 中删除脚本时，会转到 OneDrive 或 SharePoint 回收站。 若要还原已删除的脚本，请按照 [如何恢复 SharePoint 和 OneDrive 中用于工作或学校的缺失、已删除或损坏的项](https://support.microsoft.com/office/how-to-recover-missing-deleted-or-corrupted-items-in-sharepoint-and-onedrive-for-work-or-school-3d748edf-c072-46c9-81a4-4989056ebc87)所列步骤进行操作。 还原 **.osts** 文件会将其返回到 **“所有脚本”** 列表。

已删除的脚本与工作簿未共享。 还原脚本时， **它不会保留其** 脚本访问权限。 需要再次共享脚本。

还原的脚本仍按预期使用 Power Automate 流。 无需重新创建流连接器。

## <a name="file-ownership-and-retention"></a>文件所有权和保留

Office 脚本遵循 Microsoft OneDrive 和 Microsoft SharePoint 指定的保留和删除策略。 若要了解如何处理由从组织中删除的用户创建和共享的脚本，请 [参阅了解 SharePoint 和 OneDrive 的保留情况](/microsoft-365/compliance/retention-policies-sharepoint?view=o365-worldwide&preserve-view=true)。

在编辑过程中，文件暂时存储在浏览器中。 在关闭 Excel 窗口之前，必须保存脚本，才能将其保存到 OneDrive 位置。 不要忘记在编辑后保存文件，否则这些编辑将仅在浏览器的文件版本中。

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>管理级别的 Audit Office 脚本使用情况

使用合规中心审核日志发现组织中使用 Office 脚本的人员。 有关审核日志的详细信息，请参阅 [安全&合规中心中的审核日志](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log)。

若要以管理员身份专门审核 Office 脚本相关活动，请执行以下步骤。

1. 在 InPrivate 浏览器窗口 (或 Incognito 或其他特定于浏览器的受限跟踪模式) 中，打开并登录到 [合规中心](https://compliance.microsoft.com/)。
1. 转到 **“审核** ”页。
1. *(一次仅)* 在 **“搜索** ”选项卡上，选择 **“开始录制用户和管理员活动**”。

    > [!IMPORTANT]
    > 在打开录制后，可能需要一两个小时才能记录整个租户的所有活动。

1. 设置所需的搜索选项并按 **“搜索**”。 筛选 **工作簿上的“活动到运行”脚本**，以查看任何运行脚本的时间。 还可以将 **“文件”、“文件夹”或“站点”字段筛选为“文件”、“文件夹”或“站点**”字段。`.osts` 这会显示组织中创建或修改脚本的人员。

    :::image type="content" source="../images/audit-log-example.png" alt-text="几行审核日志搜索结果，包括“在工作簿上运行脚本”操作以及上传和修改 .osts 文件。":::

## <a name="see-also"></a>另请参阅

- [在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [M365 中的 Office 脚本设置](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [消除 Office 脚本的影响](../testing/undo.md)
