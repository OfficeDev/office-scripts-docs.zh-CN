---
title: 使用Office在Excel中运行脚本
description: 将按钮添加到工作簿，以控制Office脚本Excel。
ms.topic: overview
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0d88a6bcd928e6b4931b2374313cc17f4161ebf7
ms.sourcegitcommit: 49f527a7f54aba00e843ad4a92385af59c1d7bfa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63352195"
---
# <a name="run-office-scripts-in-excel-with-buttons"></a>使用Office在Excel中运行脚本

通过将脚本按钮添加到工作簿，帮助同事查找和运行脚本。

:::image type="content" source="../images/run-from-button.png" alt-text="单击时运行脚本的工作表中的一个按钮。":::

## <a name="create-script-buttons"></a>创建脚本按钮

对于任何脚本，请转到脚本详细信息页或代码编辑器任务窗格中的"更多选项" ( **...** ) "菜单，然后选择" **添加按钮"**。 此操作将在工作簿中创建一个按钮，已在选择该按钮时运行关联的脚本。 它还与工作簿共享脚本，因此对工作簿具有写入权限的每个人都可以使用有用的自动化操作。

以下屏幕截图显示了标题为"创建数据透视表"的脚本的"脚本详细信息"页，并且"其他选项"菜单中的"添加 (**...)** 突出显示。

:::image type="content" source="../images/add-button.png" alt-text="脚本详细信息页菜单中的“添加按钮”选项。":::

## <a name="remove-script-buttons"></a>删除脚本按钮

若要停止通过按钮共享脚本，请转到脚本的"详细信息 **" (...)** "菜单并选择"停止共享 **"**。 此操作将删除运行该脚本的所有按钮。 删除单个按钮会从该按钮中删除脚本，即使撤销该操作或剪切并粘贴该按钮也是如此。

## <a name="script-buttons-on-excel-for-windows"></a>用于Excel脚本Windows

这些脚本按钮也适用于 Windows。 Create the button in Excel web 版 and users on Windows can run your script with the click of a button. 请注意，您无法在脚本的 Excel 中Windows。 只能在脚本中编辑Excel web 版。

> [!NOTE]
> 此功能正在向具有 Microsoft 365 订阅的用户推出，并非所有人都可用。 我们缓慢地向更多用户发布此功能，以确保其正常工作。 此功能可能会根据你的反馈进行更改。 不支持的平台或不带此功能的 Office 版本将显示用于脚本按钮的形状，但无法单击该按钮。
