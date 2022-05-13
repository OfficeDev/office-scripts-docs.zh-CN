---
title: 使用按钮在Excel中运行Office脚本
description: 将按钮添加到控制Excel中Office脚本的工作簿。
ms.topic: overview
ms.date: 05/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: fde34d62f9abe897a8b93195ab37a75cfc73f619
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393682"
---
# <a name="run-office-scripts-in-excel-with-buttons"></a>使用按钮在Excel中运行Office脚本

通过将脚本按钮添加到工作簿，帮助同事查找和运行脚本。

:::image type="content" source="../images/run-from-button.png" alt-text="单击时运行脚本的工作表中的一个按钮。":::

## <a name="create-script-buttons"></a>创建脚本按钮

使用任何脚本，请转到脚本的详细信息页或“代码编辑器”任务窗格中的“ **更多”选项 (...)** 菜单，然后选择 **“添加”按钮**。 此操作将在工作簿中创建一个按钮，已在选择该按钮时运行关联的脚本。 它还与工作簿共享脚本，因此对工作簿具有写入权限的每个人都可以使用有用的自动化操作。

以下屏幕截图显示了名为 **“创建数据透视表**”的脚本的脚本详细信息页，并在“**更多选项 (...)** 菜单中突出显示了 **”添加“按钮** 选项。

:::image type="content" source="../images/add-button.png" alt-text="脚本详细信息页菜单中的“添加按钮”选项。":::

## <a name="remove-script-buttons"></a>删除脚本按钮

若要停止通过按钮共享脚本，请转到“ **更多”选项 (...)** 脚本详细信息页中的菜单，然后选择 **“停止共享**”。 此操作将删除运行该脚本的所有按钮。 删除单个按钮会从该按钮中删除脚本，即使撤销该操作或剪切并粘贴该按钮也是如此。

## <a name="script-buttons-with-excel-on-windows"></a>Windows上带有Excel的脚本按钮

这些脚本按钮也适用于 Windows。 在Excel web 版中创建按钮，Windows上的用户可以通过单击按钮来运行脚本。 请注意，无法在Windows Excel中编辑脚本。 只能在Excel web 版中编辑脚本。

某些Office脚本 API 可能不受Windows上的Excel支持，尤其是较旧的版本。 其中包括用于仅 Web 功能的较新的 API 和 API。 如果脚本包含不受支持的 API，则该脚本不会运行，相反，“**脚本运行状态**”任务窗格会显示一条警告消息，指出“此脚本当前必须在Excel 网页版上运行。 在浏览器中打开工作簿，然后重试，或联系脚本所有者寻求帮助。  

> [!IMPORTANT]
> 脚本按钮需要 [WebView2](/deployoffice/webview2-install) 在Windows上使用Excel。 默认情况下，在桌面上安装最新版本的Excel，但如果无法单击脚本按钮，请访问[“下载 WebView2 运行时](https://developer.microsoft.com/en-us/microsoft-edge/webview2/#download-section)”并下载浏览器引擎。
