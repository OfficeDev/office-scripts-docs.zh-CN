---
title: 撤消由脚本Office所做的更改
description: 使用脚本的版本历史记录Excel web 版运行脚本来撤消所做的更改。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8d8a86a2361f7a5eb58cd3900b488df57d889770
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586022"
---
# <a name="undo-the-changes-made-by-office-scripts"></a>撤消由脚本Office所做的更改

不能使用"撤消"命令Excel脚本撤消对Excel工作簿 **所做的更改**。 相反，您必须从云存储还原以前版本的工作簿。

## <a name="version-history"></a>版本历史记录

Office版本历史记录是一种通过自定义 UI 还原旧工作簿Excel方法。 此功能仅适用于存储在 OneDrive 或 SharePoint Online 中的文件。

从Excel脚本的工作簿中，使用以下步骤撤消效果：

1. 转到"**FileInfoVersion** >  >  **历史记录"**。
2. 选择在运行脚本之前保存的版本。
3. 选择 **"还原"**。

## <a name="see-also"></a>另请参阅

- [查看先前版本的 Office 文件](https://support.office.com/article/View-previous-versions-of-Office-files-5c1e076f-a9c9-41b8-8ace-f77b9642e2c2#ID0EABBAAA=Web)
