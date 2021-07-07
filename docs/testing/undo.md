---
title: 撤消由脚本Office所做的更改
description: 使用脚本的版本历史记录Excel web 版运行脚本来撤消所做的更改。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 86ce59ea4715ac6d8b56ca8d165a1e0451e4ee22
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313965"
---
# <a name="undo-the-changes-made-by-office-scripts"></a><span data-ttu-id="296bb-103">撤消由脚本Office所做的更改</span><span class="sxs-lookup"><span data-stu-id="296bb-103">Undo the changes made by Office Scripts</span></span>

<span data-ttu-id="296bb-104">无法使用"撤消"命令Excel脚本撤消对工作簿Excel **所做的更改**。</span><span class="sxs-lookup"><span data-stu-id="296bb-104">You cannot undo changes made to the Excel workbook by a script with the Excel's **Undo** command.</span></span> <span data-ttu-id="296bb-105">相反，您必须从云存储还原以前版本的工作簿。</span><span class="sxs-lookup"><span data-stu-id="296bb-105">Instead, you must restore a previous version of the workbook from your cloud storage.</span></span>

## <a name="version-history"></a><span data-ttu-id="296bb-106">版本历史记录</span><span class="sxs-lookup"><span data-stu-id="296bb-106">Version history</span></span>

<span data-ttu-id="296bb-107">Office版本历史记录是一种通过自定义 UI 还原旧工作簿Excel方法。</span><span class="sxs-lookup"><span data-stu-id="296bb-107">Office's version history is an easy way to restore an older workbook through the Excel UI.</span></span> <span data-ttu-id="296bb-108">该功能仅适用于存储在 OneDrive 或 SharePoint Online 中的文件。</span><span class="sxs-lookup"><span data-stu-id="296bb-108">The feature only works for files stored in OneDrive or SharePoint Online.</span></span>

<span data-ttu-id="296bb-109">从Excel脚本的工作簿中，使用以下步骤撤消效果：</span><span class="sxs-lookup"><span data-stu-id="296bb-109">From the Excel workbook in which the script was ran, use these steps to undo the effects:</span></span>

1. <span data-ttu-id="296bb-110">转到文件  >  **信息**  >  **版本历史记录**。</span><span class="sxs-lookup"><span data-stu-id="296bb-110">Go to **File** > **Info** > **Version History**.</span></span>
2. <span data-ttu-id="296bb-111">选择在运行脚本之前保存的版本。</span><span class="sxs-lookup"><span data-stu-id="296bb-111">Select a version saved prior to the running the script.</span></span>
3. <span data-ttu-id="296bb-112">选择"**还原"。**</span><span class="sxs-lookup"><span data-stu-id="296bb-112">Select **Restore**.</span></span>

## <a name="see-also"></a><span data-ttu-id="296bb-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="296bb-113">See also</span></span>

- [<span data-ttu-id="296bb-114">查看先前版本的 Office 文件</span><span class="sxs-lookup"><span data-stu-id="296bb-114">View previous versions of Office files</span></span>](https://support.office.com/article/View-previous-versions-of-Office-files-5c1e076f-a9c9-41b8-8ace-f77b9642e2c2#ID0EABBAAA=Web)
