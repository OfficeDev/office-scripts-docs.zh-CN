---
title: 撤消由脚本Office所做的更改
description: 使用脚本的版本历史记录Excel web 版运行脚本来撤消所做的更改。
ms.date: 01/08/2019
localization_priority: Normal
ms.openlocfilehash: f9f22d4879f8a02c00a5bac9f58d9aa36ae03e38
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545544"
---
# <a name="undo-the-changes-made-by-office-scripts"></a><span data-ttu-id="397a9-103">撤消由脚本Office所做的更改</span><span class="sxs-lookup"><span data-stu-id="397a9-103">Undo the changes made by Office Scripts</span></span>

<span data-ttu-id="397a9-104">无法使用"撤消"命令Excel脚本撤消对工作簿Excel **所做的更改**。</span><span class="sxs-lookup"><span data-stu-id="397a9-104">You cannot undo changes made to the Excel workbook by a script with the Excel's **Undo** command.</span></span> <span data-ttu-id="397a9-105">相反，您必须从云存储还原以前版本的工作簿。</span><span class="sxs-lookup"><span data-stu-id="397a9-105">Instead, you must restore a previous version of the workbook from your cloud storage.</span></span>

## <a name="version-history"></a><span data-ttu-id="397a9-106">版本历史记录</span><span class="sxs-lookup"><span data-stu-id="397a9-106">Version history</span></span>

<span data-ttu-id="397a9-107">Office版本历史记录是一种通过自定义 UI 还原旧工作簿Excel方法。</span><span class="sxs-lookup"><span data-stu-id="397a9-107">Office's version history is an easy way to restore an older workbook through the Excel UI.</span></span> <span data-ttu-id="397a9-108">该功能仅适用于存储在 OneDrive 或 SharePoint Online 中的文件。</span><span class="sxs-lookup"><span data-stu-id="397a9-108">The feature only works for files stored in OneDrive or SharePoint Online.</span></span>

<span data-ttu-id="397a9-109">从Excel脚本的工作簿中，使用以下步骤撤消效果：</span><span class="sxs-lookup"><span data-stu-id="397a9-109">From the Excel workbook in which the script was ran, use these steps to undo the effects:</span></span>

1. <span data-ttu-id="397a9-110">转到文件  >  **信息**  >  **版本历史记录**。</span><span class="sxs-lookup"><span data-stu-id="397a9-110">Go to **File** > **Info** > **Version History**.</span></span>
2. <span data-ttu-id="397a9-111">选择在运行脚本之前保存的版本。</span><span class="sxs-lookup"><span data-stu-id="397a9-111">Select a version saved prior to the running the script.</span></span>
3. <span data-ttu-id="397a9-112">按 **"还原"。**</span><span class="sxs-lookup"><span data-stu-id="397a9-112">Press **Restore**.</span></span>

## <a name="see-also"></a><span data-ttu-id="397a9-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="397a9-113">See also</span></span>

- [<span data-ttu-id="397a9-114">查看先前版本的 Office 文件</span><span class="sxs-lookup"><span data-stu-id="397a9-114">View previous versions of Office files</span></span>](https://support.office.com/article/View-previous-versions-of-Office-files-5c1e076f-a9c9-41b8-8ace-f77b9642e2c2#ID0EABBAAA=Web)
