---
title: Office脚本文件存储和所有权
description: 有关脚本Office和在所有者Microsoft OneDrive传输的信息。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545799"
---
# <a name="office-scripts-file-storage-and-ownership"></a><span data-ttu-id="6fc8f-103">Office脚本文件存储和所有权</span><span class="sxs-lookup"><span data-stu-id="6fc8f-103">Office Scripts file storage and ownership</span></span>

<span data-ttu-id="6fc8f-104">Office脚本作为 **.osts** 文件存储在你的Microsoft OneDrive。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-104">Office Scripts are stored as **.osts** files in your Microsoft OneDrive.</span></span> <span data-ttu-id="6fc8f-105">它们独立于工作簿存储。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-105">They are stored separately from a workbook.</span></span> <span data-ttu-id="6fc8f-106">若要向其他人授予访问权限，[请与一个工作簿Excel脚本](excel.md#sharing-scripts)。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-106">To give others access, [share the script with an Excel workbook](excel.md#sharing-scripts).</span></span> <span data-ttu-id="6fc8f-107">这意味着你要将脚本与文件链接，而不是附加它。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-107">This means you're linking the script with the file, not attaching it.</span></span> <span data-ttu-id="6fc8f-108">有权访问脚本Excel用户也能够查看、运行或制作脚本副本。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-108">Whoever has access to the Excel file will also be able to view, run, or make a copy of the script.</span></span>

<span data-ttu-id="6fc8f-109">除非你共享脚本，否则其他人无法访问它们。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-109">Unless you share your scripts, no one else can access them.</span></span> <span data-ttu-id="6fc8f-110">你的OneDrive设置控制所有脚本 **.osts** 文件的共享访问和权限，而不受Excel设置。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-110">Your OneDrive settings control the shared access and permissions for all script **.osts** files, independent of any Excel settings.</span></span> <span data-ttu-id="6fc8f-111">无法从本地磁盘或自定义云位置链接脚本。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-111">Scripts can't be linked from a local disk or custom cloud locations.</span></span> <span data-ttu-id="6fc8f-112">Office脚本仅识别并运行脚本（如果它在 OneDrive文件夹中或与工作簿共享）。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-112">Office Scripts only recognizes and runs a script if it's in your OneDrive folder or shared with the workbook.</span></span>

## <a name="file-storage"></a><span data-ttu-id="6fc8f-113">文件存储</span><span class="sxs-lookup"><span data-stu-id="6fc8f-113">File storage</span></span>

<span data-ttu-id="6fc8f-114">You Office Scripts are stored in your OneDrive.</span><span class="sxs-lookup"><span data-stu-id="6fc8f-114">You Office Scripts are stored in your OneDrive.</span></span> <span data-ttu-id="6fc8f-115">**.osts** 文件位于 **/Documents/Office Scripts/** 文件夹中。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-115">The **.osts** files are found in the **/Documents/Office Scripts/** folder.</span></span> <span data-ttu-id="6fc8f-116">对这些 **.osts** 文件进行的任何编辑（如重命名或删除文件）都将反映在代码编辑器和脚本库中。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-116">Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.</span></span>

<span data-ttu-id="6fc8f-117">与工作簿之一共享的脚本将保留在脚本创建者的OneDrive。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-117">Scripts that are shared with one of your workbooks remain in the script creator's OneDrive.</span></span> <span data-ttu-id="6fc8f-118">在 OneDrive 中运行共享脚本时，不会将文件复制到任何本地或Excel。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-118">They are not copied to any of your local or OneDrive folders when you run the shared script in Excel.</span></span> <span data-ttu-id="6fc8f-119">代码 **编辑器的"创建** 副本"按钮会将脚本的单独副本保存在OneDrive。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-119">The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive.</span></span> <span data-ttu-id="6fc8f-120">对副本所做的更改不会影响原始脚本。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-120">Changes to the copy don't affect the original script.</span></span>

## <a name="file-ownership-and-retention"></a><span data-ttu-id="6fc8f-121">文件所有权和保留</span><span class="sxs-lookup"><span data-stu-id="6fc8f-121">File ownership and retention</span></span>

<span data-ttu-id="6fc8f-122">Office脚本存储在用户的 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-122">Office Scripts are stored in a user's OneDrive.</span></span> <span data-ttu-id="6fc8f-123">它们遵循由用户指定的保留和删除Microsoft OneDrive。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-123">They follow the retention and deletion policies specified by Microsoft OneDrive.</span></span> <span data-ttu-id="6fc8f-124">若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-124">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

<span data-ttu-id="6fc8f-125">在编辑过程中，文件会临时存储在浏览器中。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-125">During editing, files are temporarily stored in the browser.</span></span> <span data-ttu-id="6fc8f-126">必须先保存脚本，然后再关闭Excel，以将其保存到OneDrive位置。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-126">You must save the script before closing the Excel window to save it to the OneDrive location.</span></span> <span data-ttu-id="6fc8f-127">不要忘记在编辑后保存文件，否则这些编辑将仅在浏览器版本的文件中。</span><span class="sxs-lookup"><span data-stu-id="6fc8f-127">Don't forget to save the file after edits, or else those edits will only be in the browser's version of the file.</span></span>

## <a name="see-also"></a><span data-ttu-id="6fc8f-128">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6fc8f-128">See also</span></span>

- [<span data-ttu-id="6fc8f-129">在 Excel 网页版中共享 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="6fc8f-129">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [<span data-ttu-id="6fc8f-130">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="6fc8f-130">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="6fc8f-131">M365 中的 Office 脚本设置</span><span class="sxs-lookup"><span data-stu-id="6fc8f-131">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="6fc8f-132">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="6fc8f-132">Undo the effects of Office Scripts</span></span>](../testing/undo.md)
