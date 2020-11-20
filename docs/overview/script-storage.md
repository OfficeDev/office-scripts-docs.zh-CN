---
title: Office 脚本文件存储和所有权
description: 有关 Office 脚本存储在 Microsoft OneDrive 中并在所有者之间进行传输的信息。
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 648f3b2cf7e7d8d3bab2cf07a090e116e267a99a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49346860"
---
# <a name="office-scripts-file-storage-and-ownership"></a><span data-ttu-id="36cb7-103">Office 脚本文件存储和所有权</span><span class="sxs-lookup"><span data-stu-id="36cb7-103">Office Scripts file storage and ownership</span></span>

<span data-ttu-id="36cb7-104">Office 脚本以 **. osts** 文件的形式存储在 Microsoft OneDrive 中。</span><span class="sxs-lookup"><span data-stu-id="36cb7-104">Office Scripts are stored as **.osts** files in your Microsoft OneDrive.</span></span> <span data-ttu-id="36cb7-105">这将允许您的脚本位于任何特定工作簿的外部。</span><span class="sxs-lookup"><span data-stu-id="36cb7-105">This allows your scripts to exist outside any particular workbook.</span></span> <span data-ttu-id="36cb7-106">您的 OneDrive 设置控制所有 **osts** 文件的共享访问权限和权限。独立于任何 Excel 设置。</span><span class="sxs-lookup"><span data-stu-id="36cb7-106">Your OneDrive settings control the shared access and permissions for all script **.osts** files; independent of any Excel settings.</span></span>

## <a name="file-storage"></a><span data-ttu-id="36cb7-107">文件存储</span><span class="sxs-lookup"><span data-stu-id="36cb7-107">File storage</span></span>

<span data-ttu-id="36cb7-108">你的 Office 脚本存储在你的 OneDrive 中。</span><span class="sxs-lookup"><span data-stu-id="36cb7-108">You Office Scripts are stored in your OneDrive.</span></span> <span data-ttu-id="36cb7-109">在 **/Documents/Office 脚本/** 文件夹中找到 **osts** 文件。</span><span class="sxs-lookup"><span data-stu-id="36cb7-109">The **.osts** files are found in the **/Documents/Office Scripts/** folder.</span></span> <span data-ttu-id="36cb7-110">对这些 **osts** 文件所做的任何编辑（如重命名或删除文件）都将反映在代码编辑器和脚本库中。</span><span class="sxs-lookup"><span data-stu-id="36cb7-110">Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.</span></span>

<span data-ttu-id="36cb7-111">与其中一个工作簿共享的脚本将保留在脚本创建者的 OneDrive 中。</span><span class="sxs-lookup"><span data-stu-id="36cb7-111">Scripts that are shared with one of your workbooks remain in the script creator's OneDrive.</span></span> <span data-ttu-id="36cb7-112">当您在 Excel 中运行共享脚本时，不会将它们复制到您的任何本地或 OneDrive 文件夹中。</span><span class="sxs-lookup"><span data-stu-id="36cb7-112">They are not copied to any of your local or OneDrive folders when you run the shared script in Excel.</span></span> <span data-ttu-id="36cb7-113">" **创建** 代码编辑器的副本" 按钮在 OneDrive 中保存脚本的单独副本。</span><span class="sxs-lookup"><span data-stu-id="36cb7-113">The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive.</span></span> <span data-ttu-id="36cb7-114">对副本所做的更改不会影响原始脚本。</span><span class="sxs-lookup"><span data-stu-id="36cb7-114">Changes to the copy don't affect the original script.</span></span>

### <a name="script-folders"></a><span data-ttu-id="36cb7-115">脚本文件夹</span><span class="sxs-lookup"><span data-stu-id="36cb7-115">Script folders</span></span>

<span data-ttu-id="36cb7-116">将文件夹添加到你的 OneDrive 有助于组织组织的脚本。</span><span class="sxs-lookup"><span data-stu-id="36cb7-116">Adding folders to your OneDrive helps keep your scripts organized.</span></span> <span data-ttu-id="36cb7-117">" **/Documents/Office scripts/** " 下的任何文件夹都显示在代码编辑器的 " **我的脚本** " 部分下。</span><span class="sxs-lookup"><span data-stu-id="36cb7-117">Any folders under **/Documents/Office Scripts/** are displayed under the **My Scripts** section of the Code Editor.</span></span> <span data-ttu-id="36cb7-118">请注意，无法使用代码编辑器创建或删除这些文件夹。</span><span class="sxs-lookup"><span data-stu-id="36cb7-118">Please note that these folders cannot be created or deleted by using the Code Editor.</span></span> <span data-ttu-id="36cb7-119">同样，脚本也不能放在文件夹中，也不能通过使用代码编辑器在文件夹中移动。</span><span class="sxs-lookup"><span data-stu-id="36cb7-119">Likewise, scripts cannot be placed in folders, or moved across folders by using the Code Editor.</span></span>

![在 "代码编辑器" 任务窗格中显示的文件夹中包含的一些脚本](../images/script-folders.png)

## <a name="file-ownership-and-retention"></a><span data-ttu-id="36cb7-121">文件所有权和保留</span><span class="sxs-lookup"><span data-stu-id="36cb7-121">File ownership and retention</span></span>

<span data-ttu-id="36cb7-122">Office 脚本存储在用户的 OneDrive 中。</span><span class="sxs-lookup"><span data-stu-id="36cb7-122">Office Scripts are stored in a user's OneDrive.</span></span> <span data-ttu-id="36cb7-123">它们遵循 Microsoft OneDrive 指定的保留和删除策略。</span><span class="sxs-lookup"><span data-stu-id="36cb7-123">They follow the retention and deletion policies specified by Microsoft OneDrive.</span></span> <span data-ttu-id="36cb7-124">若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。</span><span class="sxs-lookup"><span data-stu-id="36cb7-124">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

## <a name="see-also"></a><span data-ttu-id="36cb7-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="36cb7-125">See also</span></span>

- [<span data-ttu-id="36cb7-126">在 Excel 网页版中共享 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="36cb7-126">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [<span data-ttu-id="36cb7-127">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="36cb7-127">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="36cb7-128">M365 中的 Office 脚本设置</span><span class="sxs-lookup"><span data-stu-id="36cb7-128">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="36cb7-129">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="36cb7-129">Undo the effects of an Office Script</span></span>](../testing/undo.md)
