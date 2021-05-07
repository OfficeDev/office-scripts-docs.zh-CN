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
# <a name="office-scripts-file-storage-and-ownership"></a><span data-ttu-id="b758f-103">Office脚本文件存储和所有权</span><span class="sxs-lookup"><span data-stu-id="b758f-103">Office Scripts file storage and ownership</span></span>

<span data-ttu-id="b758f-104">Office脚本作为 **.osts** 文件存储在你的Microsoft OneDrive。</span><span class="sxs-lookup"><span data-stu-id="b758f-104">Office Scripts are stored as **.osts** files in your Microsoft OneDrive.</span></span> <span data-ttu-id="b758f-105">这允许您的脚本存在于任何特定工作簿外部。</span><span class="sxs-lookup"><span data-stu-id="b758f-105">This allows your scripts to exist outside any particular workbook.</span></span> <span data-ttu-id="b758f-106">你的OneDrive设置控制所有脚本 **.osts** 文件的共享访问和权限;独立于任何Excel设置。</span><span class="sxs-lookup"><span data-stu-id="b758f-106">Your OneDrive settings control the shared access and permissions for all script **.osts** files; independent of any Excel settings.</span></span>

## <a name="file-storage"></a><span data-ttu-id="b758f-107">文件存储</span><span class="sxs-lookup"><span data-stu-id="b758f-107">File storage</span></span>

<span data-ttu-id="b758f-108">You Office Scripts are stored in your OneDrive.</span><span class="sxs-lookup"><span data-stu-id="b758f-108">You Office Scripts are stored in your OneDrive.</span></span> <span data-ttu-id="b758f-109">**.osts** 文件位于 **/Documents/Office Scripts/** 文件夹中。</span><span class="sxs-lookup"><span data-stu-id="b758f-109">The **.osts** files are found in the **/Documents/Office Scripts/** folder.</span></span> <span data-ttu-id="b758f-110">对这些 **.osts** 文件进行的任何编辑（如重命名或删除文件）都将反映在代码编辑器和脚本库中。</span><span class="sxs-lookup"><span data-stu-id="b758f-110">Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.</span></span>

<span data-ttu-id="b758f-111">与工作簿之一共享的脚本将保留在脚本创建者的OneDrive。</span><span class="sxs-lookup"><span data-stu-id="b758f-111">Scripts that are shared with one of your workbooks remain in the script creator's OneDrive.</span></span> <span data-ttu-id="b758f-112">在 OneDrive 中运行共享脚本时，不会将文件复制到任何本地或Excel。</span><span class="sxs-lookup"><span data-stu-id="b758f-112">They are not copied to any of your local or OneDrive folders when you run the shared script in Excel.</span></span> <span data-ttu-id="b758f-113">代码 **编辑器的"创建** 副本"按钮会将脚本的单独副本保存在OneDrive。</span><span class="sxs-lookup"><span data-stu-id="b758f-113">The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive.</span></span> <span data-ttu-id="b758f-114">对副本所做的更改不会影响原始脚本。</span><span class="sxs-lookup"><span data-stu-id="b758f-114">Changes to the copy don't affect the original script.</span></span>

### <a name="script-folders"></a><span data-ttu-id="b758f-115">脚本文件夹</span><span class="sxs-lookup"><span data-stu-id="b758f-115">Script folders</span></span>

<span data-ttu-id="b758f-116">将文件夹添加到OneDrive有助于保持脚本的条理。</span><span class="sxs-lookup"><span data-stu-id="b758f-116">Adding folders to your OneDrive helps keep your scripts organized.</span></span> <span data-ttu-id="b758f-117">**/Documents/Office Scripts/** 下的任何文件夹都显示在代码编辑器的 **"我的** 脚本"部分下。</span><span class="sxs-lookup"><span data-stu-id="b758f-117">Any folders under **/Documents/Office Scripts/** are displayed under the **My Scripts** section of the Code Editor.</span></span> <span data-ttu-id="b758f-118">请注意，不能使用代码编辑器创建或删除这些文件夹。</span><span class="sxs-lookup"><span data-stu-id="b758f-118">Please note that these folders cannot be created or deleted by using the Code Editor.</span></span> <span data-ttu-id="b758f-119">同样，脚本也不能放置在文件夹中，也不能使用代码编辑器跨文件夹移动。</span><span class="sxs-lookup"><span data-stu-id="b758f-119">Likewise, scripts cannot be placed in folders, or moved across folders by using the Code Editor.</span></span>

:::image type="content" source="../images/script-folders.png" alt-text="代码编辑器中的&quot;新建脚本&quot;对话框显示文件夹中包含的脚本，如任务窗格中显示":::

## <a name="file-ownership-and-retention"></a><span data-ttu-id="b758f-121">文件所有权和保留</span><span class="sxs-lookup"><span data-stu-id="b758f-121">File ownership and retention</span></span>

<span data-ttu-id="b758f-122">Office脚本存储在用户的 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="b758f-122">Office Scripts are stored in a user's OneDrive.</span></span> <span data-ttu-id="b758f-123">它们遵循由用户指定的保留和删除Microsoft OneDrive。</span><span class="sxs-lookup"><span data-stu-id="b758f-123">They follow the retention and deletion policies specified by Microsoft OneDrive.</span></span> <span data-ttu-id="b758f-124">若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。</span><span class="sxs-lookup"><span data-stu-id="b758f-124">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

## <a name="see-also"></a><span data-ttu-id="b758f-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b758f-125">See also</span></span>

- [<span data-ttu-id="b758f-126">在 Excel 网页版中共享 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="b758f-126">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [<span data-ttu-id="b758f-127">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="b758f-127">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="b758f-128">M365 中的 Office 脚本设置</span><span class="sxs-lookup"><span data-stu-id="b758f-128">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="b758f-129">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="b758f-129">Undo the effects of an Office Script</span></span>](../testing/undo.md)
