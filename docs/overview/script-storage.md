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
# <a name="office-scripts-file-storage-and-ownership"></a><span data-ttu-id="0d872-103">Office 脚本文件存储和所有权</span><span class="sxs-lookup"><span data-stu-id="0d872-103">Office Scripts file storage and ownership</span></span>

<span data-ttu-id="0d872-104">Office 脚本存储为 Microsoft OneDrive 中的 **.osts** 文件。</span><span class="sxs-lookup"><span data-stu-id="0d872-104">Office Scripts are stored as **.osts** files in your Microsoft OneDrive.</span></span> <span data-ttu-id="0d872-105">这允许您的脚本存在于任何特定工作簿外部。</span><span class="sxs-lookup"><span data-stu-id="0d872-105">This allows your scripts to exist outside any particular workbook.</span></span> <span data-ttu-id="0d872-106">OneDrive 设置控制所有脚本 **.osts** 文件的共享访问和权限;独立于任何 Excel 设置。</span><span class="sxs-lookup"><span data-stu-id="0d872-106">Your OneDrive settings control the shared access and permissions for all script **.osts** files; independent of any Excel settings.</span></span>

## <a name="file-storage"></a><span data-ttu-id="0d872-107">文件存储</span><span class="sxs-lookup"><span data-stu-id="0d872-107">File storage</span></span>

<span data-ttu-id="0d872-108">Office 脚本存储在 OneDrive 中。</span><span class="sxs-lookup"><span data-stu-id="0d872-108">You Office Scripts are stored in your OneDrive.</span></span> <span data-ttu-id="0d872-109">**.osts** 文件位于 **/Documents/Office Scripts/** 文件夹中。</span><span class="sxs-lookup"><span data-stu-id="0d872-109">The **.osts** files are found in the **/Documents/Office Scripts/** folder.</span></span> <span data-ttu-id="0d872-110">对这些 **.osts** 文件进行的任何编辑（如重命名或删除文件）都将反映在代码编辑器和脚本库中。</span><span class="sxs-lookup"><span data-stu-id="0d872-110">Any edits made to these **.osts** files, such as renaming or deleting files, will be reflected in the Code Editor and Script Gallery.</span></span>

<span data-ttu-id="0d872-111">与其中一个工作簿共享的脚本保留在脚本创建者的 OneDrive 中。</span><span class="sxs-lookup"><span data-stu-id="0d872-111">Scripts that are shared with one of your workbooks remain in the script creator's OneDrive.</span></span> <span data-ttu-id="0d872-112">在 Excel 中运行共享脚本时，它们不会复制到任何本地或 OneDrive 文件夹。</span><span class="sxs-lookup"><span data-stu-id="0d872-112">They are not copied to any of your local or OneDrive folders when you run the shared script in Excel.</span></span> <span data-ttu-id="0d872-113">代码 **编辑器的"复制** "按钮在 OneDrive 中保存脚本的单独副本。</span><span class="sxs-lookup"><span data-stu-id="0d872-113">The **Make a Copy** button of the Code Editor saves a separate copy of the script in your OneDrive.</span></span> <span data-ttu-id="0d872-114">对副本所做的更改不会影响原始脚本。</span><span class="sxs-lookup"><span data-stu-id="0d872-114">Changes to the copy don't affect the original script.</span></span>

### <a name="script-folders"></a><span data-ttu-id="0d872-115">脚本文件夹</span><span class="sxs-lookup"><span data-stu-id="0d872-115">Script folders</span></span>

<span data-ttu-id="0d872-116">将文件夹添加到 OneDrive 有助于保持脚本组织。</span><span class="sxs-lookup"><span data-stu-id="0d872-116">Adding folders to your OneDrive helps keep your scripts organized.</span></span> <span data-ttu-id="0d872-117">**/Documents/Office Scripts/** 下的任何文件夹都显示在代码编辑器的 **"我的脚本**"部分下。</span><span class="sxs-lookup"><span data-stu-id="0d872-117">Any folders under **/Documents/Office Scripts/** are displayed under the **My Scripts** section of the Code Editor.</span></span> <span data-ttu-id="0d872-118">请注意，不能使用代码编辑器创建或删除这些文件夹。</span><span class="sxs-lookup"><span data-stu-id="0d872-118">Please note that these folders cannot be created or deleted by using the Code Editor.</span></span> <span data-ttu-id="0d872-119">同样，脚本也不能放置在文件夹中，也不能使用代码编辑器跨文件夹移动。</span><span class="sxs-lookup"><span data-stu-id="0d872-119">Likewise, scripts cannot be placed in folders, or moved across folders by using the Code Editor.</span></span>

:::image type="content" source="../images/script-folders.png" alt-text="代码编辑器中的&quot;新建脚本&quot;对话框显示文件夹中包含的脚本，如任务窗格中显示。":::

## <a name="file-ownership-and-retention"></a><span data-ttu-id="0d872-121">文件所有权和保留</span><span class="sxs-lookup"><span data-stu-id="0d872-121">File ownership and retention</span></span>

<span data-ttu-id="0d872-122">Office 脚本存储在用户的 OneDrive 中。</span><span class="sxs-lookup"><span data-stu-id="0d872-122">Office Scripts are stored in a user's OneDrive.</span></span> <span data-ttu-id="0d872-123">它们遵循 Microsoft OneDrive 指定的保留和删除策略。</span><span class="sxs-lookup"><span data-stu-id="0d872-123">They follow the retention and deletion policies specified by Microsoft OneDrive.</span></span> <span data-ttu-id="0d872-124">若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。</span><span class="sxs-lookup"><span data-stu-id="0d872-124">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

## <a name="see-also"></a><span data-ttu-id="0d872-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0d872-125">See also</span></span>

- [<span data-ttu-id="0d872-126">在 Excel 网页版中共享 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="0d872-126">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [<span data-ttu-id="0d872-127">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="0d872-127">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="0d872-128">M365 中的 Office 脚本设置</span><span class="sxs-lookup"><span data-stu-id="0d872-128">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="0d872-129">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="0d872-129">Undo the effects of an Office Script</span></span>](../testing/undo.md)
