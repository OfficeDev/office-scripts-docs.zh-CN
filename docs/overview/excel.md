---
title: Excel 网页版中的 Office 脚本
description: Office 脚本中的操作录制器和代码编辑器简介。
ms.date: 09/29/2020
localization_priority: Priority
ms.openlocfilehash: 965e28be285d59d79d46fe005ab16f29b271041f
ms.sourcegitcommit: ce72354381561dc167ea0092efd915642a9161b3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/30/2020
ms.locfileid: "48319670"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="e4a51-103">Excel 网页版中的 Office 脚本（预览版）</span><span class="sxs-lookup"><span data-stu-id="e4a51-103">Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="e4a51-104">Excel 网页版中的 Office 脚本可以让您可以自动化日常任务。</span><span class="sxs-lookup"><span data-stu-id="e4a51-104">Office Scripts in Excel on the web let you automate your day-to-day tasks.</span></span> <span data-ttu-id="e4a51-105">你可以使用操作录制器录制 Excel 操作，这会创建一个脚本。</span><span class="sxs-lookup"><span data-stu-id="e4a51-105">You can record your Excel actions with the Action Recorder, which creates a script.</span></span> <span data-ttu-id="e4a51-106">此外，你还可以使用代码编辑器创建和编辑脚本。</span><span class="sxs-lookup"><span data-stu-id="e4a51-106">You can also create and edit scripts with the Code Editor.</span></span> <span data-ttu-id="e4a51-107">然后，可在组织中共享你的脚本，以便同事也可实现其工作流的自动化。</span><span class="sxs-lookup"><span data-stu-id="e4a51-107">Your scripts can then be shared across your organization so your coworkers can also automate their workflows.</span></span>

<span data-ttu-id="e4a51-108">本文档系列将指导你如何使用这些工具。</span><span class="sxs-lookup"><span data-stu-id="e4a51-108">This series of documents teaches you how to use these tools.</span></span> <span data-ttu-id="e4a51-109">我们将向你介绍操作录制器，让你了解如何录制频繁的 Excel 操作。</span><span class="sxs-lookup"><span data-stu-id="e4a51-109">You'll be introduced to the Action Recorder and see how to record your frequent Excel actions.</span></span> <span data-ttu-id="e4a51-110">你还将学习如何使用代码编辑器创建或更新自己的脚本。</span><span class="sxs-lookup"><span data-stu-id="e4a51-110">You'll also learn how to make or update your own scripts with the Code Editor.</span></span>

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a><span data-ttu-id="e4a51-111">Requirements</span><span class="sxs-lookup"><span data-stu-id="e4a51-111">Requirements</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

<span data-ttu-id="e4a51-112">若要使用 Office 脚本，需要以下内容。</span><span class="sxs-lookup"><span data-stu-id="e4a51-112">To use Office Scripts, you'll need the following.</span></span>

1. <span data-ttu-id="e4a51-113">[Excel 网页版](https://www.office.com/launch/excel)（不支持桌面等其他平台）。</span><span class="sxs-lookup"><span data-stu-id="e4a51-113">[Excel on the web](https://www.office.com/launch/excel) (other platforms, such as desktop, are not supported).</span></span>
1. <span data-ttu-id="e4a51-114">OneDrive for Business。</span><span class="sxs-lookup"><span data-stu-id="e4a51-114">OneDrive for Business.</span></span>
1. <span data-ttu-id="e4a51-115">[管理员已启用](/microsoft-365/admin/manage/manage-office-scripts-settings) Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="e4a51-115">Office Scripts [enabled by your administrator](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>
1. <span data-ttu-id="e4a51-116">可访问 Microsoft 365 Office 桌面应用的任何商业版或教育版 Microsoft 365 许可证，例如：</span><span class="sxs-lookup"><span data-stu-id="e4a51-116">Any commercial or educational Microsoft 365 license with access to the Microsoft 365 Office desktop apps, such as:</span></span>

    - <span data-ttu-id="e4a51-117">Office 365 商业版</span><span class="sxs-lookup"><span data-stu-id="e4a51-117">Office 365 Business</span></span>
    - <span data-ttu-id="e4a51-118">Office 365 商业高级版</span><span class="sxs-lookup"><span data-stu-id="e4a51-118">Office 365 Business Premium</span></span>
    - <span data-ttu-id="e4a51-119">Office 365 专业增强版</span><span class="sxs-lookup"><span data-stu-id="e4a51-119">Office 365 ProPlus</span></span>
    - <span data-ttu-id="e4a51-120">Office 365 专业增强版（设备）</span><span class="sxs-lookup"><span data-stu-id="e4a51-120">Office 365 ProPlus for Devices</span></span>
    - <span data-ttu-id="e4a51-121">Office 365 企业版 E3</span><span class="sxs-lookup"><span data-stu-id="e4a51-121">Office 365 Enterprise E3</span></span>
    - <span data-ttu-id="e4a51-122">Office 365 企业版 E5</span><span class="sxs-lookup"><span data-stu-id="e4a51-122">Office 365 Enterprise E5</span></span>
    - <span data-ttu-id="e4a51-123">Office 365 A3</span><span class="sxs-lookup"><span data-stu-id="e4a51-123">Office 365 A3</span></span>
    - <span data-ttu-id="e4a51-124">Office 365 A5</span><span class="sxs-lookup"><span data-stu-id="e4a51-124">Office 365 A5</span></span>

## <a name="when-to-use-office-scripts"></a><span data-ttu-id="e4a51-125">何时使用 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="e4a51-125">When to use Office Scripts</span></span>

<span data-ttu-id="e4a51-126">你可以使用脚本录制和重播不同工作簿和工作表上的 Excel 操作。</span><span class="sxs-lookup"><span data-stu-id="e4a51-126">Scripts allow you to record and replay your Excel actions on different workbooks and worksheets.</span></span> <span data-ttu-id="e4a51-127">如果你发现自己正在重复执行相同的操作，则可以将所有工作转变为易于运行的 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="e4a51-127">If you find yourself doing the same things over and over again, you can turn all that work into an easy-to-run Office Script.</span></span> <span data-ttu-id="e4a51-128">通过 Excel 中的一个按钮运行脚本，或将其与 Power Automate 结合使用，简化整个工作流程。</span><span class="sxs-lookup"><span data-stu-id="e4a51-128">Run your script with a button-press in Excel or combine it with Power Automate to streamline your entire workflow.</span></span>

<span data-ttu-id="e4a51-129">例如，假如你在 Excel 中打开一个会计网站的 .csv 文件，以此开始一天的工作。</span><span class="sxs-lookup"><span data-stu-id="e4a51-129">As an example, say you start your work day by opening a .csv file from an accounting site in Excel.</span></span> <span data-ttu-id="e4a51-130">你需要花几分钟删除不必要的列，设置表格格式，添加公式和在新工作表中创建一个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="e4a51-130">You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet.</span></span> <span data-ttu-id="e4a51-131">你可以使用操作录制器录制这些每天重复的操作。</span><span class="sxs-lookup"><span data-stu-id="e4a51-131">Those actions you repeat daily can be recorded once with the Action Recorder.</span></span> <span data-ttu-id="e4a51-132">录制之后，运行脚本即可处理整个 .csv 转换。</span><span class="sxs-lookup"><span data-stu-id="e4a51-132">From then on, running the script will take care of your entire .csv conversion.</span></span> <span data-ttu-id="e4a51-133">这样不仅可以消除忘记步骤的风险，而且还能够与他们共享流程，无需为他们提供任何指导。</span><span class="sxs-lookup"><span data-stu-id="e4a51-133">You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything.</span></span> <span data-ttu-id="e4a51-134">Office 脚本可以自动化常见任务，使你和你的工作空间可以更有效率、更加高效。</span><span class="sxs-lookup"><span data-stu-id="e4a51-134">Office Scripts automate your common tasks so you and your workplace can be more efficient and productive.</span></span>

## <a name="action-recorder"></a><span data-ttu-id="e4a51-135">操作录制器</span><span class="sxs-lookup"><span data-stu-id="e4a51-135">Action Recorder</span></span>

![录制若干操作之后的操作录制器。](../images/action-recorder-intro.png)

<span data-ttu-id="e4a51-137">操作录制器可以录制你在 Excel 中进行的操作，并将它们转换为脚本。</span><span class="sxs-lookup"><span data-stu-id="e4a51-137">The Action Recorder records actions you take in Excel and saves them as a script.</span></span> <span data-ttu-id="e4a51-138">运行操作录制器之后，你可以在编辑单元格、更改格式和创建表格时捕获 Excel 操作。</span><span class="sxs-lookup"><span data-stu-id="e4a51-138">With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables.</span></span> <span data-ttu-id="e4a51-139">可以在其他工作表和工作簿上运行生成的脚本，以重复创建原始操作。</span><span class="sxs-lookup"><span data-stu-id="e4a51-139">The resulting script can be run on other worksheets and workbooks to recreate your original actions.</span></span>

## <a name="code-editor"></a><span data-ttu-id="e4a51-140">代码编辑器</span><span class="sxs-lookup"><span data-stu-id="e4a51-140">Code Editor</span></span>

![显示以上脚本的脚本代码的代码编辑器。](../images/code-editor-intro.png)

<span data-ttu-id="e4a51-142">使用操作录制器录制的所有脚本均可通过代码编辑器编辑。</span><span class="sxs-lookup"><span data-stu-id="e4a51-142">All scripts recorded with the Action Recorder can be edited through the Code Editor.</span></span> <span data-ttu-id="e4a51-143">这使你能够调整和自定义脚本，以更好地满足你的准确需求。</span><span class="sxs-lookup"><span data-stu-id="e4a51-143">This lets you tweak and customize the script to better suit your exact needs.</span></span> <span data-ttu-id="e4a51-144">此外，你还可以添加不能直接通过 Excel UI 访问的逻辑和功能，例如条件语句 (if/else) 和循环。</span><span class="sxs-lookup"><span data-stu-id="e4a51-144">You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.</span></span>

<span data-ttu-id="e4a51-145">一种简单的开始学习 Office 脚本方式就是在 Excel 网页版上录制脚本，然后查看生成的代码。</span><span class="sxs-lookup"><span data-stu-id="e4a51-145">One easy way to start learning the capabilities of Office Scripts is to record scripts in Excel on the web and view the resulting code.</span></span> <span data-ttu-id="e4a51-146">另一种选择是按照我们的[教程](../tutorials/excel-tutorial.md)进行，以更具指导性的结构化方式进行学习。</span><span class="sxs-lookup"><span data-stu-id="e4a51-146">Another option is to follow our [tutorials](../tutorials/excel-tutorial.md) to learn in a more guided and structured way.</span></span>

## <a name="sharing-scripts"></a><span data-ttu-id="e4a51-147">共享脚本</span><span class="sxs-lookup"><span data-stu-id="e4a51-147">Sharing scripts</span></span>

![显示“在此工作簿中与其他人共享”选项的脚本“详细信息”页面。](../images/script-sharing.png)

<span data-ttu-id="e4a51-149">Office 脚本可与 Excel 工作簿的其他用户共享。</span><span class="sxs-lookup"><span data-stu-id="e4a51-149">Office Scripts can be shared with other users of an Excel workbook.</span></span> <span data-ttu-id="e4a51-150">在工作簿中与其他人共享脚本时，该脚本将附加到工作簿中。</span><span class="sxs-lookup"><span data-stu-id="e4a51-150">When you share a script with others in a workbook, the script is attached to the workbook.</span></span> <span data-ttu-id="e4a51-151">你的脚本存储在你的 OneDrive 中，当你共享一个脚本时，你将在打开的工作簿中创建指向该脚本的链接。</span><span class="sxs-lookup"><span data-stu-id="e4a51-151">Your scripts are stored in your OneDrive, and when you share one, you create a link to it in the workbook you have open.</span></span>

<span data-ttu-id="e4a51-152">有关共享和取消共享脚本的详细信息，请参阅[在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)一文。</span><span class="sxs-lookup"><span data-stu-id="e4a51-152">More details about sharing and unsharing scripts can be in the article [Sharing Office Scripts in Excel for the Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b).</span></span>

> [!NOTE]
> <span data-ttu-id="e4a51-153">由于 Office 脚本存储在用户的 OneDrive 上，因此它们遵循相同的保留和删除策略。</span><span class="sxs-lookup"><span data-stu-id="e4a51-153">Since Office Scripts are stored on a user's OneDrive, they follow the same retention and deletion policies.</span></span> <span data-ttu-id="e4a51-154">若要了解如何处理从组织中删除了用户所创建和共享的脚本，请参阅 [OneDrive 保留和删除](/onedrive/retention-and-deletion)。</span><span class="sxs-lookup"><span data-stu-id="e4a51-154">To learn how to handle scripts that were created and shared by a user being removed from your organization, see [OneDrive retention and deletion](/onedrive/retention-and-deletion).</span></span>

## <a name="connecting-office-scripts-to-power-automate"></a><span data-ttu-id="e4a51-155">将 Office 脚本连接到 Power Automate</span><span class="sxs-lookup"><span data-stu-id="e4a51-155">Connecting Office Scripts to Power Automate</span></span>

<span data-ttu-id="e4a51-156">[Power Automate](https://flow.microsoft.com/) 是一种可帮助你在多个应用和服务之间创建自动化工作流的服务。</span><span class="sxs-lookup"><span data-stu-id="e4a51-156">[Power Automate](https://flow.microsoft.com/) is a service that helps you create automated workflows between multiple apps and services.</span></span> <span data-ttu-id="e4a51-157">Office 脚本可以在这些工作流中使用，以便你在工作簿之外控制脚本。</span><span class="sxs-lookup"><span data-stu-id="e4a51-157">Office Scripts can be used in these workflows, giving you control of your scripts outside of the workbook.</span></span> <span data-ttu-id="e4a51-158">你可以按计划运行脚本，在回复电子邮件时触发它们，等等。</span><span class="sxs-lookup"><span data-stu-id="e4a51-158">You can run your scripts on a schedule, trigger them in response to emails, and much more.</span></span> <span data-ttu-id="e4a51-159">若要了解有关连接这些自动化服务的基础知识，请访问[使用 Power Automate 在 Excel 网页版中运行 Office 脚本](../tutorials/excel-power-automate-manual.md)教程。</span><span class="sxs-lookup"><span data-stu-id="e4a51-159">Visit the [Run Office Scripts in Excel on the web with Power Automate](../tutorials/excel-power-automate-manual.md) tutorial to learn the basics of connecting these automation services.</span></span>

## <a name="next-steps"></a><span data-ttu-id="e4a51-160">后续步骤</span><span class="sxs-lookup"><span data-stu-id="e4a51-160">Next steps</span></span>

<span data-ttu-id="e4a51-161">完成 [Excel 网页版上的 Office 脚本教程](../tutorials/excel-tutorial.md)，以了解如何创建你的第一个 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="e4a51-161">Complete the [Office Scripts in Excel on the web tutorial](../tutorials/excel-tutorial.md) to learn how to create your first Office Scripts.</span></span>

## <a name="see-also"></a><span data-ttu-id="e4a51-162">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e4a51-162">See also</span></span>

- [<span data-ttu-id="e4a51-163">Excel 网页版中 Office 脚本的脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="e4a51-163">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="e4a51-164">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="e4a51-164">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="e4a51-165">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="e4a51-165">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="e4a51-166">M365 中的 Office 脚本设置</span><span class="sxs-lookup"><span data-stu-id="e4a51-166">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="e4a51-167">Excel 中的 Office 脚本简介 (support.office.com)</span><span class="sxs-lookup"><span data-stu-id="e4a51-167">Introduction to Office Scripts in Excel (on support.office.com)</span></span>](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [<span data-ttu-id="e4a51-168">在 Excel 网页版中共享 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="e4a51-168">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US)
