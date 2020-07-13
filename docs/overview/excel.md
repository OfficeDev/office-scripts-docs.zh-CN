---
title: Excel 网页版中的 Office 脚本
description: Office 脚本中的操作录制器和代码编辑器简介。
ms.date: 06/29/2020
localization_priority: Priority
ms.openlocfilehash: 046dd4eac0cce14117da75199841f0b2f72031bc
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043403"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="ec1e5-103">Excel 网页版中的 Office 脚本（预览版）</span><span class="sxs-lookup"><span data-stu-id="ec1e5-103">Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="ec1e5-104">Excel 网页版中的 Office 脚本可以让您可以自动化日常任务。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-104">Office Scripts in Excel on the web let you automate your day-to-day tasks.</span></span> <span data-ttu-id="ec1e5-105">你可以使用操作录制器录制 Excel 操作，这会创建一个脚本。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-105">You can record your Excel actions with the Action Recorder, which creates a script.</span></span> <span data-ttu-id="ec1e5-106">此外，你还可以使用代码编辑器创建和编辑脚本。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-106">You can also create and edit scripts with the Code Editor.</span></span> <span data-ttu-id="ec1e5-107">然后，可在组织中共享你的脚本，以便同事也可实现其工作流的自动化。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-107">Your scripts can then be shared across your organization so your coworkers can also automate their workflows.</span></span>

<span data-ttu-id="ec1e5-108">本文档系列将指导你如何使用这些工具。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-108">This series of documents teaches you how to use these tools.</span></span> <span data-ttu-id="ec1e5-109">我们将向你介绍操作录制器，让你了解如何录制频繁的 Excel 操作。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-109">You'll be introduced to the Action Recorder and see how to record your frequent Excel actions.</span></span> <span data-ttu-id="ec1e5-110">你还将学习如何使用代码编辑器创建或更新自己的脚本。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-110">You'll also learn how to make or update your own scripts with the Code Editor.</span></span>

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="when-to-use-office-scripts"></a><span data-ttu-id="ec1e5-111">何时使用 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="ec1e5-111">When to use Office Scripts</span></span>

<span data-ttu-id="ec1e5-112">你可以使用脚本录制和重播不同工作簿和工作表上的 Excel 操作。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-112">Scripts allow you to record and replay your Excel actions on different workbooks and worksheets.</span></span> <span data-ttu-id="ec1e5-113">如果你发现自己正在执行重复操作，则 Office 脚本可以将整个工作流程缩减为按一下按钮，从而为你提供帮助。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-113">If you find yourself doing the same things over and over again, an Office Script can help you by reducing your whole workflow to a single button press.</span></span>

<span data-ttu-id="ec1e5-114">例如，假如你在 Excel 中打开一个会计网站的 .csv 文件，以此开始一天的工作。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-114">As an example, say you start your work day by opening a .csv file from an accounting site in Excel.</span></span> <span data-ttu-id="ec1e5-115">你需要花几分钟删除不必要的列，设置表格格式，添加公式和在新工作表中创建一个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-115">You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet.</span></span> <span data-ttu-id="ec1e5-116">你可以使用操作录制器录制这些每天重复的操作。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-116">Those actions you repeat daily can be recorded once with the Action Recorder.</span></span> <span data-ttu-id="ec1e5-117">录制之后，运行脚本即可处理整个 .csv 转换。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-117">From then on, running the script will take care of your entire .csv conversion.</span></span> <span data-ttu-id="ec1e5-118">这样不仅可以消除忘记步骤的风险，而且还能够与他们共享流程，无需为他们提供任何指导。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-118">You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything.</span></span> <span data-ttu-id="ec1e5-119">Office 脚本可以自动化常见任务，使你和你的工作空间可以更有效率、更加高效。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-119">Office Scripts automate your common tasks so you and your workplace can be more efficient and productive.</span></span>

## <a name="action-recorder"></a><span data-ttu-id="ec1e5-120">操作录制器</span><span class="sxs-lookup"><span data-stu-id="ec1e5-120">Action Recorder</span></span>

![录制若干操作之后的操作录制器。](../images/action-recorder-intro.png)

<span data-ttu-id="ec1e5-122">操作录制器可以录制你在 Excel 中进行的操作，并将它们转换为脚本。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-122">The Action Recorder records actions you take in Excel and translates them into a script.</span></span> <span data-ttu-id="ec1e5-123">运行操作录制器之后，你可以在编辑单元格、更改格式和创建表格时捕获 Excel 操作。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-123">With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables.</span></span> <span data-ttu-id="ec1e5-124">可以在其他工作表和工作簿上运行生成的脚本，以重复创建原始操作。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-124">The resulting script can be run on other worksheets and workbooks to recreate your original actions.</span></span>

## <a name="code-editor"></a><span data-ttu-id="ec1e5-125">代码编辑器</span><span class="sxs-lookup"><span data-stu-id="ec1e5-125">Code Editor</span></span>

![显示以上脚本的脚本代码的代码编辑器。](../images/code-editor-intro.png)

<span data-ttu-id="ec1e5-127">使用操作录制器录制的所有脚本均可通过代码编辑器编辑。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-127">All scripts recorded with the Action Recorder can be edited through the Code Editor.</span></span> <span data-ttu-id="ec1e5-128">这使你能够调整和自定义脚本，以更好地满足你的准确需求。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-128">This lets you tweak and customize the script to better suit your exact needs.</span></span> <span data-ttu-id="ec1e5-129">此外，你还可以添加不能直接通过 Excel UI 访问的逻辑和功能，例如条件语句 (if/else) 和循环。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-129">You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.</span></span>

<span data-ttu-id="ec1e5-130">一种简单的开始学习 Office 脚本方式就是在 Excel 网页版上录制脚本，然后查看生成的代码。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-130">One easy way to start learning the capabilities of Office Scripts is to record scripts in Excel on the web and view the resulting code.</span></span> <span data-ttu-id="ec1e5-131">另一种选择是按照我们的[教程](../tutorials/excel-tutorial.md)进行，以更具指导性的结构化方式进行学习。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-131">Another option is to follow our [tutorials](../tutorials/excel-tutorial.md) to learn in a more guided and structured way.</span></span>

## <a name="sharing-scripts"></a><span data-ttu-id="ec1e5-132">共享脚本</span><span class="sxs-lookup"><span data-stu-id="ec1e5-132">Sharing scripts</span></span>

![显示“在此工作簿中与其他人共享”选项的脚本“详细信息”页面。](../images/script-sharing.png)

<span data-ttu-id="ec1e5-134">Office 脚本可与 Excel 工作簿的其他用户共享。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-134">Office Scripts can be shared with other users of an Excel workbook.</span></span> <span data-ttu-id="ec1e5-135">在工作簿中与其他人共享脚本时，该脚本将附加到工作簿中。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-135">When you share a script with others in a workbook, the script is attached to the workbook.</span></span> <span data-ttu-id="ec1e5-136">你的脚本存储在你的 OneDrive 中，当你共享一个脚本时，你将在打开的工作簿中创建指向该脚本的链接。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-136">Your scripts are stored in your OneDrive, and when you share one, you create a link to it in the workbook you have open.</span></span>

<span data-ttu-id="ec1e5-137">有关共享和取消共享脚本的详细信息，请参阅[在 Excel 网页版中共享 Office 脚本](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US)一文。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-137">More details about sharing and unsharing scripts can be in the article [Sharing Office Scripts in Excel for the Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US).</span></span>

## <a name="connecting-office-scripts-to-power-automate"></a><span data-ttu-id="ec1e5-138">将 Office 脚本连接到 Power Automate</span><span class="sxs-lookup"><span data-stu-id="ec1e5-138">Connecting Office Scripts to Power Automate</span></span>

<span data-ttu-id="ec1e5-139">[Power Automate](https://flow.microsoft.com/) 是一种可帮助你在多个应用和服务之间创建自动化工作流的服务。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-139">[Power Automate](https://flow.microsoft.com/) is a service that helps you create automated workflows between multiple apps and services.</span></span> <span data-ttu-id="ec1e5-140">Office 脚本可以在这些工作流中使用，以便你在工作簿之外控制脚本。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-140">Office Scripts can be used in these workflows, giving you control of your scripts outside of the workbook.</span></span> <span data-ttu-id="ec1e5-141">你可以按计划运行脚本，在回复电子邮件时触发它们，等等。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-141">You can run your scripts on a schedule, trigger them in response to emails, and much more.</span></span> <span data-ttu-id="ec1e5-142">若要了解有关连接这些自动化服务的基础知识，请访问[使用 Power Automate 在 Excel 网页版中运行 Office 脚本](../tutorials/excel-power-automate-manual.md)教程。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-142">Visit the [Run Office Scripts in Excel on the web with Power Automate](../tutorials/excel-power-automate-manual.md) tutorial to learn the basics of connecting these automation services.</span></span>

## <a name="next-steps"></a><span data-ttu-id="ec1e5-143">后续步骤</span><span class="sxs-lookup"><span data-stu-id="ec1e5-143">Next steps</span></span>

<span data-ttu-id="ec1e5-144">完成 [Excel 网页版上的 Office 脚本教程](../tutorials/excel-tutorial.md)，以了解如何创建你的第一个 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="ec1e5-144">Complete the [Office Scripts in Excel on the web tutorial](../tutorials/excel-tutorial.md) to learn how to create your first Office Scripts.</span></span>

## <a name="see-also"></a><span data-ttu-id="ec1e5-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ec1e5-145">See also</span></span>

- [<span data-ttu-id="ec1e5-146">Excel 网页版中 Office 脚本的脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="ec1e5-146">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="ec1e5-147">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="ec1e5-147">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="ec1e5-148">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="ec1e5-148">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="ec1e5-149">M365 中的 Office 脚本设置</span><span class="sxs-lookup"><span data-stu-id="ec1e5-149">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="ec1e5-150">Excel 中的 Office 脚本简介 (support.office.com)</span><span class="sxs-lookup"><span data-stu-id="ec1e5-150">Introduction to Office Scripts in Excel (on support.office.com)</span></span>](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [<span data-ttu-id="ec1e5-151">在 Excel 网页版中共享 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="ec1e5-151">Sharing Office Scripts in Excel for the Web</span></span>](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US)
