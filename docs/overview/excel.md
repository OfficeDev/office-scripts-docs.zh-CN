---
title: Excel 网页版中的 Office 脚本
description: Office 脚本中的操作录制器和代码编辑器简介。
ms.date: 02/24/2020
localization_priority: Priority
ms.openlocfilehash: fb1d32068f9a738bb99412c2892cf22b4119b9b1
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978347"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="e4e86-103">Excel 网页版中的 Office 脚本（预览版）</span><span class="sxs-lookup"><span data-stu-id="e4e86-103">Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="e4e86-104">Excel 网页版中的 Office 脚本可以让您可以自动化日常任务。</span><span class="sxs-lookup"><span data-stu-id="e4e86-104">Office Scripts in Excel on the web let you automate your day-to-day tasks.</span></span> <span data-ttu-id="e4e86-105">你可以使用操作录制器录制 Excel 操作，这会创建一个脚本。</span><span class="sxs-lookup"><span data-stu-id="e4e86-105">You can record your Excel actions with the Action Recorder, which creates a script.</span></span> <span data-ttu-id="e4e86-106">此外，你还可以使用代码编辑器创建和编辑脚本。</span><span class="sxs-lookup"><span data-stu-id="e4e86-106">You can also create and edit scripts with the Code Editor.</span></span> <span data-ttu-id="e4e86-107">本文档系列将指导你如何使用这些工具。</span><span class="sxs-lookup"><span data-stu-id="e4e86-107">This series of documents teaches you how to use these tools.</span></span> <span data-ttu-id="e4e86-108">我们将向你介绍操作录制器，让你了解如何录制频繁的 Excel 操作。</span><span class="sxs-lookup"><span data-stu-id="e4e86-108">You'll be introduced to the Action Recorder and see how to record your frequent Excel actions.</span></span> <span data-ttu-id="e4e86-109">你还将学习如何使用代码编辑器创建或更新自己的脚本。</span><span class="sxs-lookup"><span data-stu-id="e4e86-109">You'll also learn how to make or update your own scripts with the Code Editor.</span></span>

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="when-to-use-office-scripts"></a><span data-ttu-id="e4e86-110">何时使用 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="e4e86-110">When to use Office Scripts</span></span>

<span data-ttu-id="e4e86-111">你可以使用脚本录制和重播不同工作簿和工作表上的 Excel 操作。</span><span class="sxs-lookup"><span data-stu-id="e4e86-111">Scripts allow you to record and replay your Excel actions on different workbooks and worksheets.</span></span> <span data-ttu-id="e4e86-112">如果你发现自己正在执行重复操作，则 Office 脚本可以将整个工作流程缩减为按一下按钮，从而为你提供帮助。</span><span class="sxs-lookup"><span data-stu-id="e4e86-112">If you find yourself doing the same things over and over again, an Office Script can help you by reducing your whole workflow to a single button press.</span></span>

<span data-ttu-id="e4e86-113">例如，假如你在 Excel 中打开一个会计网站的 .csv 文件，以此开始一天的工作。</span><span class="sxs-lookup"><span data-stu-id="e4e86-113">As an example, say you start your work day by opening a .csv file from an accounting site in Excel.</span></span> <span data-ttu-id="e4e86-114">你需要花几分钟删除不必要的列，设置表格格式，添加公式和在新工作表中创建一个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="e4e86-114">You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet.</span></span> <span data-ttu-id="e4e86-115">你可以使用操作录制器录制这些每天重复的操作。</span><span class="sxs-lookup"><span data-stu-id="e4e86-115">Those actions you repeat daily can be recorded once with the Action Recorder.</span></span> <span data-ttu-id="e4e86-116">录制之后，运行脚本即可处理整个 .csv 转换。</span><span class="sxs-lookup"><span data-stu-id="e4e86-116">From then on, running the script will take care of your entire .csv conversion.</span></span> <span data-ttu-id="e4e86-117">这样不仅可以消除忘记步骤的风险，而且还能够与他们共享流程，无需为他们提供任何指导。</span><span class="sxs-lookup"><span data-stu-id="e4e86-117">You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything.</span></span> <span data-ttu-id="e4e86-118">Office 脚本可以自动化常见任务，使你和你的工作空间可以更有效率、更加高效。</span><span class="sxs-lookup"><span data-stu-id="e4e86-118">Office Scripts automate your common tasks so you and your workplace can be more efficient and productive.</span></span>

## <a name="action-recorder"></a><span data-ttu-id="e4e86-119">操作录制器</span><span class="sxs-lookup"><span data-stu-id="e4e86-119">Action Recorder</span></span>

![录制若干操作之后的操作录制器。](../images/action-recorder-intro.png)

<span data-ttu-id="e4e86-121">操作录制器可以录制你在 Excel 中进行的操作，并将它们转换为脚本。</span><span class="sxs-lookup"><span data-stu-id="e4e86-121">The Action Recorder records actions you take in Excel and translates them into a script.</span></span> <span data-ttu-id="e4e86-122">运行操作录制器之后，你可以在编辑单元格、更改格式和创建表格时捕获 Excel 操作。</span><span class="sxs-lookup"><span data-stu-id="e4e86-122">With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables.</span></span> <span data-ttu-id="e4e86-123">可以在其他工作表和工作簿上运行生成的脚本，以重复创建原始操作。</span><span class="sxs-lookup"><span data-stu-id="e4e86-123">The resulting script can be run on other worksheets and workbooks to recreate your original actions.</span></span>

## <a name="code-editor"></a><span data-ttu-id="e4e86-124">代码编辑器</span><span class="sxs-lookup"><span data-stu-id="e4e86-124">Code Editor</span></span>

![显示以上脚本的脚本代码的代码编辑器。](../images/code-editor-intro.png)

<span data-ttu-id="e4e86-126">使用操作录制器录制的所有脚本均可通过代码编辑器编辑。</span><span class="sxs-lookup"><span data-stu-id="e4e86-126">All scripts recorded with the Action Recorder can be edited through the Code Editor.</span></span> <span data-ttu-id="e4e86-127">这使你能够调整和自定义脚本，以更好地满足你的准确需求。</span><span class="sxs-lookup"><span data-stu-id="e4e86-127">This lets you tweak and customize the script to better suit your exact needs.</span></span> <span data-ttu-id="e4e86-128">此外，你还可以添加不能直接通过 Excel UI 访问的逻辑和功能，例如条件语句 (if/else) 和循环。</span><span class="sxs-lookup"><span data-stu-id="e4e86-128">You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.</span></span>

<span data-ttu-id="e4e86-129">一种简单的开始学习 Office 脚本方式就是在 Excel 网页版上录制脚本，然后查看生成的代码。</span><span class="sxs-lookup"><span data-stu-id="e4e86-129">One easy way to start learning the capabilities of Office Scripts is to record scripts in Excel on the web and view the resulting code.</span></span> <span data-ttu-id="e4e86-130">另一种选择是按照我们的[教程](../tutorials/excel-tutorial.md)进行，以更具指导性的结构化方式进行学习。</span><span class="sxs-lookup"><span data-stu-id="e4e86-130">Another option is to follow our [tutorials](../tutorials/excel-tutorial.md) to learn in a more guided and structured way.</span></span>

## <a name="next-steps"></a><span data-ttu-id="e4e86-131">后续步骤</span><span class="sxs-lookup"><span data-stu-id="e4e86-131">Next steps</span></span>

<span data-ttu-id="e4e86-132">完成 [Excel 网页版上的 Office 脚本教程](../tutorials/excel-tutorial.md)，以了解如何创建你的第一个 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="e4e86-132">Complete the [Office Scripts in Excel on the web tutorial](../tutorials/excel-tutorial.md) to learn how to create your first Office Scripts.</span></span>

## <a name="see-also"></a><span data-ttu-id="e4e86-133">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e4e86-133">See also</span></span>

- [<span data-ttu-id="e4e86-134">Excel 网页版上的 Office 脚本的脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="e4e86-134">Scripting fundamentals for Office Script in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="e4e86-135">Office 脚本 API 参考</span><span class="sxs-lookup"><span data-stu-id="e4e86-135">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="e4e86-136">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="e4e86-136">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="e4e86-137">M365 中的 Office 脚本设置</span><span class="sxs-lookup"><span data-stu-id="e4e86-137">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="e4e86-138">Excel 中的 Office 脚本简介 (support.office.com)</span><span class="sxs-lookup"><span data-stu-id="e4e86-138">Introduction to Office Scripts in Excel (on support.office.com)</span></span>](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
