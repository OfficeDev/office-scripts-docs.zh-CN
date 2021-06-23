---
title: Office 脚本与 Office 加载项之间的差异
description: 脚本和加载项Office API 的行为Office API 差异。
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: 17d66e37c7bf2b1263c0232bb0afb3ee4d29aa36
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074562"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="40faf-103">Office 脚本与 Office 加载项之间的差异</span><span class="sxs-lookup"><span data-stu-id="40faf-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="40faf-104">了解Office脚本Office外接程序之间的差异，以了解何时使用每个脚本和外接程序。</span><span class="sxs-lookup"><span data-stu-id="40faf-104">Understand the differences between Office Scripts and Office Add-ins to know when to use each one.</span></span> <span data-ttu-id="40faf-105">Office脚本旨在让任何希望改进其工作流的人快速创建脚本。</span><span class="sxs-lookup"><span data-stu-id="40faf-105">Office Scripts are designed to be quickly made by anyone looking to improve their workflow.</span></span> <span data-ttu-id="40faf-106">Office外接程序与 Office UI 集成，通过功能区按钮和任务窗格实现更具交互性的体验。</span><span class="sxs-lookup"><span data-stu-id="40faf-106">Office Add-ins integrate with the Office UI for a more interactive experience through ribbon buttons and task panes.</span></span> <span data-ttu-id="40faf-107">Office加载项还可以通过提供自定义函数来扩展内置Excel函数。</span><span class="sxs-lookup"><span data-stu-id="40faf-107">Office Add-ins can also expand built-in Excel functions by providing custom functions.</span></span>

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="四象限图，显示不同扩展性解决方案Office区域。Office 脚本和 Office Web 外接程序都侧重于 Web 和协作，但 Office 脚本适合最终用户 (而 Office Web 外接程序面向专业开发人员) 。":::

<span data-ttu-id="40faf-109">Office脚本通过手动按下按钮或作为 Power Automate 中的步骤运行以[](https://flow.microsoft.com/)完成，Office外接程序将继续运行，具体取决于其配置方式。</span><span class="sxs-lookup"><span data-stu-id="40faf-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins continue running depending on how they are configured.</span></span> <span data-ttu-id="40faf-110">例如，可以将加载项Office，即使任务窗格关闭，也可以继续运行。</span><span class="sxs-lookup"><span data-stu-id="40faf-110">For example, you can configure an Office Add-in to continue running even when its task pane is closed.</span></span> <span data-ttu-id="40faf-111">这意味着Office加载项在会话期间保持状态，而Office脚本不会在运行之间保持内部状态。</span><span class="sxs-lookup"><span data-stu-id="40faf-111">This means that Office Add-ins maintain state during a session, whereas Office Scripts don't maintain an internal state between runs.</span></span> <span data-ttu-id="40faf-112">如果您要构建的解决方案需要保持状态，则应该访问 Office[外接程序](/office/dev/add-ins)文档，以了解有关Office外接程序的信息。</span><span class="sxs-lookup"><span data-stu-id="40faf-112">If the solution you are building requires a maintained state, you should visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="40faf-113">本文的其余部分介绍加载项和脚本Office之间的主要Office区别。</span><span class="sxs-lookup"><span data-stu-id="40faf-113">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="40faf-114">平台支持</span><span class="sxs-lookup"><span data-stu-id="40faf-114">Platform Support</span></span>

<span data-ttu-id="40faf-115">Office外接程序是跨平台的。</span><span class="sxs-lookup"><span data-stu-id="40faf-115">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="40faf-116">它们跨桌面Windows、Mac、iOS 和 Web 平台运行，并在每个平台上提供相同的体验。</span><span class="sxs-lookup"><span data-stu-id="40faf-116">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="40faf-117">有关此情况的任何例外情况都记录在单个 API 的文档中。</span><span class="sxs-lookup"><span data-stu-id="40faf-117">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="40faf-118">Office脚本当前仅受 Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="40faf-118">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="40faf-119">所有录制、编辑和运行均在 Web 平台上完成。</span><span class="sxs-lookup"><span data-stu-id="40faf-119">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="40faf-120">API</span><span class="sxs-lookup"><span data-stu-id="40faf-120">APIs</span></span>

<span data-ttu-id="40faf-121">尽管Office加载项Office JavaScript API 和 Office 脚本 API 共享一些功能，但两者是不同的平台。</span><span class="sxs-lookup"><span data-stu-id="40faf-121">While the Office JavaScript APIs for Office Add-ins and the Office Scripts APIs share some functionality, they are different platforms.</span></span> <span data-ttu-id="40faf-122">Office脚本 API 是 JavaScript API 模型的优化Excel子集。</span><span class="sxs-lookup"><span data-stu-id="40faf-122">The Office Scripts APIs are an optimized, synchronous subset of the Excel JavaScript API model.</span></span> <span data-ttu-id="40faf-123">主要区别是范例 `load` / `sync` 与加载项的用法。此外，加载项还提供事件 API 以及 Excel 之外的一组更广泛的功能，称为通用 API。</span><span class="sxs-lookup"><span data-stu-id="40faf-123">The major difference is usage of the `load`/`sync` paradigm with add-ins. Additionally, add-ins offer APIs for events and a broader set of functionality outside of Excel, known as the Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="40faf-124">活动</span><span class="sxs-lookup"><span data-stu-id="40faf-124">Events</span></span>

<span data-ttu-id="40faf-125">Office脚本不支持工作簿级[事件](/office/dev/add-ins/excel/excel-add-ins-events)。</span><span class="sxs-lookup"><span data-stu-id="40faf-125">Office Scripts do not support workbook-level [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="40faf-126">脚本由用户按脚本的 **"运行**"按钮触发，或Power Automate。</span><span class="sxs-lookup"><span data-stu-id="40faf-126">Scripts are either triggered by users pressing the **Run** button for a script or through Power Automate.</span></span> <span data-ttu-id="40faf-127">每个脚本在一个方法中运行 `main` 代码，然后结束。</span><span class="sxs-lookup"><span data-stu-id="40faf-127">Every script runs the code in a single `main` method, then ends.</span></span>

### <a name="common-apis"></a><span data-ttu-id="40faf-128">通用 API</span><span class="sxs-lookup"><span data-stu-id="40faf-128">Common APIs</span></span>

<span data-ttu-id="40faf-129">Office脚本不能使用[通用 API。](/javascript/api/office)</span><span class="sxs-lookup"><span data-stu-id="40faf-129">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="40faf-130">如果你需要身份验证、对话框窗口或其他仅受通用 API 支持的功能，你可能需要创建一个 Office 外接程序，而不是一个 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="40faf-130">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="40faf-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="40faf-131">See also</span></span>

- [<span data-ttu-id="40faf-132">Excel web 版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="40faf-132">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="40faf-133">脚本Office VBA 宏之间的差异</span><span class="sxs-lookup"><span data-stu-id="40faf-133">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="40faf-134">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="40faf-134">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="40faf-135">生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="40faf-135">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
