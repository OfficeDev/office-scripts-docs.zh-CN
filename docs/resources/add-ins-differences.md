---
title: Office 脚本与 Office 外接程序之间的差异
description: Office 脚本与 Office 外接程序之间的行为和 API 差异。
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700127"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="9c06d-103">Office 脚本与 Office 外接程序之间的差异</span><span class="sxs-lookup"><span data-stu-id="9c06d-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="9c06d-104">Office 外接程序和 Office 脚本具有很多共同之处。</span><span class="sxs-lookup"><span data-stu-id="9c06d-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="9c06d-105">它们都通过 Office JavaScript API 的`Excel`命名空间提供对 Excel 工作簿的自动控制。</span><span class="sxs-lookup"><span data-stu-id="9c06d-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="9c06d-106">但是，Office 脚本的作用范围更有限。</span><span class="sxs-lookup"><span data-stu-id="9c06d-106">However, Office Scripts are more limited in their scope.</span></span>

<span data-ttu-id="9c06d-107">Office 脚本运行到完成时只需要按手动按钮，而 Office 外接程序依赖于用户交互并在工作簿使用时保持不变。</span><span class="sxs-lookup"><span data-stu-id="9c06d-107">Office Scripts run to completion with a manual button press, whereas Office Add-ins rely on user interaction and persist while the workbook is in use.</span></span> <span data-ttu-id="9c06d-108">如果发现您的 Excel 扩展需要超过脚本平台的功能，请访问[Office 外接程序文档](/office/dev/add-ins)以了解有关 Office 外接程序的详细信息。</span><span class="sxs-lookup"><span data-stu-id="9c06d-108">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="9c06d-109">本文的其余部分将介绍 Office 外接程序和 Office 脚本之间的主要差异。</span><span class="sxs-lookup"><span data-stu-id="9c06d-109">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="9c06d-110">平台支持</span><span class="sxs-lookup"><span data-stu-id="9c06d-110">Platform Support</span></span>

<span data-ttu-id="9c06d-111">Office 外接程序是跨平台的。</span><span class="sxs-lookup"><span data-stu-id="9c06d-111">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="9c06d-112">它们在 Windows 桌面、Mac、iOS 和 web 平台上工作，并在每个平台上提供相同的体验。</span><span class="sxs-lookup"><span data-stu-id="9c06d-112">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="9c06d-113">每个 API 的文档中注明了此错误的任何例外。</span><span class="sxs-lookup"><span data-stu-id="9c06d-113">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="9c06d-114">Office 脚本目前仅对 web 上的 Excel 受支持。</span><span class="sxs-lookup"><span data-stu-id="9c06d-114">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="9c06d-115">所有录制、编辑和运行都是在 web 平台上完成的。</span><span class="sxs-lookup"><span data-stu-id="9c06d-115">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="9c06d-116">API</span><span class="sxs-lookup"><span data-stu-id="9c06d-116">APIs</span></span>

<span data-ttu-id="9c06d-117">Office 脚本支持大多数 Excel JavaScript Api，这意味着这两个平台之间存在许多功能重叠。</span><span class="sxs-lookup"><span data-stu-id="9c06d-117">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="9c06d-118">有两个例外：事件和常见 Api。</span><span class="sxs-lookup"><span data-stu-id="9c06d-118">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="9c06d-119">活动</span><span class="sxs-lookup"><span data-stu-id="9c06d-119">Events</span></span>

<span data-ttu-id="9c06d-120">Office 脚本不支持[事件](/office/dev/add-ins/excel/excel-add-ins-events)。</span><span class="sxs-lookup"><span data-stu-id="9c06d-120">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="9c06d-121">每个脚本在一个`main`方法中运行代码，然后结束。</span><span class="sxs-lookup"><span data-stu-id="9c06d-121">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="9c06d-122">触发事件时不会重新激活，因此无法注册事件。</span><span class="sxs-lookup"><span data-stu-id="9c06d-122">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="9c06d-123">通用 API</span><span class="sxs-lookup"><span data-stu-id="9c06d-123">Common APIs</span></span>

<span data-ttu-id="9c06d-124">Office 脚本无法使用[通用 api](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="9c06d-124">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="9c06d-125">如果需要身份验证、对话窗口或其他仅受常见 Api 支持的功能，则您可能需要创建 Office 加载项，而不是 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="9c06d-125">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="9c06d-126">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9c06d-126">See also</span></span>

- [<span data-ttu-id="9c06d-127">Web 上的 Excel 中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="9c06d-127">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="9c06d-128">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="9c06d-128">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="9c06d-129">生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="9c06d-129">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)