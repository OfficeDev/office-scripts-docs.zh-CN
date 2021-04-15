---
title: Office 脚本和 VBA 宏之间的差异
description: Office 脚本和 Excel VBA 宏之间的行为和 API 差异。
ms.date: 12/14/2020
localization_priority: Normal
ms.openlocfilehash: a56409a5de3eb07876faa88bfbfe78eeca59f70f
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755019"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a><span data-ttu-id="d859e-103">Office 脚本和 VBA 宏之间的差异</span><span class="sxs-lookup"><span data-stu-id="d859e-103">Differences between Office Scripts and VBA macros</span></span>

<span data-ttu-id="d859e-104">Office 脚本和 VBA 宏有很多共同之处。</span><span class="sxs-lookup"><span data-stu-id="d859e-104">Office Scripts and VBA macros have a lot in common.</span></span> <span data-ttu-id="d859e-105">它们都允许用户通过易于使用的操作录制器自动处理解决方案，并允许编辑这些录制。</span><span class="sxs-lookup"><span data-stu-id="d859e-105">They both allow users to automate solutions through an easy-to-use action recorder and allow edits of those recordings.</span></span> <span data-ttu-id="d859e-106">这两个框架旨在让不将自己认为是程序员的人在 Excel 中创建小型程序。</span><span class="sxs-lookup"><span data-stu-id="d859e-106">Both frameworks are designed to empower people who may not consider themselves programmers to create small programs in Excel.</span></span>
<span data-ttu-id="d859e-107">基本区别在于，VBA 宏是为桌面解决方案开发的，而 Office 脚本的设计以跨平台支持和安全性作为指导原则。</span><span class="sxs-lookup"><span data-stu-id="d859e-107">The fundamental difference is that VBA macros are developed for desktop solutions and Office Scripts are designed with cross-platform support and security as the guiding principles.</span></span> <span data-ttu-id="d859e-108">目前，Office 脚本仅在 Excel 网页中受支持。</span><span class="sxs-lookup"><span data-stu-id="d859e-108">Currently, Office Scripts are only supported in Excel on the web.</span></span>

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="一个四象限图表，显示不同 Office 扩展性解决方案的重点区域。Office 脚本和 VBA 宏旨在帮助最终用户创建解决方案，但 Office 脚本是为 Web 和协作 (而 VBA 用于桌面) 。":::

<span data-ttu-id="d859e-110">本文介绍 VBA 宏与常规 (和 Office 脚本) VBA 之间的主要区别。</span><span class="sxs-lookup"><span data-stu-id="d859e-110">This article describes the main differences between VBA macros (as well as VBA in general) and Office Scripts.</span></span> <span data-ttu-id="d859e-111">由于 Office 脚本仅适用于 Excel，因此这是此处讨论的唯一主机。</span><span class="sxs-lookup"><span data-stu-id="d859e-111">Since Office Scripts are only available for Excel, that is the only host being discussed here.</span></span>

## <a name="platform-and-ecosystem"></a><span data-ttu-id="d859e-112">平台和生态系统</span><span class="sxs-lookup"><span data-stu-id="d859e-112">Platform and ecosystem</span></span>

<span data-ttu-id="d859e-113">VBA 专为桌面设计，Office 脚本专为 Web 设计。</span><span class="sxs-lookup"><span data-stu-id="d859e-113">VBA is designed for the desktop and Office Scripts are designed for the web.</span></span> <span data-ttu-id="d859e-114">VBA 可以与用户的桌面进行交互，以使用类似技术（如 COM 和 OLE）进行连接。</span><span class="sxs-lookup"><span data-stu-id="d859e-114">VBA can interact with a user's desktop to connect with similar technologies, such as COM and OLE.</span></span> <span data-ttu-id="d859e-115">但是，VBA 无法方便地调用 Internet。</span><span class="sxs-lookup"><span data-stu-id="d859e-115">However, VBA has no convenient way to call out to the internet.</span></span>

<span data-ttu-id="d859e-116">Office 脚本使用 JavaScript 的通用运行时。</span><span class="sxs-lookup"><span data-stu-id="d859e-116">Office Scripts use a universal runtime for JavaScript.</span></span> <span data-ttu-id="d859e-117">这将提供一致的行为和辅助功能，而不考虑用于运行脚本的机器。</span><span class="sxs-lookup"><span data-stu-id="d859e-117">This gives consistent behavior and accessibility, regardless of the machine being used to run the script.</span></span> <span data-ttu-id="d859e-118">他们还可以调用其他 Web 服务。</span><span class="sxs-lookup"><span data-stu-id="d859e-118">They can also make calls to other web services.</span></span>

## <a name="security"></a><span data-ttu-id="d859e-119">安全性</span><span class="sxs-lookup"><span data-stu-id="d859e-119">Security</span></span>

<span data-ttu-id="d859e-120">VBA 宏具有与 Excel 相同的安全许可。</span><span class="sxs-lookup"><span data-stu-id="d859e-120">VBA macros have the same security clearance as Excel.</span></span> <span data-ttu-id="d859e-121">这样，他们可以访问你的桌面。</span><span class="sxs-lookup"><span data-stu-id="d859e-121">This gives them full access to your desktop.</span></span> <span data-ttu-id="d859e-122">Office 脚本只能访问工作簿，而无法访问托管工作簿的机器。</span><span class="sxs-lookup"><span data-stu-id="d859e-122">Office Scripts only have access to the workbook, not the machine hosting the workbook.</span></span> <span data-ttu-id="d859e-123">此外，无法与脚本共享 JavaScript 身份验证令牌。</span><span class="sxs-lookup"><span data-stu-id="d859e-123">Additionally, no JavaScript authentication tokens can be shared with scripts.</span></span> <span data-ttu-id="d859e-124">这意味着脚本既不具有已登录用户的令牌，也没有用于登录到外部服务的任何 API 功能，因此它们无法使用现有令牌代表用户进行外部调用。</span><span class="sxs-lookup"><span data-stu-id="d859e-124">This means the script has neither the tokens of the signed-in user nor are there any API capabilities for signing in to an external service, so they are unable to use existing tokens to make external calls on behalf of the user.</span></span>

<span data-ttu-id="d859e-125">管理员有三个 VBA 宏选项：允许租户上的所有宏、不允许在租户上运行宏或只允许使用签名证书的宏。</span><span class="sxs-lookup"><span data-stu-id="d859e-125">Admins have three options for VBA macros: allow all macros on the tenant, allow no macros on the tenant, or allow only macros with signed certificates.</span></span> <span data-ttu-id="d859e-126">这种缺少粒度会使隔离单个错误参与者变得困难。</span><span class="sxs-lookup"><span data-stu-id="d859e-126">This lack of granularity makes it hard to isolate a single bad actor.</span></span> <span data-ttu-id="d859e-127">目前，租户的 Office 脚本处于打开或关闭状态。</span><span class="sxs-lookup"><span data-stu-id="d859e-127">Currently, Office Scripts are either on or off for a tenant.</span></span> <span data-ttu-id="d859e-128">但是，我们正在努力使管理员能够更加控制单个脚本和脚本创建者。</span><span class="sxs-lookup"><span data-stu-id="d859e-128">However, we are working to give admins more control over individual scripts and script creators.</span></span>

## <a name="coverage"></a><span data-ttu-id="d859e-129">覆盖范围</span><span class="sxs-lookup"><span data-stu-id="d859e-129">Coverage</span></span>

<span data-ttu-id="d859e-130">目前，VBA 更全面涵盖 Excel 功能，尤其是桌面客户端上提供的功能。</span><span class="sxs-lookup"><span data-stu-id="d859e-130">Currently, VBA offers a more complete coverage of Excel features, particularly those available on the desktop client.</span></span> <span data-ttu-id="d859e-131">Office 脚本几乎涵盖 Excel 网页应用的所有方案。</span><span class="sxs-lookup"><span data-stu-id="d859e-131">Office Scripts cover nearly all of the scenarios for Excel on the web.</span></span> <span data-ttu-id="d859e-132">此外，随着新功能在 Web 上首次推出，Office 脚本将同时支持操作录制器和 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="d859e-132">Additionally, as new features debut on the web, Office Scripts will support them for both the Action Recorder and JavaScript APIs.</span></span>

<span data-ttu-id="d859e-133">Office 脚本不支持 Excel 级 [事件](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)。</span><span class="sxs-lookup"><span data-stu-id="d859e-133">Office Scripts don't support Excel-level [events](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects).</span></span> <span data-ttu-id="d859e-134">脚本仅在用户手动启动或 Power Automate 流调用脚本时运行。</span><span class="sxs-lookup"><span data-stu-id="d859e-134">Scripts are only run when a user manually starts them or when a Power Automate flow calls the script.</span></span>

## <a name="power-automate"></a><span data-ttu-id="d859e-135">Power Automate</span><span class="sxs-lookup"><span data-stu-id="d859e-135">Power Automate</span></span>

<span data-ttu-id="d859e-136">Office 脚本可以通过 Power Automate 运行。</span><span class="sxs-lookup"><span data-stu-id="d859e-136">Office Scripts can be run through Power Automate.</span></span> <span data-ttu-id="d859e-137">工作簿可以通过计划流或事件驱动的流进行更新，让你无需打开 Excel 即可自动执行工作流。</span><span class="sxs-lookup"><span data-stu-id="d859e-137">Your workbook can be updated through scheduled or event-driven flows, letting you automate workflows without even opening Excel.</span></span> <span data-ttu-id="d859e-138">这意味着，只要工作簿存储在 OneDrive (且可供 Power Automate) 访问，无论您和组织是使用 Excel 桌面版、Mac 还是 Web 客户端，流都可以运行脚本。</span><span class="sxs-lookup"><span data-stu-id="d859e-138">This means that as long as your workbook is stored in OneDrive (and accessible to Power Automate), a flow can run your scripts regardless of whether you and your organization use Excel's desktop, Mac, or web client.</span></span>

<span data-ttu-id="d859e-139">VBA 没有 Power Automate 连接器。</span><span class="sxs-lookup"><span data-stu-id="d859e-139">VBA doesn't have a Power Automate connector.</span></span> <span data-ttu-id="d859e-140">所有支持的 VBA 方案都涉及用户参与宏的执行。</span><span class="sxs-lookup"><span data-stu-id="d859e-140">All supported VBA scenarios involved a user attending to the macro's execution.</span></span>

## <a name="see-also"></a><span data-ttu-id="d859e-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d859e-141">See also</span></span>

- [<span data-ttu-id="d859e-142">Excel web 版中的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="d859e-142">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="d859e-143">Office 脚本与 Office 加载项之间的差异</span><span class="sxs-lookup"><span data-stu-id="d859e-143">Differences between Office Scripts and Office Add-ins</span></span>](add-ins-differences.md)
- [<span data-ttu-id="d859e-144">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="d859e-144">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="d859e-145">Excel VBA 参考</span><span class="sxs-lookup"><span data-stu-id="d859e-145">Excel VBA reference</span></span>](/office/vba/api/overview/excel)
