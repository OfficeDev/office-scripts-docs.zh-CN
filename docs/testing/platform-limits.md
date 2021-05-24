---
title: Office 脚本的平台限制和要求
description: 与脚本一Office脚本的资源限制和浏览器Excel web 版
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545579"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="5816c-103">Office 脚本的平台限制和要求</span><span class="sxs-lookup"><span data-stu-id="5816c-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="5816c-104">开发脚本时应注意一些平台Office限制。</span><span class="sxs-lookup"><span data-stu-id="5816c-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="5816c-105">本文详细介绍了 Office Scripts for Excel web 版 的浏览器支持和Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="5816c-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="5816c-106">浏览器支持</span><span class="sxs-lookup"><span data-stu-id="5816c-106">Browser support</span></span>

<span data-ttu-id="5816c-107">Office脚本在任何支持 Web Office[的浏览器中工作](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)。</span><span class="sxs-lookup"><span data-stu-id="5816c-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="5816c-108">但是，IE 11 版本 11 Internet Explorer不支持 (JavaScript) 。</span><span class="sxs-lookup"><span data-stu-id="5816c-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="5816c-109">[ES6 或更高版本中引入](https://www.w3schools.com/Js/js_es6.asp)的任何功能将不能与 IE 11 一起使用。</span><span class="sxs-lookup"><span data-stu-id="5816c-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="5816c-110">如果组织成员仍使用该浏览器，请务必在共享脚本时测试该环境中脚本。</span><span class="sxs-lookup"><span data-stu-id="5816c-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="5816c-111">第三方 Cookie</span><span class="sxs-lookup"><span data-stu-id="5816c-111">Third-party cookies</span></span>

<span data-ttu-id="5816c-112">浏览器需要启用第三方 Cookie，以在浏览器中显示"**自动Excel web 版。**</span><span class="sxs-lookup"><span data-stu-id="5816c-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="5816c-113">如果未显示选项卡，请检查浏览器设置。</span><span class="sxs-lookup"><span data-stu-id="5816c-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="5816c-114">如果使用的是专用浏览器会话，可能需要每次重新启用此设置。</span><span class="sxs-lookup"><span data-stu-id="5816c-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="5816c-115">某些浏览器将此设置视为"所有 Cookie"，而不是"第三方 Cookie"。</span><span class="sxs-lookup"><span data-stu-id="5816c-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="5816c-116">在热门浏览器中调整 Cookie 设置的说明</span><span class="sxs-lookup"><span data-stu-id="5816c-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="5816c-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="5816c-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="5816c-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="5816c-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="5816c-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="5816c-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="5816c-120">Safari</span><span class="sxs-lookup"><span data-stu-id="5816c-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="5816c-121">数据限制</span><span class="sxs-lookup"><span data-stu-id="5816c-121">Data limits</span></span>

<span data-ttu-id="5816c-122">对于一次可Excel的数据量以及可以执行单个数据传输Power Automate存在限制。</span><span class="sxs-lookup"><span data-stu-id="5816c-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="5816c-123">Excel</span><span class="sxs-lookup"><span data-stu-id="5816c-123">Excel</span></span>

<span data-ttu-id="5816c-124">Excel通过脚本调用工作簿时，Web 应用程序具有以下限制：</span><span class="sxs-lookup"><span data-stu-id="5816c-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="5816c-125">请求和响应限制为 **5MB。**</span><span class="sxs-lookup"><span data-stu-id="5816c-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="5816c-126">范围限制为五百 **万个单元格**。</span><span class="sxs-lookup"><span data-stu-id="5816c-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="5816c-127">如果在处理大型数据集时遇到错误，请尝试使用多个较小的范围，而不是较大的区域。</span><span class="sxs-lookup"><span data-stu-id="5816c-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="5816c-128">有关示例，请参阅编写 [大型数据集](../resources/samples/write-large-dataset.md) 示例。</span><span class="sxs-lookup"><span data-stu-id="5816c-128">For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span></span> <span data-ttu-id="5816c-129">您还可以使用 [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) 等 API 来定位特定单元格，而不是大型区域。</span><span class="sxs-lookup"><span data-stu-id="5816c-129">You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="5816c-130">Power Automate</span><span class="sxs-lookup"><span data-stu-id="5816c-130">Power Automate</span></span>

<span data-ttu-id="5816c-131">在将Office脚本与Power Automate时，每个用户每天只能调用 **400 次运行脚本操作**。</span><span class="sxs-lookup"><span data-stu-id="5816c-131">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="5816c-132">此限制在 UTC 时间上午 12：00 重置。</span><span class="sxs-lookup"><span data-stu-id="5816c-132">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="5816c-133">the Power Automate platform also has usage limitations， which can be found in the following articles：</span><span class="sxs-lookup"><span data-stu-id="5816c-133">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="5816c-134">中的限制和Power Automate</span><span class="sxs-lookup"><span data-stu-id="5816c-134">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="5816c-135">Excel Online (Business) 连接器的已知问题和限制</span><span class="sxs-lookup"><span data-stu-id="5816c-135">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="5816c-136">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5816c-136">See also</span></span>

- [<span data-ttu-id="5816c-137">脚本Office疑难解答</span><span class="sxs-lookup"><span data-stu-id="5816c-137">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="5816c-138">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="5816c-138">Undo the effects of Office Scripts</span></span>](undo.md)
- [<span data-ttu-id="5816c-139">提高脚本Office性能</span><span class="sxs-lookup"><span data-stu-id="5816c-139">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="5816c-140">Office脚本的脚本Excel web 版</span><span class="sxs-lookup"><span data-stu-id="5816c-140">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
