---
title: Office 脚本的平台限制和要求
description: 与 Excel 网页 Excel 一同使用时，Office 脚本的资源限制和浏览器支持
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: ef733562fb3caa8261fbbd8382923927a46cb7d4
ms.sourcegitcommit: 5ca286615a11d282e3f80023d22d36a039800eed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/13/2021
ms.locfileid: "51689764"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="46c90-103">Office 脚本的平台限制和要求</span><span class="sxs-lookup"><span data-stu-id="46c90-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="46c90-104">开发 Office 脚本时应注意一些平台限制。</span><span class="sxs-lookup"><span data-stu-id="46c90-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="46c90-105">本文详细介绍了 Excel 网页 Office 脚本的浏览器支持和数据限制。</span><span class="sxs-lookup"><span data-stu-id="46c90-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="46c90-106">浏览器支持</span><span class="sxs-lookup"><span data-stu-id="46c90-106">Browser support</span></span>

<span data-ttu-id="46c90-107">Office 脚本适用于任何支持 [Office 网页的浏览器](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)。</span><span class="sxs-lookup"><span data-stu-id="46c90-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="46c90-108">但是，IE 11 版本 11 Internet Explorer不支持 (JavaScript) 。</span><span class="sxs-lookup"><span data-stu-id="46c90-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="46c90-109">[ES6 或更高版本中引入](https://www.w3schools.com/Js/js_es6.asp)的任何功能将不能与 IE 11 一起使用。</span><span class="sxs-lookup"><span data-stu-id="46c90-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="46c90-110">如果组织成员仍使用该浏览器，请务必在共享脚本时测试该环境中脚本。</span><span class="sxs-lookup"><span data-stu-id="46c90-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="46c90-111">第三方 Cookie</span><span class="sxs-lookup"><span data-stu-id="46c90-111">Third-party cookies</span></span>

<span data-ttu-id="46c90-112">浏览器需要启用第三方 Cookie，以在Excel 网页中显示"自动"选项卡。</span><span class="sxs-lookup"><span data-stu-id="46c90-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="46c90-113">如果未显示选项卡，请检查浏览器设置。</span><span class="sxs-lookup"><span data-stu-id="46c90-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="46c90-114">如果使用的是专用浏览器会话，可能需要每次重新启用此设置。</span><span class="sxs-lookup"><span data-stu-id="46c90-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="46c90-115">某些浏览器将此设置视为"所有 Cookie"，而不是"第三方 Cookie"。</span><span class="sxs-lookup"><span data-stu-id="46c90-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="46c90-116">在热门浏览器中调整 Cookie 设置的说明</span><span class="sxs-lookup"><span data-stu-id="46c90-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="46c90-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="46c90-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="46c90-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="46c90-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="46c90-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="46c90-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="46c90-120">Safari</span><span class="sxs-lookup"><span data-stu-id="46c90-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="46c90-121">数据限制</span><span class="sxs-lookup"><span data-stu-id="46c90-121">Data limits</span></span>

<span data-ttu-id="46c90-122">一次可传输的 Excel 数据量以及可以执行单个 Power Automate 事务数存在限制。</span><span class="sxs-lookup"><span data-stu-id="46c90-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="46c90-123">Excel</span><span class="sxs-lookup"><span data-stu-id="46c90-123">Excel</span></span>

<span data-ttu-id="46c90-124">通过脚本调用工作簿时，Excel 网页具有以下限制：</span><span class="sxs-lookup"><span data-stu-id="46c90-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="46c90-125">请求和响应限制为 **5MB。**</span><span class="sxs-lookup"><span data-stu-id="46c90-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="46c90-126">范围限制为五百 **万个单元格**。</span><span class="sxs-lookup"><span data-stu-id="46c90-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="46c90-127">如果在处理大型数据集时遇到错误，请尝试使用多个较小的范围，而不是较大的区域。</span><span class="sxs-lookup"><span data-stu-id="46c90-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="46c90-128">还可以将 [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) 等 API 定向到特定单元格，而不是大型区域。</span><span class="sxs-lookup"><span data-stu-id="46c90-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="46c90-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="46c90-129">Power Automate</span></span>

<span data-ttu-id="46c90-130">将 Office 脚本与 Power Automate 一同使用时，每个用户每天只能调用 **400 次运行脚本操作**。</span><span class="sxs-lookup"><span data-stu-id="46c90-130">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="46c90-131">此限制在 UTC 时间上午 12：00 重置。</span><span class="sxs-lookup"><span data-stu-id="46c90-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="46c90-132">Power Automate 平台还具有使用限制，可在以下文章中找到这些限制：</span><span class="sxs-lookup"><span data-stu-id="46c90-132">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="46c90-133">Power Automate 中的限制和配置</span><span class="sxs-lookup"><span data-stu-id="46c90-133">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="46c90-134">Excel Online (Business) 连接器的已知问题和限制</span><span class="sxs-lookup"><span data-stu-id="46c90-134">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="46c90-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="46c90-135">See also</span></span>

- [<span data-ttu-id="46c90-136">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="46c90-136">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="46c90-137">消除 Office 脚本的影响</span><span class="sxs-lookup"><span data-stu-id="46c90-137">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="46c90-138">提高 Office 脚本的性能</span><span class="sxs-lookup"><span data-stu-id="46c90-138">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="46c90-139">Excel 网页中的 Office 脚本脚本基础</span><span class="sxs-lookup"><span data-stu-id="46c90-139">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
