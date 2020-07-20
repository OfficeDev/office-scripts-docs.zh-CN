---
title: Office 脚本中的外部 API 调用支持
description: 在 Office 脚本中进行外部 API 调用的支持和指南。
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: ec8281551cbe7c500eee40ec86067e5efbfcfc31
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878723"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="bef28-103">Office 脚本中的外部 API 调用支持</span><span class="sxs-lookup"><span data-stu-id="bef28-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="bef28-104">Office 脚本平台不支持对[外部 api](https://developer.mozilla.org/docs/Web/API)的调用。</span><span class="sxs-lookup"><span data-stu-id="bef28-104">The Office Scripts platform doesn't support calls to [external APIs](https://developer.mozilla.org/docs/Web/API).</span></span> <span data-ttu-id="bef28-105">但是，在适当的情况下，可以运行这些呼叫。</span><span class="sxs-lookup"><span data-stu-id="bef28-105">However, these calls can be run under the right circumstances.</span></span> <span data-ttu-id="bef28-106">外部呼叫只能通过 Excel 客户端进行，而不是[在正常情况下](#external-calls-from-power-automate)的电源自动运行。</span><span class="sxs-lookup"><span data-stu-id="bef28-106">External calls can be only be made through the Excel client, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

<span data-ttu-id="bef28-107">在平台的预览阶段使用外部 Api 时，脚本作者不应预期一致的行为。</span><span class="sxs-lookup"><span data-stu-id="bef28-107">Script authors shouldn't expect consistent behavior when using external APIs during the platform's preview phase.</span></span> <span data-ttu-id="bef28-108">这是因为 JavaScript 运行时如何管理与工作簿的交互。</span><span class="sxs-lookup"><span data-stu-id="bef28-108">This is due how the JavaScript runtime manages interacting with the workbook.</span></span> <span data-ttu-id="bef28-109">脚本可能在 API 调用完成之前结束（或 `Promise` 完全解决）。</span><span class="sxs-lookup"><span data-stu-id="bef28-109">The script may end before the API call completes (or its `Promise` is fully resolved).</span></span> <span data-ttu-id="bef28-110">因此，不要依赖于对关键脚本方案的外部 Api。</span><span class="sxs-lookup"><span data-stu-id="bef28-110">As such, do not rely on external APIs for critical script scenarios.</span></span>

> [!CAUTION]
> <span data-ttu-id="bef28-111">外部调用可能会导致敏感数据暴露给不需要的终结点。</span><span class="sxs-lookup"><span data-stu-id="bef28-111">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="bef28-112">你的管理员可以针对此类呼叫建立防火墙保护。</span><span class="sxs-lookup"><span data-stu-id="bef28-112">Your admin can establish firewall protection against such calls.</span></span>

## <a name="definition-files-for-external-apis"></a><span data-ttu-id="bef28-113">外部 Api 的定义文件</span><span class="sxs-lookup"><span data-stu-id="bef28-113">Definition files for external APIs</span></span>

<span data-ttu-id="bef28-114">外部 Api 的定义文件不包含在 Office 脚本中。</span><span class="sxs-lookup"><span data-stu-id="bef28-114">The definition files for external APIs aren't included with Office Scripts.</span></span> <span data-ttu-id="bef28-115">使用此类 Api 会为缺少的定义生成编译时错误。</span><span class="sxs-lookup"><span data-stu-id="bef28-115">The use of such APIs generates compile-time errors for missing definitions.</span></span> <span data-ttu-id="bef28-116">Api 仍运行（尽管仅在通过 Excel 客户端运行时），如以下脚本所示：</span><span class="sxs-lookup"><span data-stu-id="bef28-116">The APIs still run (though only when run through the Excel client), as shown in the following script:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="bef28-117">来自电源自动执行的外部呼叫</span><span class="sxs-lookup"><span data-stu-id="bef28-117">External calls from Power Automate</span></span>

<span data-ttu-id="bef28-118">使用 Power 自动化运行脚本时，任何外部 API 调用都将失败。</span><span class="sxs-lookup"><span data-stu-id="bef28-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="bef28-119">这是通过 Excel 客户端和 Power 自动化运行脚本之间的行为差异。</span><span class="sxs-lookup"><span data-stu-id="bef28-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="bef28-120">在将脚本生成到流中之前，请务必检查脚本中的这些引用。</span><span class="sxs-lookup"><span data-stu-id="bef28-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="bef28-121">如果外部呼叫[Excel Online 连接器](/connectors/excelonlinebusiness)在 Power 自动化中出现故障，则可以帮助您降低现有数据丢失防护策略。</span><span class="sxs-lookup"><span data-stu-id="bef28-121">The failure of external calls [Excel Online connector](/connectors/excelonlinebusiness) in Power Automate is there to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="bef28-122">但是，在您的组织外和组织的防火墙外部，这些脚本将通过 "电源自动完成" 运行。</span><span class="sxs-lookup"><span data-stu-id="bef28-122">However, the scripts run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="bef28-123">若要进一步保护此外部环境中的恶意用户，管理员可以控制 Office 脚本的使用。</span><span class="sxs-lookup"><span data-stu-id="bef28-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="bef28-124">您的管理员可以通过[Office 脚本管理员控件](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)在 web 上禁用 excel Online 连接器，或在 web 上关闭 Excel 的 office 脚本。</span><span class="sxs-lookup"><span data-stu-id="bef28-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="see-also"></a><span data-ttu-id="bef28-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bef28-125">See also</span></span>

- [<span data-ttu-id="bef28-126">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="bef28-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)