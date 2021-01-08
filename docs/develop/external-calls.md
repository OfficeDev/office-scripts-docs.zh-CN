---
title: Office 脚本中的外部 API 呼叫支持
description: 支持和指导在 Office 脚本中调用外部 API。
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 1091031bc2e12f3e1e79b177c69874ee4ce61dd8
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784142"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="b3c81-103">Office 脚本中的外部 API 呼叫支持</span><span class="sxs-lookup"><span data-stu-id="b3c81-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="b3c81-104">脚本作者不应期望在平台预览阶段使用外部 [API](https://developer.mozilla.org/docs/Web/API) 时出现一致的行为。</span><span class="sxs-lookup"><span data-stu-id="b3c81-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="b3c81-105">因此，不要依赖外部 API 实现关键脚本方案。</span><span class="sxs-lookup"><span data-stu-id="b3c81-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="b3c81-106">通常，只能通过 Excel 应用程序（而不是 Power Automate）调用[外部 API。](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="b3c81-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="b3c81-107">外部调用可能会导致敏感数据向不需要的终结点公开。</span><span class="sxs-lookup"><span data-stu-id="b3c81-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="b3c81-108">管理员可以针对此类呼叫建立防火墙保护。</span><span class="sxs-lookup"><span data-stu-id="b3c81-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="working-with-fetch"></a><span data-ttu-id="b3c81-109">使用 `fetch`</span><span class="sxs-lookup"><span data-stu-id="b3c81-109">Working with `fetch`</span></span>

<span data-ttu-id="b3c81-110">提取 [API](https://developer.mozilla.org/docs/Web/API/Fetch_API) 从外部服务检索信息。</span><span class="sxs-lookup"><span data-stu-id="b3c81-110">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="b3c81-111">它是一 `async` 个 API，因此你需要调整 `main` 脚本的签名。</span><span class="sxs-lookup"><span data-stu-id="b3c81-111">It is an `async` API, so you will need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="b3c81-112">创建 `main` 函数 `async` ，并返回一个 `Promise<void>` 。</span><span class="sxs-lookup"><span data-stu-id="b3c81-112">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="b3c81-113">还应确保调用 `await` `fetch` 和 `json` 检索。</span><span class="sxs-lookup"><span data-stu-id="b3c81-113">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="b3c81-114">这将确保在脚本结束之前完成这些操作。</span><span class="sxs-lookup"><span data-stu-id="b3c81-114">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="b3c81-115">以下脚本用于 `fetch` 从给定 URL 中的测试服务器检索 JSON 数据。</span><span class="sxs-lookup"><span data-stu-id="b3c81-115">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

<span data-ttu-id="b3c81-116">[Office 脚本示例方案：来自 NOAA 的 Graph](../resources/scenarios/noaa-data-fetch.md)水级数据演示用于从国家远洋和保存管理局的"三项工程"和"当前"数据库检索记录的提取命令。</span><span class="sxs-lookup"><span data-stu-id="b3c81-116">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="b3c81-117">来自 Power Automate 的外部调用</span><span class="sxs-lookup"><span data-stu-id="b3c81-117">External calls from Power Automate</span></span>

<span data-ttu-id="b3c81-118">当使用 Power Automate 运行脚本时，任何外部 API 调用都失败。</span><span class="sxs-lookup"><span data-stu-id="b3c81-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="b3c81-119">这是通过 Excel 客户端和 Power Automate 运行脚本的行为差异。</span><span class="sxs-lookup"><span data-stu-id="b3c81-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="b3c81-120">在将脚本构建到流中之前，请务必检查脚本中是否包含此类引用。</span><span class="sxs-lookup"><span data-stu-id="b3c81-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="b3c81-121">通过 Power Automate [Excel Online](/connectors/excelonlinebusiness) 连接器进行的外部调用失败，以帮助构建现有数据丢失防护策略。</span><span class="sxs-lookup"><span data-stu-id="b3c81-121">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="b3c81-122">但是，通过 Power Automate 运行的脚本在组织外部和组织的防火墙之外执行。</span><span class="sxs-lookup"><span data-stu-id="b3c81-122">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="b3c81-123">对于此外部环境中恶意用户的额外保护，管理员可控制 Office 脚本的使用。</span><span class="sxs-lookup"><span data-stu-id="b3c81-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="b3c81-124">管理员可以在 Power Automate 中禁用 Excel Online 连接器，或者通过 Office 脚本管理员控件关闭 Excel 网页版 [Office 脚本](/microsoft-365/admin/manage/manage-office-scripts-settings)。</span><span class="sxs-lookup"><span data-stu-id="b3c81-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="b3c81-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b3c81-125">See also</span></span>

- [<span data-ttu-id="b3c81-126">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="b3c81-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="b3c81-127">Office 脚本示例方案：绘制 NOAA 中的水级数据</span><span class="sxs-lookup"><span data-stu-id="b3c81-127">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
