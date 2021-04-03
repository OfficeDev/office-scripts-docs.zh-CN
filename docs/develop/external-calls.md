---
title: Office 脚本中的外部 API 呼叫支持
description: 在 Office 脚本中执行外部 API 调用的支持和指导。
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 74b8750f609370370759ca4a4a1daa998363ac2e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570309"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="571c6-103">Office 脚本中的外部 API 呼叫支持</span><span class="sxs-lookup"><span data-stu-id="571c6-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="571c6-104">在平台的预览阶段使用外部 API[](https://developer.mozilla.org/docs/Web/API)时，脚本作者不应期望行为一致。</span><span class="sxs-lookup"><span data-stu-id="571c6-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="571c6-105">因此，不要依赖外部 API 实现关键脚本方案。</span><span class="sxs-lookup"><span data-stu-id="571c6-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="571c6-106">对外部 API 的调用只能通过 Excel 应用程序进行，而不是在正常情况下通过 Power Automate [进行](#external-calls-from-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="571c6-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="571c6-107">外部调用可能会导致敏感数据向不需要的终结点公开。</span><span class="sxs-lookup"><span data-stu-id="571c6-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="571c6-108">管理员可以针对此类呼叫建立防火墙保护。</span><span class="sxs-lookup"><span data-stu-id="571c6-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="working-with-fetch"></a><span data-ttu-id="571c6-109">使用 `fetch`</span><span class="sxs-lookup"><span data-stu-id="571c6-109">Working with `fetch`</span></span>

<span data-ttu-id="571c6-110">提取 [API](https://developer.mozilla.org/docs/Web/API/Fetch_API) 从外部服务检索信息。</span><span class="sxs-lookup"><span data-stu-id="571c6-110">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="571c6-111">它是 `async` 一个 API，因此你需要调整 `main` 脚本的签名。</span><span class="sxs-lookup"><span data-stu-id="571c6-111">It is an `async` API, so you will need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="571c6-112">创建 `main` 函数 `async` ，并返回 `Promise<void>` 。</span><span class="sxs-lookup"><span data-stu-id="571c6-112">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="571c6-113">还应确保进行 `await` `fetch` 呼叫和 `json` 检索。</span><span class="sxs-lookup"><span data-stu-id="571c6-113">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="571c6-114">这将确保在脚本结束之前完成这些操作。</span><span class="sxs-lookup"><span data-stu-id="571c6-114">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="571c6-115">以下脚本使用 `fetch` 从给定 URL 中的测试服务器检索 JSON 数据。</span><span class="sxs-lookup"><span data-stu-id="571c6-115">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span>

```TypeScript
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

<span data-ttu-id="571c6-116">[Office 脚本示例方案：来自 NOAA 的图形](../resources/scenarios/noaa-data-fetch.md)水级数据演示了用于从"国家/地区"和"国家/地区管理中心"的"目录和当前"数据库中检索记录的提取命令。</span><span class="sxs-lookup"><span data-stu-id="571c6-116">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="571c6-117">来自 Power Automate 的外部呼叫</span><span class="sxs-lookup"><span data-stu-id="571c6-117">External calls from Power Automate</span></span>

<span data-ttu-id="571c6-118">使用 Power Automate 运行脚本时，任何外部 API 调用都失败。</span><span class="sxs-lookup"><span data-stu-id="571c6-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="571c6-119">这是通过 Excel 客户端和 Power Automate 运行脚本的行为差异。</span><span class="sxs-lookup"><span data-stu-id="571c6-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="571c6-120">在将脚本构建到流中之前，请务必检查脚本中是否包含此类引用。</span><span class="sxs-lookup"><span data-stu-id="571c6-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="571c6-121">通过 Power Automate [Excel Online](/connectors/excelonlinebusiness) 连接器进行的外部调用失败，以帮助制定现有数据丢失防护策略。</span><span class="sxs-lookup"><span data-stu-id="571c6-121">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="571c6-122">但是，通过 Power Automate 运行的脚本在组织外部和组织的防火墙之外执行。</span><span class="sxs-lookup"><span data-stu-id="571c6-122">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="571c6-123">要在此外部环境中对恶意用户进行其他保护，管理员可以控制 Office 脚本的使用。</span><span class="sxs-lookup"><span data-stu-id="571c6-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="571c6-124">管理员可以在 Power Automate 中禁用 Excel Online 连接器，或者通过 Office 脚本管理员控件关闭 Excel 网页版 [Office 脚本](/microsoft-365/admin/manage/manage-office-scripts-settings)。</span><span class="sxs-lookup"><span data-stu-id="571c6-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="571c6-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="571c6-125">See also</span></span>

- [<span data-ttu-id="571c6-126">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="571c6-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="571c6-127">Office 脚本示例方案：GRAPH NOAA 中的水级数据</span><span class="sxs-lookup"><span data-stu-id="571c6-127">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
