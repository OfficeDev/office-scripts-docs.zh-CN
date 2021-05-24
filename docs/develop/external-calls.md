---
title: Office 脚本中的外部 API 呼叫支持
description: 在脚本中执行外部 API 调用Office指南。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fd6ba0c57bf4cabb2d07421355cacff373f6706c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545080"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="ea70b-103">Office 脚本中的外部 API 呼叫支持</span><span class="sxs-lookup"><span data-stu-id="ea70b-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="ea70b-104">在平台的预览阶段使用外部 API[](https://developer.mozilla.org/docs/Web/API)时，脚本作者不应期望行为一致。</span><span class="sxs-lookup"><span data-stu-id="ea70b-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="ea70b-105">因此，不要依赖外部 API 实现关键脚本方案。</span><span class="sxs-lookup"><span data-stu-id="ea70b-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="ea70b-106">对外部 API 的调用只能通过 Excel 应用程序进行，而在正常情况下Power Automate[调用](#external-calls-from-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="ea70b-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="ea70b-107">外部调用可能会导致敏感数据向不需要的终结点公开。</span><span class="sxs-lookup"><span data-stu-id="ea70b-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="ea70b-108">管理员可以针对此类呼叫建立防火墙保护。</span><span class="sxs-lookup"><span data-stu-id="ea70b-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="configure-your-script-for-external-calls"></a><span data-ttu-id="ea70b-109">为外部调用配置脚本</span><span class="sxs-lookup"><span data-stu-id="ea70b-109">Configure your script for external calls</span></span>

<span data-ttu-id="ea70b-110">外部调用 [是异步](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) 的，需要将脚本标记为 `async` 。</span><span class="sxs-lookup"><span data-stu-id="ea70b-110">External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`.</span></span> <span data-ttu-id="ea70b-111">将 `async` 前缀添加到 函数 `main` ，并返回 `Promise` ，如下所示：</span><span class="sxs-lookup"><span data-stu-id="ea70b-111">Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> <span data-ttu-id="ea70b-112">返回其他信息的脚本可以返回 `Promise` 该类型的 。</span><span class="sxs-lookup"><span data-stu-id="ea70b-112">Scripts that return other information can return a `Promise` of that type.</span></span> <span data-ttu-id="ea70b-113">例如，如果您的脚本需要返回对象 `Employee` ，则返回签名为 `: Promise <Employee>`</span><span class="sxs-lookup"><span data-stu-id="ea70b-113">For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`</span></span>

<span data-ttu-id="ea70b-114">您需要了解外部服务的接口，以调用该服务。</span><span class="sxs-lookup"><span data-stu-id="ea70b-114">You'll need to learn the external service's interfaces to make calls to that service.</span></span> <span data-ttu-id="ea70b-115">如果使用 或 `fetch` [REST API，](https://wikipedia.org/wiki/Representational_state_transfer)则需要确定返回数据的 JSON 结构。</span><span class="sxs-lookup"><span data-stu-id="ea70b-115">If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data.</span></span> <span data-ttu-id="ea70b-116">对于脚本的输入和输出，请考虑使 与所需的 `interface` JSON 结构相匹配。</span><span class="sxs-lookup"><span data-stu-id="ea70b-116">For both input to and output from your script, consider making an `interface` to match the needed JSON structures.</span></span> <span data-ttu-id="ea70b-117">这为脚本提供了更多的类型安全性。</span><span class="sxs-lookup"><span data-stu-id="ea70b-117">This gives the script more type safety.</span></span> <span data-ttu-id="ea70b-118">有关此内容的示例，请参阅[Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md)。</span><span class="sxs-lookup"><span data-stu-id="ea70b-118">You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span></span>

### <a name="limitations-with-external-calls-from-office-scripts"></a><span data-ttu-id="ea70b-119">来自脚本的外部调用Office限制</span><span class="sxs-lookup"><span data-stu-id="ea70b-119">Limitations with external calls from Office Scripts</span></span>

* <span data-ttu-id="ea70b-120">无法登录或使用 OAuth2 类型的身份验证流。</span><span class="sxs-lookup"><span data-stu-id="ea70b-120">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="ea70b-121">所有密钥和凭据必须硬编码 (源文件进行硬编码) 。</span><span class="sxs-lookup"><span data-stu-id="ea70b-121">All keys and credentials have to be hardcoded (or read from another source).</span></span>
* <span data-ttu-id="ea70b-122">没有用于存储 API 凭据和密钥的基础结构。</span><span class="sxs-lookup"><span data-stu-id="ea70b-122">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="ea70b-123">这必须由用户管理。</span><span class="sxs-lookup"><span data-stu-id="ea70b-123">This will have to be managed by the user.</span></span>
* <span data-ttu-id="ea70b-124">不支持文档 `localStorage` Cookie、和 `sessionStorage` 对象。</span><span class="sxs-lookup"><span data-stu-id="ea70b-124">Document cookies, `localStorage`, and `sessionStorage` objects are not supported.</span></span> 
* <span data-ttu-id="ea70b-125">外部调用可能会导致向不需要的终结点公开敏感数据，或导致外部数据进入内部工作簿。</span><span class="sxs-lookup"><span data-stu-id="ea70b-125">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="ea70b-126">管理员可以针对此类呼叫建立防火墙保护。</span><span class="sxs-lookup"><span data-stu-id="ea70b-126">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="ea70b-127">在依赖外部调用之前，请务必检查本地策略。</span><span class="sxs-lookup"><span data-stu-id="ea70b-127">Be sure to check with local policies prior to relying on external calls.</span></span>
* <span data-ttu-id="ea70b-128">请务必在依赖关系之前检查数据吞吐量。</span><span class="sxs-lookup"><span data-stu-id="ea70b-128">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="ea70b-129">例如，下拉整个外部数据集可能不是最佳选择，而应该使用分页获取区块中的数据。</span><span class="sxs-lookup"><span data-stu-id="ea70b-129">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="retrieve-information-with-fetch"></a><span data-ttu-id="ea70b-130">使用 检索信息 `fetch`</span><span class="sxs-lookup"><span data-stu-id="ea70b-130">Retrieve information with `fetch`</span></span>

<span data-ttu-id="ea70b-131">提取 [API](https://developer.mozilla.org/docs/Web/API/Fetch_API) 从外部服务检索信息。</span><span class="sxs-lookup"><span data-stu-id="ea70b-131">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="ea70b-132">它是一 `async` 个 API，因此你需要调整 `main` 脚本的签名。</span><span class="sxs-lookup"><span data-stu-id="ea70b-132">It is an `async` API, so you need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="ea70b-133">创建 `main` 函数 `async` ，并返回 `Promise<void>` 。</span><span class="sxs-lookup"><span data-stu-id="ea70b-133">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="ea70b-134">还应确保进行 `await` `fetch` 呼叫和 `json` 检索。</span><span class="sxs-lookup"><span data-stu-id="ea70b-134">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="ea70b-135">这将确保在脚本结束之前完成这些操作。</span><span class="sxs-lookup"><span data-stu-id="ea70b-135">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="ea70b-136">检索到的任何 JSON 数据 `fetch` 都必须与脚本中定义的接口匹配。</span><span class="sxs-lookup"><span data-stu-id="ea70b-136">Any JSON data retrieved by `fetch` must match an interface defined in the script.</span></span> <span data-ttu-id="ea70b-137">返回的值必须分配给特定类型，因为Office[脚本不支持 `any` 类型](typescript-restrictions.md#no-any-type-in-office-scripts)。</span><span class="sxs-lookup"><span data-stu-id="ea70b-137">The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts).</span></span> <span data-ttu-id="ea70b-138">应参考服务文档，以查看返回的属性的名称和类型。</span><span class="sxs-lookup"><span data-stu-id="ea70b-138">You should refer to the documentation for your service to see what the names and types of the returned properties are.</span></span> <span data-ttu-id="ea70b-139">然后，将匹配的接口添加到脚本。</span><span class="sxs-lookup"><span data-stu-id="ea70b-139">Then, add the matching interface or interfaces to your script.</span></span>

<span data-ttu-id="ea70b-140">以下脚本使用 `fetch` 从给定 URL 中的测试服务器检索 JSON 数据。</span><span class="sxs-lookup"><span data-stu-id="ea70b-140">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span> <span data-ttu-id="ea70b-141">请注意 `JSONData` 用于将数据存储为匹配类型的接口。</span><span class="sxs-lookup"><span data-stu-id="ea70b-141">Note the `JSONData` interface to store the data as a matching type.</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### <a name="other-fetch-samples"></a><span data-ttu-id="ea70b-142">其他 `fetch` 示例</span><span class="sxs-lookup"><span data-stu-id="ea70b-142">Other `fetch` samples</span></span>

* <span data-ttu-id="ea70b-143">Use [external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md)示例演示如何获取有关用户的 GitHub 存储库的基本信息。</span><span class="sxs-lookup"><span data-stu-id="ea70b-143">The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.</span></span>
* <span data-ttu-id="ea70b-144">the [Office Scripts sample scenario： Graph water-level](../resources/scenarios/noaa-data-fetch.md) data from NOAA 演示了提取命令，用于从美国国家/地区和省政府管理署的"目录和当前"数据库中检索记录。</span><span class="sxs-lookup"><span data-stu-id="ea70b-144">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="ea70b-145">外部呼叫Power Automate</span><span class="sxs-lookup"><span data-stu-id="ea70b-145">External calls from Power Automate</span></span>

<span data-ttu-id="ea70b-146">在使用脚本运行时，任何外部 API 调用Power Automate。</span><span class="sxs-lookup"><span data-stu-id="ea70b-146">Any external API call fails when a script is run with Power Automate.</span></span> <span data-ttu-id="ea70b-147">这是通过应用程序运行脚本和Excel脚本的行为Power Automate。</span><span class="sxs-lookup"><span data-stu-id="ea70b-147">This is a behavioral difference between running a script through the Excel application and through Power Automate.</span></span> <span data-ttu-id="ea70b-148">在将脚本构建到流中之前，请务必检查脚本中是否包含此类引用。</span><span class="sxs-lookup"><span data-stu-id="ea70b-148">Be sure to check your scripts for such references before building them into a flow.</span></span>

<span data-ttu-id="ea70b-149">你必须将 HTTP 与 [Azure AD](/connectors/webcontents/) 或其他等效操作一同使用，以从外部服务提取数据或将其推送到外部服务。</span><span class="sxs-lookup"><span data-stu-id="ea70b-149">You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.</span></span>

> [!WARNING]
> <span data-ttu-id="ea70b-150">通过 Power Automate [Excel Online](/connectors/excelonlinebusiness)连接器进行的外部呼叫失败，以帮助制定现有数据丢失防护策略。</span><span class="sxs-lookup"><span data-stu-id="ea70b-150">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="ea70b-151">但是，通过 Power Automate运行的脚本在组织外部和组织的防火墙之外执行。</span><span class="sxs-lookup"><span data-stu-id="ea70b-151">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="ea70b-152">对于此外部环境中恶意用户的额外保护，管理员可以控制对脚本Office的使用。</span><span class="sxs-lookup"><span data-stu-id="ea70b-152">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="ea70b-153">管理员可以在 Power Automate 中禁用 Excel Online 连接器，或Office脚本Excel web 版禁用 Office Scripts for [Office Scripts。](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="ea70b-153">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="ea70b-154">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ea70b-154">See also</span></span>

* [<span data-ttu-id="ea70b-155">在 Office 脚本中使用内置的 JavaScript 对象</span><span class="sxs-lookup"><span data-stu-id="ea70b-155">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
* [<span data-ttu-id="ea70b-156">在 Office 脚本中使用外部提取呼叫</span><span class="sxs-lookup"><span data-stu-id="ea70b-156">Use external fetch calls in Office Scripts</span></span>](../resources/samples/external-fetch-calls.md)
* [<span data-ttu-id="ea70b-157">Office脚本示例方案：Graph NOAA 中的水级数据</span><span class="sxs-lookup"><span data-stu-id="ea70b-157">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
