---
title: 在 Office 脚本中执行外部 API 调用
description: 了解如何在 Office 脚本中执行外部 API 调用。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: d0abfa0bb1adedc7535059ed359b8053d9f1c84d
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571243"
---
# <a name="external-api-calls-from-office-scripts"></a><span data-ttu-id="de3c1-103">Office 脚本中的外部 API 调用</span><span class="sxs-lookup"><span data-stu-id="de3c1-103">External API calls from Office Scripts</span></span>

<span data-ttu-id="de3c1-104">Office 脚本允许 [有限的外部 API 调用支持](../../develop/external-calls.md)。</span><span class="sxs-lookup"><span data-stu-id="de3c1-104">Office Scripts allows [limited external API call support](../../develop/external-calls.md).</span></span>

> [!IMPORTANT]
>
> * <span data-ttu-id="de3c1-105">无法登录或使用 OAuth2 类型的身份验证流。</span><span class="sxs-lookup"><span data-stu-id="de3c1-105">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="de3c1-106">所有密钥和凭据必须硬编码 (源文件进行硬编码) 。</span><span class="sxs-lookup"><span data-stu-id="de3c1-106">All keys and credentials have to be hardcoded (or read from another source).</span></span>
> * <span data-ttu-id="de3c1-107">没有用于存储 API 凭据和密钥的基础结构。</span><span class="sxs-lookup"><span data-stu-id="de3c1-107">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="de3c1-108">这必须由用户管理。</span><span class="sxs-lookup"><span data-stu-id="de3c1-108">This will have to be managed by the user.</span></span>
> * <span data-ttu-id="de3c1-109">外部调用可能会导致向不需要的终结点公开敏感数据，或导致外部数据进入内部工作簿。</span><span class="sxs-lookup"><span data-stu-id="de3c1-109">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="de3c1-110">管理员可以针对此类呼叫建立防火墙保护。</span><span class="sxs-lookup"><span data-stu-id="de3c1-110">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="de3c1-111">在依赖外部调用之前，请务必检查本地策略。</span><span class="sxs-lookup"><span data-stu-id="de3c1-111">Be sure to check with local policies prior to relying on external calls.</span></span>
> * <span data-ttu-id="de3c1-112">如果脚本使用 API 调用，则它在 Power Automate 方案中将无法正常工作。</span><span class="sxs-lookup"><span data-stu-id="de3c1-112">If a script uses an API call, it will not function in a Power Automate scenario.</span></span> <span data-ttu-id="de3c1-113">您必须使用 Power Automate 的 HTTP 操作或等效操作从外部服务提取数据或将其推送到外部服务。</span><span class="sxs-lookup"><span data-stu-id="de3c1-113">You'll have to use Power Automate's HTTP action or equivalent actions to pull data from or push it to an external service.</span></span>
> * <span data-ttu-id="de3c1-114">外部 API 调用涉及异步 API 语法，并且需要稍微高级了解异步通信的工作方式。</span><span class="sxs-lookup"><span data-stu-id="de3c1-114">An external API call involves asynchronous API syntax and requires slightly advanced knowledge of the way async communication works.</span></span>
> * <span data-ttu-id="de3c1-115">请务必在依赖关系之前检查数据吞吐量。</span><span class="sxs-lookup"><span data-stu-id="de3c1-115">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="de3c1-116">例如，下拉整个外部数据集可能不是最佳选择，而应该使用分页获取区块中的数据。</span><span class="sxs-lookup"><span data-stu-id="de3c1-116">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="useful-knowledge-and-resources"></a><span data-ttu-id="de3c1-117">有用的知识和资源</span><span class="sxs-lookup"><span data-stu-id="de3c1-117">Useful knowledge and resources</span></span>

* <span data-ttu-id="de3c1-118">[REST API：](https://en.wikipedia.org/wiki/Representational_state_transfer)最有可能使用 API 调用的方式。</span><span class="sxs-lookup"><span data-stu-id="de3c1-118">[REST API](https://en.wikipedia.org/wiki/Representational_state_transfer): Most likely way you'll use the API call.</span></span>
* <span data-ttu-id="de3c1-119">[ `async` ：了解工作原理 `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)。</span><span class="sxs-lookup"><span data-stu-id="de3c1-119">[`async` `await`](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await): Understand how this works.</span></span>
* <span data-ttu-id="de3c1-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch)：了解工作原理。</span><span class="sxs-lookup"><span data-stu-id="de3c1-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Understand how this works.</span></span>

## <a name="steps"></a><span data-ttu-id="de3c1-121">步骤</span><span class="sxs-lookup"><span data-stu-id="de3c1-121">Steps</span></span>

1. <span data-ttu-id="de3c1-122">通过添加 `main` 前缀将函数标记为异步 `async` 函数。</span><span class="sxs-lookup"><span data-stu-id="de3c1-122">Mark your `main` function as an asynchronous function by adding `async` prefix.</span></span> <span data-ttu-id="de3c1-123">例如，`async function main(workbook: ExcelScript.Workbook)`。</span><span class="sxs-lookup"><span data-stu-id="de3c1-123">For example, `async function main(workbook: ExcelScript.Workbook)`.</span></span>
1. <span data-ttu-id="de3c1-124">你进行哪种类型的 API 调用？</span><span class="sxs-lookup"><span data-stu-id="de3c1-124">Which type of API call are you making?</span></span> <span data-ttu-id="de3c1-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span><span class="sxs-lookup"><span data-stu-id="de3c1-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span></span> <span data-ttu-id="de3c1-126">有关详细信息，请参阅 REST API 材料。</span><span class="sxs-lookup"><span data-stu-id="de3c1-126">Refer to REST API material for details.</span></span>
1. <span data-ttu-id="de3c1-127">获取服务 API 终结点、身份验证要求、标头等。</span><span class="sxs-lookup"><span data-stu-id="de3c1-127">Obtain the service API endpoint, authentication requirements, headers, etc.</span></span>
1. <span data-ttu-id="de3c1-128">定义输入或输出 `interface` 以帮助完成代码和开发时间验证。</span><span class="sxs-lookup"><span data-stu-id="de3c1-128">Define the input or output `interface` to help with code completion and development time verification.</span></span> <span data-ttu-id="de3c1-129">有关详细信息 [，](#training-video-how-to-make-external-api-calls) 请参阅视频。</span><span class="sxs-lookup"><span data-stu-id="de3c1-129">See [video](#training-video-how-to-make-external-api-calls) for details.</span></span>
1. <span data-ttu-id="de3c1-130">代码、测试、优化。</span><span class="sxs-lookup"><span data-stu-id="de3c1-130">Code, test, optimize.</span></span> <span data-ttu-id="de3c1-131">你可以为 API 调用例程创建一个函数，使其从脚本的其他部分重复使用，或在其他脚本中重复使用 (复制粘贴将变得更容易) 。</span><span class="sxs-lookup"><span data-stu-id="de3c1-131">You can create a function for your API call routine to make it reusable from other parts of your script or for reuse in a different script (copy-paste becomes much easier this way).</span></span>

## <a name="scenario"></a><span data-ttu-id="de3c1-132">方案</span><span class="sxs-lookup"><span data-stu-id="de3c1-132">Scenario</span></span>

<span data-ttu-id="de3c1-133">此脚本获取有关用户的 GitHub 存储库的基本信息。</span><span class="sxs-lookup"><span data-stu-id="de3c1-133">This script gets basic information about the user's GitHub repositories.</span></span>

![获取存储库信息示例](../../images/git.png)

## <a name="resources-used-in-the-sample"></a><span data-ttu-id="de3c1-135">示例中使用的资源</span><span class="sxs-lookup"><span data-stu-id="de3c1-135">Resources used in the sample</span></span>

1. [<span data-ttu-id="de3c1-136">获取存储库 Github API 参考。</span><span class="sxs-lookup"><span data-stu-id="de3c1-136">Get repositories Github API reference.</span></span>](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. <span data-ttu-id="de3c1-137">API 调用输出：转到 Web 浏览器或任何 HTTP 界面并键入 ，将 {USERNAME} 占位符替换为 `https://api.github.com/users/{USERNAME}/repos` Github ID。</span><span class="sxs-lookup"><span data-stu-id="de3c1-137">API call output: Go to a web browser or any HTTP interface and type in `https://api.github.com/users/{USERNAME}/repos`, replacing the {USERNAME} placeholder with your Github ID.</span></span>
1. <span data-ttu-id="de3c1-138">获取的信息：repo.name、repo.size、repo.owner.id、repo.license？。name</span><span class="sxs-lookup"><span data-stu-id="de3c1-138">Information fetched: repo.name, repo.size, repo.owner.id, repo.license?.name</span></span>

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="de3c1-139">示例代码：获取有关用户的 GitHub 存储库的基本信息</span><span class="sxs-lookup"><span data-stu-id="de3c1-139">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  name: string,
  id: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="de3c1-140">培训视频：如何进行外部 API 调用</span><span class="sxs-lookup"><span data-stu-id="de3c1-140">Training video: How to make external API calls</span></span>

<span data-ttu-id="de3c1-141">[![观看有关如何进行外部 API 调用的视频](../../images/api-vid.png)](https://youtu.be/fulP29J418E "如何进行外部 API 调用的视频")</span><span class="sxs-lookup"><span data-stu-id="de3c1-141">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
