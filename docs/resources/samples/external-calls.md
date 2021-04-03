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
# <a name="external-api-calls-from-office-scripts"></a>Office 脚本中的外部 API 调用

Office 脚本允许 [有限的外部 API 调用支持](../../develop/external-calls.md)。

> [!IMPORTANT]
>
> * 无法登录或使用 OAuth2 类型的身份验证流。 所有密钥和凭据必须硬编码 (源文件进行硬编码) 。
> * 没有用于存储 API 凭据和密钥的基础结构。 这必须由用户管理。
> * 外部调用可能会导致向不需要的终结点公开敏感数据，或导致外部数据进入内部工作簿。 管理员可以针对此类呼叫建立防火墙保护。 在依赖外部调用之前，请务必检查本地策略。
> * 如果脚本使用 API 调用，则它在 Power Automate 方案中将无法正常工作。 您必须使用 Power Automate 的 HTTP 操作或等效操作从外部服务提取数据或将其推送到外部服务。
> * 外部 API 调用涉及异步 API 语法，并且需要稍微高级了解异步通信的工作方式。
> * 请务必在依赖关系之前检查数据吞吐量。 例如，下拉整个外部数据集可能不是最佳选择，而应该使用分页获取区块中的数据。

## <a name="useful-knowledge-and-resources"></a>有用的知识和资源

* [REST API：](https://en.wikipedia.org/wiki/Representational_state_transfer)最有可能使用 API 调用的方式。
* [ `async` ：了解工作原理 `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)。
* [`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch)：了解工作原理。

## <a name="steps"></a>步骤

1. 通过添加 `main` 前缀将函数标记为异步 `async` 函数。 例如，`async function main(workbook: ExcelScript.Workbook)`。
1. 你进行哪种类型的 API 调用？ `GET`, `POST`, `PUT`, `DELETE`, `PATCH`? 有关详细信息，请参阅 REST API 材料。
1. 获取服务 API 终结点、身份验证要求、标头等。
1. 定义输入或输出 `interface` 以帮助完成代码和开发时间验证。 有关详细信息 [，](#training-video-how-to-make-external-api-calls) 请参阅视频。
1. 代码、测试、优化。 你可以为 API 调用例程创建一个函数，使其从脚本的其他部分重复使用，或在其他脚本中重复使用 (复制粘贴将变得更容易) 。

## <a name="scenario"></a>方案

此脚本获取有关用户的 GitHub 存储库的基本信息。

![获取存储库信息示例](../../images/git.png)

## <a name="resources-used-in-the-sample"></a>示例中使用的资源

1. [获取存储库 Github API 参考。](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. API 调用输出：转到 Web 浏览器或任何 HTTP 界面并键入 ，将 {USERNAME} 占位符替换为 `https://api.github.com/users/{USERNAME}/repos` Github ID。
1. 获取的信息：repo.name、repo.size、repo.owner.id、repo.license？。name

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>示例代码：获取有关用户的 GitHub 存储库的基本信息

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

## <a name="training-video-how-to-make-external-api-calls"></a>培训视频：如何进行外部 API 调用

[![观看有关如何进行外部 API 调用的视频](../../images/api-vid.png)](https://youtu.be/fulP29J418E "如何进行外部 API 调用的视频")
