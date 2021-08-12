---
title: 在 Office 脚本中使用外部提取呼叫
description: 了解如何在脚本中执行外部 API Office调用。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 46e83e290224d03cdb1f87eeb6215e2fb9510189c210a7dfa330f438ced0a53e
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847526"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>在 Office 脚本中使用外部提取呼叫

此脚本获取有关用户存储库GitHub信息。 它显示了如何在 `fetch` 简单方案中使用。 有关使用或其他外部调用的信息，请阅读外部 `fetch` [API 调用支持（Office脚本）](../../develop/external-calls.md)

你可以了解有关正在应用 API 参考中使用的 GItHub GITHUB [API。](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) 您还可以通过访问 Web 浏览器来查看原始 API 调用输出 (请务必将 {USERNAME} 占位符替换为 GitHub `https://api.github.com/users/{USERNAME}/repos` ID) 。

![获取存储库信息示例](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>示例代码：获取有关用户存储库GitHub信息

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }

  // Add the data to the current worksheet, starting at "A2".
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
}

// An interface matching the returned JSON for a GitHub repository.
interface Repository {
  name: string,
  id: string,
  license?: License 
}

// An interface matching the returned JSON for a GitHub repo license.
interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a>培训视频：如何进行外部 API 调用

[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/fulP29J418E)。
