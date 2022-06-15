---
title: 在 Office 脚本中使用外部提取呼叫
description: 了解如何在Office脚本中进行外部 API 调用。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 569d74f1ca8996cd8fe8a4ba3163445d57676d27
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088090"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>在 Office 脚本中使用外部提取呼叫

此脚本获取有关用户GitHub存储库的基本信息。 它演示如何在简单方案中使用 `fetch` 。 有关使用或其他外部调用`fetch`的详细信息，请阅读[Office脚本中的外部 API 调用支持](../../develop/external-calls.md)。 有关使用 [JSON]] (https://www.w3schools.com/whatis/whatis_json.asp)对象的信息（例如GitHub API 返回的内容），请阅读[使用 JSON 将数据传入和传入Office脚本](../../develop/use-json.md)。

详细了解在 [GitHub API 参考](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)中使用的 GItHub API。 还可以通过访问 `https://api.github.com/users/{USERNAME}/repos` Web 浏览器来查看原始 API 调用输出 (请务必将 {USERNAME} 占位符替换为GitHub ID) 。

![获取存储库信息示例](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>示例代码：获取有关用户GitHub存储库的基本信息

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();

  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos) {
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url]);
  }
  // Create a header row.
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange('A1:D1').setValues([["ID", "Name", "License Name", "License URL"]]);

  // Add the data to the current worksheet, starting at "A2".
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

[观看苏迪 · 拉马穆尔西在 YouTube 上浏览这个示例](https://youtu.be/fulP29J418E)。
