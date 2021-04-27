---
title: 在脚本中使用外部Office调用
description: 了解如何在脚本中执行外部 API Office调用。
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: a77ceb61c2ff46a7b6226b798462b7be2c8e1c54
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026990"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>在脚本中使用外部Office调用

此脚本获取有关用户存储库GitHub信息。 它显示了如何在 `fetch` 简单方案中使用。

你可以了解有关正在应用 API 参考中使用的 GItHub GITHUB [API。](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) 您还可以通过访问 Web 浏览器中的内容来查看原始 API 调用输出 (请务必将 {USERNAME} 占位符替换为 `https://api.github.com/users/{USERNAME}/repos` Github ID) 。

![获取存储库信息示例](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>示例代码：获取有关用户数据库GitHub信息

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
