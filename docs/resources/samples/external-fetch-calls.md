---
title: 在 Office 脚本中使用外部提取呼叫
description: 了解如何在脚本中执行外部 API Office调用。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 721bfa39eea1e9973efc7fd13efa5bac734b76dd
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232520"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="56b8c-103">在 Office 脚本中使用外部提取呼叫</span><span class="sxs-lookup"><span data-stu-id="56b8c-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="56b8c-104">此脚本获取有关用户存储库GitHub信息。</span><span class="sxs-lookup"><span data-stu-id="56b8c-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="56b8c-105">它显示了如何在 `fetch` 简单方案中使用。</span><span class="sxs-lookup"><span data-stu-id="56b8c-105">It shows how to use `fetch` in a simple scenario.</span></span>

<span data-ttu-id="56b8c-106">你可以了解有关正在应用 API 参考中使用的 GItHub GITHUB [API。](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)</span><span class="sxs-lookup"><span data-stu-id="56b8c-106">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="56b8c-107">您还可以通过访问 Web 浏览器中的内容来查看原始 API 调用输出 (请务必将 {USERNAME} 占位符替换为 `https://api.github.com/users/{USERNAME}/repos` Github ID) 。</span><span class="sxs-lookup"><span data-stu-id="56b8c-107">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your Github ID).</span></span>

![获取存储库信息示例](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="56b8c-109">示例代码：获取有关用户数据库GitHub信息</span><span class="sxs-lookup"><span data-stu-id="56b8c-109">Sample code: Get basic information about user's GitHub repositories</span></span>

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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="56b8c-110">培训视频：如何进行外部 API 调用</span><span class="sxs-lookup"><span data-stu-id="56b8c-110">Training video: How to make external API calls</span></span>

<span data-ttu-id="56b8c-111">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/fulP29J418E)。</span><span class="sxs-lookup"><span data-stu-id="56b8c-111">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/fulP29J418E).</span></span>
