---
title: Office 脚本中的外部 API 呼叫支持
description: 在Office脚本中进行外部 API 调用的支持和指导。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: b847400893184533c250ab99b640563ff0cbdb3e
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088041"
---
# <a name="external-api-call-support-in-office-scripts"></a>Office 脚本中的外部 API 呼叫支持

脚本支持对外部服务的调用。 使用这些服务向工作簿提供数据和其他信息。

> [!CAUTION]
> 外部调用可能导致敏感数据暴露到不受欢迎的终结点。 管理员可以针对此类调用建立防火墙保护。

> [!IMPORTANT]
> 对外部 API 的调用只能通过Excel应用程序进行，而不能通过[正常情况下](#external-calls-from-power-automate)的Power Automate进行。

## <a name="configure-your-script-for-external-calls"></a>为外部调用配置脚本

外部调用是 [异步的](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) ，要求脚本标记为 `async`。 将 `async` 前缀添加到函 `main` 数并让其返回， `Promise`如下所示：

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> 返回其他信息的脚本可以返回 `Promise` 该类型的脚本。 例如，如果脚本需要返回 `Employee` 对象，则返回签名将为 `: Promise <Employee>`

需要了解外部服务的接口才能对该服务进行调用。 如果使用的 `fetch` 是 [REST API 或 REST API](https://wikipedia.org/wiki/Representational_state_transfer)，则需要确定返回的数据的 JSON 结构。 对于脚本的输入和输出，请考虑使之 `interface` 与所需的 JSON 结构相匹配。 这为脚本提供了更多类型安全性。 可以在[“使用从Office脚本提取](../resources/samples/external-fetch-calls.md)”中看到此示例。

### <a name="limitations-with-external-calls-from-office-scripts"></a>来自Office脚本的外部调用的限制

* 无法登录或使用 OAuth2 类型的身份验证流。 所有密钥和凭据都必须硬编码 (或从其他源) 读取。
* 没有用于存储 API 凭据和密钥的基础结构。 这必须由用户管理。
* 不支持文档 Cookie `localStorage`和 `sessionStorage` 对象。
* 外部调用可能导致敏感数据公开到不受欢迎的终结点，或将外部数据引入内部工作簿。 管理员可以针对此类调用建立防火墙保护。 在依赖外部调用之前，请务必检查本地策略。
* 在获取依赖项之前，请务必检查数据吞吐量。 例如，向下拉取整个外部数据集可能不是最佳选项，而应使用分页以区块形式获取数据。

## <a name="retrieve-information-with-fetch"></a>使用 >0> `fetch`

[提取 API](https://developer.mozilla.org/docs/Web/API/Fetch_API) 从外部服务检索信息。 它是一个 `async` API，因此需要调整 `main` 脚本的签名。 使函 `main` 数 `async`。 还应确保`await``fetch`调用和`json`检索。 这可确保这些操作在脚本结束之前完成。

检 `fetch` 索到的任何 JSON 数据都必须与脚本中定义的接口匹配。 返回的值必须分配给特定类型，因为[Office脚本不支持该`any`类型](typescript-restrictions.md#no-any-type-in-office-scripts)。 应参阅服务的文档，了解返回的属性的名称和类型。 然后，将匹配的接口或接口添加到脚本。

以下脚本用于 `fetch` 从给定 URL 中的测试服务器检索 JSON 数据。 `JSONData`请注意将数据存储为匹配类型的接口。

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
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

### <a name="other-fetch-samples"></a>其他 `fetch` 示例

* Office[脚本示例中的“使用外部提取调](../resources/samples/external-fetch-calls.md)用”演示如何获取有关用户GitHub存储库的基本信息。
* [Office脚本示例方案：来自 NOAA 的Graph水位数据](../resources/scenarios/noaa-data-fetch.md)演示了用于从国家海洋和大气管理局的潮汐和电流数据库中检索记录的提取命令。

## <a name="external-calls-from-power-automate"></a>来自Power Automate的外部调用

使用Power Automate运行脚本时，任何外部 API 调用都失败。 这是通过Excel应用程序运行脚本和通过Power Automate运行脚本之间的行为差异。 在将这些引用生成到流之前，请务必检查脚本是否具有此类引用。

必须将 [HTTP 与 Azure AD](/connectors/webcontents/) 或其他等效操作配合使用，才能从外部服务中提取数据或将其推送到外部服务。

> [!WARNING]
> 通过Power Automate [Excel联机连接器](/connectors/excelonlinebusiness)进行的外部调用失败，以帮助维护现有的数据丢失防护策略。 但是，通过Power Automate运行的脚本会在组织外部和组织防火墙之外执行此操作。 为了在此外部环境中保护恶意用户，管理员可以控制Office脚本的使用。 管理员可以在Power Automate中禁用 Excel Online 连接器，也可以通过Office[脚本管理员控件关闭Excel web 版Office脚本](/microsoft-365/admin/manage/manage-office-scripts-settings)。

## <a name="see-also"></a>另请参阅

* [使用 JSON 将数据传入和传入Office脚本](use-json.md)
* [在 Office 脚本中使用内置的 JavaScript 对象](javascript-objects.md)
* [在 Office 脚本中使用外部提取呼叫](../resources/samples/external-fetch-calls.md)
* [Office脚本示例方案：Graph来自 NOAA 的水位数据](../resources/scenarios/noaa-data-fetch.md)
