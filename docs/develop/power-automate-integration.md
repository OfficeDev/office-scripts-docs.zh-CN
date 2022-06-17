---
title: 使用Power Automate运行Office脚本
description: 如何获取Excel web 版使用Power Automate工作流的Office脚本。
ms.date: 05/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85c335eeb736ec544eccb2fbdbe819bdbef6848c
ms.sourcegitcommit: aecbd5baf1e2122d836c3eef3b15649e132bc68e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/16/2022
ms.locfileid: "66128228"
---
# <a name="run-office-scripts-with-power-automate"></a>使用Power Automate运行Office脚本

[Power Automate](https://flow.microsoft.com)可将Office脚本添加到更大的自动化工作流。 可以使用Power Automate执行诸如将电子邮件内容添加到工作表表或基于工作簿注释在项目管理工具中创建操作之类的操作。

## <a name="get-started"></a>入门

如果你不熟悉Power Automate，建议[使用Power Automate访问开始](/power-automate/getting-started)。 可在此处详细了解所有可用的自动化可能性。 此处的文档重点介绍Office脚本如何使用Power Automate，以及如何帮助改善Excel体验。

若要开始合并Power Automate和Office脚本，请按照教程["开始"菜单使用脚本和Power Automate](../tutorials/excel-power-automate-manual.md)。 这将教你如何创建调用简单脚本的流。 完成本教程并将[数据传递给自动运行Power Automate流教程中的脚本](../tutorials/excel-power-automate-trigger.md)后，返回此处，了解有关将Office脚本连接到Power Automate流的详细信息。

## <a name="excel-online-business-connector"></a>Excel联机 (业务) 连接器

[连接器](/connectors/connectors)是Power Automate和应用程序之间的桥梁。 [Excel联机 (业务) 连接器](/connectors/excelonlinebusiness)允许流访问Excel工作簿。 通过“运行脚本”操作，可以调用可通过所选工作簿访问的任何Office脚本。 还可以为脚本提供输入参数，以便流可以提供数据，或者让脚本返回信息，以便在流中执行后续步骤。

> [!IMPORTANT]
> “运行脚本”操作为使用Excel连接器的人员提供了对工作簿及其数据的重要访问权限。 此外，执行外部 API 调用的脚本存在安全风险，如[来自 Power Automate 的外部调](external-calls.md)用中所述。 如果你的管理员担心高度敏感数据的泄露，他们可以关闭Excel联机连接器，或者通过Office[脚本管理员控件限制对Office脚本](/microsoft-365/admin/manage/manage-office-scripts-settings)的访问。

> [!IMPORTANT]
> Power Automate目前 **不** 支持存储在SharePoint上的脚本。

## <a name="data-transfer-in-flows-for-scripts"></a>脚本的流中的数据传输

Power Automate使你可以在流的步骤之间传递数据片段。 脚本可以配置为接受所需的任何类型的信息，并从工作簿返回所需的任何信息。 除了) ，还通过向函数 (添加参数来 `main` `workbook: ExcelScript.Workbook` 指定脚本的输入。 脚本的输出是通过向其添加返回类型来声明的 `main`。

> [!NOTE]
> 在流中创建“运行脚本”块时，会填充接受的参数和返回的类型。 如果更改脚本的参数或返回类型，则需要重新创建流的“运行脚本”块。 这可确保正确分析数据。

以下部分介绍Power Automate中使用的脚本的输入和输出的详细信息。 如果想要使用动手方法来学习本主题，请[在自动运行Power Automate流教程中尝试将数据传递给脚本](../tutorials/excel-power-automate-trigger.md)，或探索[自动任务提醒](../resources/scenarios/task-reminders.md)示例方案。

### <a name="main-parameters-pass-data-to-a-script"></a>`main` 参数：将数据传递到脚本

所有脚本输入都指定为函 `main` 数的其他参数。 例如，如果希望脚本接受 `string` 表示名称作为输入的脚本，则会将签名更改 `main` 为 `function main(workbook: ExcelScript.Workbook, name: string)`。

在Power Automate中配置流时，可以将脚本输入指定为静态值、[表达式](/power-automate/use-expressions-in-conditions)或动态内容。 有关单个服务连接器的详细信息，请参[阅Power Automate连接器文档](/connectors/)。

#### <a name="type-restrictions"></a>类型限制

将输入参数添加到脚本函 `main` 数时，请考虑以下限制和限制。 这些也适用于脚本的返回类型。

1. 第一个参数必须为类型 `ExcelScript.Workbook`。 其参数名称并不重要。

1. 支持类型`string``number`、类型`boolean``unknown`、类型`object`和`undefined`类型。

1. 支持 (`[]` 数组和 `Array<T>` 以前列出类型的样式) 。 也支持嵌套数组。

1. 如果联合类型是属于单个类型 (（例如 `"Left" | "Right"`，而不是 `"Left", 5`) ）的文本的联合，则允许联合类型。 支持未定义类型的联合也支持 (，例如 `string | undefined`) 。

1. 如果对象类型包含类型`string`、、`number``boolean`支持的数组或其他受支持的对象的属性，则允许它们。 以下示例显示了作为参数类型支持的嵌套对象。

    ```TypeScript
    // The Employee object is supported because Position is also composed of supported types.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

1. 对象必须在脚本中定义其接口或类定义。 也可以匿名内联定义对象，如以下示例所示。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

#### <a name="optional-and-default-parameters"></a>可选参数和默认参数

1. 允许使用可选参数，并使用可选修饰符 `?` (（例如) `function main(workbook: ExcelScript.Workbook, Name?: string)` ）进行表示。

1. 例如 `function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`，允许 (默认参数值。

### <a name="return-data-from-a-script"></a>从脚本返回数据

脚本可以从工作簿返回数据，以用作Power Automate流中的动态内容。 [以前列出的相同类型限制](#type-restrictions)适用于返回类型。 若要返回对象，请将返回类型语法添加到函 `main` 数。 例如，如果要从脚本返回值 `string` ，则 `main` 签名将为 `function main(workbook: ExcelScript.Workbook): string`。

## <a name="example"></a>示例

以下屏幕截图显示了在向你分配GitHub问题时触发的[Power Automate](https://github.com/)流。 该流运行一个脚本，该脚本将问题添加到Excel工作簿中的表。 如果该表中有五个或更多问题，则流将发送电子邮件提醒。

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="显示示例流的Power Automate流编辑器。":::

脚 `main` 本的函数将问题 ID 和问题标题指定为输入参数，脚本返回问题表中的行数。

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a>另请参阅

- [通过 Power Automate 手动流调用脚本](../tutorials/excel-power-automate-manual.md)
- [将数据传递到自动运行的 Power Automate 流中的脚本](../tutorials/excel-power-automate-trigger.md)
- [从脚本返回数据到自动运行 Power Automated 流](../tutorials/excel-power-automate-returns.md)
- [使用Office脚本排查Power Automate的信息](../testing/power-automate-troubleshooting.md)
- [Power Automate 入门](/power-automate/getting-started)
- [Excel Online (Business) 连接器参考文档](/connectors/excelonlinebusiness/)
