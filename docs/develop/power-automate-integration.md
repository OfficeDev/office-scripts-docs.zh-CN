---
title: 使用 Power Automate 运行 Office 脚本
description: 如何获取使用 Power Automate 工作流Excel web 版的 Office 脚本。
ms.date: 06/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 61e51861bd2c987c25d40e9ac6d2247122256918
ms.sourcegitcommit: c5ffe0a95b962936ee92e7ffe17388bef6d4fad8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/29/2022
ms.locfileid: "66241839"
---
# <a name="run-office-scripts-with-power-automate"></a>使用 Power Automate 运行 Office 脚本

[借助 Power Automate](https://flow.microsoft.com) ，可将 Office 脚本添加到更大的自动化工作流。 可以使用 Power Automate 执行诸如将电子邮件内容添加到工作表表或基于工作簿注释在项目管理工具中创建操作之类的操作。

## <a name="get-started"></a>入门

如果你不熟悉 Power Automate，建议访问 [Power Automate 入](/power-automate/getting-started)门。 可在此处详细了解所有可用的自动化可能性。 此处的文档重点介绍 Office 脚本如何与 Power Automate 配合使用，以及如何帮助改善 Excel 体验。

### <a name="step-by-step-tutorials"></a>分步教程

Power Automate 和 Office 脚本有三个分步教程。 这些演示如何合并自动服务，并在工作簿和流之间传递数据。

- [通过 Power Automate 手动流调用脚本](../tutorials/excel-power-automate-manual.md)
- [将数据传递到自动运行的 Power Automate 流中的脚本](../tutorials/excel-power-automate-trigger.md)
- [从脚本返回数据到自动运行 Power Automated 流](../tutorials//excel-power-automate-returns.md)

### <a name="create-a-flow-from-excel"></a>从 Excel 创建流

可以通过各种流模板开始使用 Excel 中的 Power Automate。 在 **“自动执行”** 选项卡下，选择 **“自动执行任务**”。

:::image type="content" source="../images/automate-a-task-button.png" alt-text="功能区中的“自动执行任务”按钮。":::

这将打开一个任务窗格，其中包含多个选项，用于开始将 Office 脚本连接到更大的自动化解决方案。 选择要开始的任何选项。 流随当前工作簿一起提供。

:::image type="content" source="../images/automate-a-task-choices.png" alt-text="显示流模板选项的任务窗格，例如“计划 Office 脚本在 Excel 中运行，然后发送电子邮件”和“收到Microsoft Forms响应时在 Excel 中运行 Office 脚本”。":::

> [!TIP]
> 还可以从单个脚本 **上的“更多”选项 (...)** 菜单开始创建流。

## <a name="excel-online-business-connector"></a>Excel Online (Business) 连接器

[连接器](/connectors/connectors) 是 Power Automate 与应用程序之间的桥梁。 [Excel Online (Business) 连接器](/connectors/excelonlinebusiness)为流提供对 Excel 工作簿的访问权限。 通过“运行脚本”操作，可以调用可通过所选工作簿访问的任何 Office 脚本。 还可以为脚本提供输入参数，以便流可以提供数据，或者让脚本返回信息，以便在流中执行后续步骤。

> [!IMPORTANT]
> “运行脚本”操作使使用 Excel 连接器的人员能够对工作簿及其数据进行大量访问。 此外，执行外部 API 调用的脚本存在安全风险，如 [Power Automate 的外部调用](external-calls.md)中所述。 如果你的管理员担心高度敏感数据的泄露，他们可以关闭 Excel Online 连接器或通过 [Office 脚本管理员控制限制对 Office 脚本](/microsoft-365/admin/manage/manage-office-scripts-settings)的访问。

> [!IMPORTANT]
> Power Automate **目前不** 支持存储在 SharePoint 上的脚本。

## <a name="data-transfer-in-flows-for-scripts"></a>脚本的流中的数据传输

通过 Power Automate，可以在流的步骤之间传递数据片段。 脚本可以配置为接受所需的任何类型的信息，并从工作簿返回所需的任何信息。 除了) ，还通过向函数 (添加参数来 `main` `workbook: ExcelScript.Workbook` 指定脚本的输入。 脚本的输出是通过向其添加返回类型来声明的 `main`。

> [!NOTE]
> 在流中创建“运行脚本”块时，会填充接受的参数和返回的类型。 如果更改脚本的参数或返回类型，则需要重新创建流的“运行脚本”块。 这可确保正确分析数据。

以下部分介绍 Power Automate 中使用的脚本的输入和输出的详细信息。 如果想要使用实践方法来学习本主题，请 [在自动运行的 Power Automate 流教程中尝试将数据传递给脚本](../tutorials/excel-power-automate-trigger.md) ，或浏览 [自动任务提醒](../resources/scenarios/task-reminders.md) 示例方案。

### <a name="main-parameters-pass-data-to-a-script"></a>`main` 参数：将数据传递到脚本

所有脚本输入都指定为函 `main` 数的其他参数。 例如，如果希望脚本接受 `string` 表示名称作为输入的脚本，则会将签名更改 `main` 为 `function main(workbook: ExcelScript.Workbook, name: string)`。

在 Power Automate 中配置流时，可以将脚本输入指定为静态值、 [表达式](/power-automate/use-expressions-in-conditions)或动态内容。 有关单个服务连接器的详细信息，请参阅 [Power Automate 连接器文档](/connectors/)。

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

脚本可以从工作簿返回数据，以用作 Power Automate 流中的动态内容。 [以前列出的相同类型限制](#type-restrictions)适用于返回类型。 若要返回对象，请将返回类型语法添加到函 `main` 数。 例如，如果要从脚本返回值 `string` ，则 `main` 签名将为 `function main(workbook: ExcelScript.Workbook): string`。

## <a name="example"></a>示例

以下屏幕截图显示了每当向你分配 [GitHub](https://github.com/) 问题时触发的 Power Automate 流。 该流运行一个脚本，该脚本将问题添加到 Excel 工作簿中的表中。 如果该表中有五个或更多问题，则流将发送电子邮件提醒。

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="显示示例流的 Power Automate 流编辑器。":::

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
- [Power Automate 与 Office 脚本的故障排除信息](../testing/power-automate-troubleshooting.md)
- [Power Automate 入门](/power-automate/getting-started)
- [Excel Online (Business) 连接器参考文档](/connectors/excelonlinebusiness/)
