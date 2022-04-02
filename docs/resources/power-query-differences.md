---
title: 何时使用 Power Query 或 Office 脚本
description: 最适合脚本平台和脚本Power Query Office方案。
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: e91077d635d66dde692c129bdd4b2f32657d5283
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585903"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>何时使用 Power Query 或 Office 脚本

[Power Query](https://powerquery.microsoft.com)脚本Office脚本都是功能强大的自动化解决方案，Excel。 这两个解决方案Excel清理和转换工作簿数据。 可以刷新Power Query脚本或Office脚本，并再次运行新数据以产生一致的结果，从而节省时间并更快地处理生成的信息。

本文概述了何时可能支持一个平台，而何时支持另一个平台。 通常，Power Query从大型外部数据源和 Office 脚本提取和转换数据非常适用于快速、Excel的解决方案和Power Automate[集成](../develop/power-automate-integration.md)。

## <a name="large-data-sources-and-data-retrieval-power-query"></a>大型数据源和数据检索：Power Query

在处理来自Power Query平台的数据源时，我们建议使用资源库。

Power Query[具有到数百个源的](https://powerquery.microsoft.com/connectors/)内置数据连接。 Power Query专门设计用于数据检索、转换和组合任务。 当你需要其中一个源的数据时，Power Query一种无代码方法将数据引入到Excel形状中。

这些Power Query连接专为大型数据集设计。 它们与用户或用户没有[](../testing/platform-limits.md)Power Automate Excel 网页版。

Office脚本为小型数据源或连接器未涵盖的数据源Power Query解决方案。 这包括[使用`fetch`或 REST API](../develop/external-calls.md)，或者从临时数据源（如自适应卡片）[Teams信息](../resources/scenarios/task-reminders.md)。

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>格式、可视化和编程控件：Office脚本

我们建议你Office导入和转换数据时，使用脚本。

几乎所有可以通过自定义 UI 手动执行Excel操作都可以通过脚本Office实现。 它们非常适用于将一致的格式应用于工作簿。 脚本创建图表、数据透视表、形状、图像和其他工作表可视化。 脚本还可以精确控制这些可视化效果的位置、大小、颜色和其他属性。

包含 TypeScript 代码可让你进行高度自定义。 语句等编程控制 `if...else` 逻辑使脚本可靠。 这样，您即可执行一些操作，如按条件读取数据，而不依赖复杂的Excel公式，或在更改工作簿之前扫描工作簿中的意外更改。

格式可以通过自定义模板Power Query Excel[应用](https://templates.office.com/power-query-tutorial-tm11414620)。 但是，模板在个人或组织级别进行更新，而Office脚本提供了更精细的访问控制。

## <a name="power-automate-integrations"></a>Power Automate集成

Office脚本提供了更多用于集成Power Automate选项。 脚本专为您的解决方案而定制。 定义脚本 [的输入和输出，](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts)以便它适用于流中任何其他连接器或数据。 以下屏幕截图显示了一个Power Automate流，该流将数据从自适应卡片Teams到Office脚本。

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="显示流设计器中的 Excel Online (Business) 连接器的屏幕截图。连接器使用 Run 脚本操作从自适应卡片Teams输入，然后向脚本提供。":::

Power Query连接器中SQL Server [Power Automate](https://powerquery.microsoft.com/flow/)连接器。 The [Transform data using Power Query](/connectors/sql/#transform-data-using-power-query) action lets you build a query in Power Automate. 虽然这是一个与 SQL Server一起使用的强大工具，但它Power Query输入源，如以下流屏幕截图中所示。

:::image type="content" source="../images/power-query-flow-option.png" alt-text="显示流设计器中的SQL Server连接器的屏幕截图。连接器使用&quot;使用转换数据转换Power Query操作。":::

## <a name="platform-dependencies"></a>平台依赖项

Office脚本当前仅适用于Excel web 版。 Power Query当前仅适用于桌面Excel应用。 这两者均可以通过 Power Automate，从而允许流处理存储在Excel工作簿OneDrive。

## <a name="see-also"></a>另请参阅

- [Power Query门户](https://powerquery.microsoft.com/)
- [Power Query Excel](https://powerquery.microsoft.com/excel/)
- [使用Office脚本运行Power Automate](../develop/power-automate-integration.md)
