---
title: Office脚本示例
description: 可用于Office脚本示例和方案。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 625db792763606e8db77abdc4665b7db2732892f
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232737"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office脚本示例和方案

本节包含[Office](../../overview/excel.md)脚本的自动化解决方案，帮助最终用户实现日常任务的自动化。 它包含业务用户面临的实际方案，并提供详细的解决方案以及分步说明视频链接。

对于 [Basics](#basics) 和 Beyond the [basics](#beyond-the-basics)中的每个项目，请查看源代码 [**、YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)视频等。

在 ["应用](#scenarios)场景"中，我们包含了几个演示实际用例的较大方案示例。

我们还欢迎 [来自社区的贡献](#community-contributions)。

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>基本信息

| Project | 详细信息 |
|---------|---------|
| [脚本基础知识](../excel-samples.md) | 这些示例演示了脚本的基本Office构建基块。 |
| [在内容中添加Excel](add-excel-comments.md) | 本示例演示如何向单元格添加注释，包括@mentioning添加注释。 |
| [将多个Excel表复制到单个表中](copy-tables-combine.md) | 此示例将来自多个Excel的数据组合到一个包含所有行的表中。 |

## <a name="beyond-the-basics"></a>超越基础设置

请查看以下端到端项目，该项目可自动执行示例方案以及 YouTube) 上承载的完整脚本、Excel 文件示例[ (。 ](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | 详细信息 |
|---------|---------|
| [计算特定工作表或所有工作表中的空行数](count-blank-rows.md) | 本示例检测工作表中是否有预计存在数据的空白行，然后报告空白行计数，以用于Power Automate流。 |
| [交叉引用和格式化Excel文件](excel-cross-reference.md) | 此解决方案演示如何使用脚本Excel脚本和脚本对两个Office文件进行交叉Power Automate。 |
| [电子邮件图表和表格图像](email-images-chart-table.md) | 此示例使用Office脚本Power Automate操作来创建图表，并通过电子邮件将图表作为图像发送。 |
| [外部提取调用](external-fetch-calls.md) | 此示例使用 `fetch` 从脚本的 GitHub获取信息。 |
| [筛选Excel并获取可见区域](filter-table-get-visible-range.md) | 此示例筛选一Excel，并作为 JSON 对象返回可见区域。 此 JSON 可以作为较大解决方案的Power Automate提供给一个流。 |
| [在工作簿中生成唯一标识符](document-number-generator.md) | 此方案可帮助用户生成具有特定格式的唯一文档编号，并添加一个范围或表中的条目。 |
| [在计算模式下管理Excel](excel-calculation.md) | 此示例演示如何在脚本中使用计算模式Excel web 版计算Office方法。 |
| [跨表移动行](move-rows-across-tables.md) | 此示例演示如何通过保存筛选器，然后处理和重新应用筛选器来跨表移动行。 |
| [输出Excel JSON](get-table-data.md) | 此解决方案演示如何将Excel数据输出为 JSON，以用于Power Automate。 |
| [从工作表的每个单元格中删除Excel超链接](remove-hyperlinks-from-cells.md) | 本示例清除当前工作表的所有超链接。 |
| [对文件夹中的所有 Excel 文件运行脚本](automate-tasks-on-all-excel-files-in-folder.md) | 此项目对位于 OneDrive for Business (文件夹中的所有文件执行一组自动化任务，SharePoint文件夹) 。 该代码对Excel文件执行计算，添加格式，并插入一个@mentions注释。 |
| [从Teams数据发送Excel会议](send-teams-invite-from-excel-data.md) | 此解决方案演示如何使用 Office 脚本和 Power Automate 操作从 Excel 文件选择行，并使用它发送 Teams 会议邀请，然后更新Excel。 |

## <a name="scenarios"></a>应用场景

Office脚本可以自动执行日常例程的某些部分。 这些日常任务通常存在于独特的生态系统中，其中Excel以特定方式设置的特定工作簿。 这些较大的方案示例演示了此类实际用例。 它们包括Office脚本和工作簿，因此您可从头到尾查看方案。

| 应用场景 | 详细信息 |
|---------|---------|
| [分析 Web 下载项](../scenarios/analyze-web-downloads.md) | 此方案具有分析 Web 流量记录的脚本，以确定用户的来源国家/地区。 它展示文本分析、在脚本中使用子功能、应用条件格式和使用表的技能。 |
| [从 NOAA 中提取图形水级别的数据](../scenarios/noaa-data-fetch.md) | 此方案使用Office脚本从 NOAA ([Currents](https://tidesandcurrents.noaa.gov/)数据库的外部源提取数据) 并绘制结果信息的图形。 它重点介绍了使用 `fetch` 获取数据和使用图表的技能。 |
| [成绩计算器](../scenarios/grade-calculator.md) | 此方案具有一个脚本，用于验证教师的课堂成绩记录。 它展示错误检查、单元格格式设置和正则表达式的技能。 |
| [任务提醒](../scenarios/task-reminders.md) | 此方案在Office流Power Automate脚本，向同事发送更新项目状态的提醒。 它重点介绍了脚本Power Automate和数据传输的专业技能。 |

## <a name="community-contributions"></a>社区参与

欢迎来自[我们的](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md)脚本Office贡献！ 可随意创建拉取请求进行审阅。

| Project | 详细信息 |
|---------|---------|
| [四年问候语动画](community-seasons-greetings.md) | 此脚本由 [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) 在假日假日的快乐中贡献！ 这是一个有趣的脚本，它使用脚本在Excel web 版中显示一Office树。 |

## <a name="try-it-out"></a>试用

这些示例是开源的。 尝试一下。 你将需要来自工作或学校的 Microsoft 工作或学校帐户，并拥有许可证才能Microsoft 365 E3 (或) 。 只需前往登录 https://office.com 帐户并开始使用。

## <a name="leave-a-comment"></a>留下注释

使用特定示例的文档页面底部的"反馈"部分，可以随意留下评论、提出建议或记录问题。
