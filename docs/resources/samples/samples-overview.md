---
title: Office 脚本示例
description: 可用的 Office 脚本示例和方案。
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de0e99cbac7fcdeb1a3d3c43dd72ce53ed5847dd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571214"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office 脚本示例和方案

本节包含 [基于 Office 脚本](../../overview/excel.md) 的自动化解决方案，帮助最终用户实现日常任务的自动化。 它包含业务用户面临的实际方案，并提供详细的解决方案以及分步说明视频链接。

对于 [Basics](#basics) 和 Beyond the [basics](#beyond-the-basics)中的每个项目，请查看源代码 [**、YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)视频等。

在 ["应用](#scenarios)场景"中，我们包含了几个演示实际用例的较大方案示例。

我们还欢迎 [来自社区的贡献](#community-contributions)。

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>基本信息

| Project | 详细信息 |
|---------|---------|
| [脚本基础知识](../excel-samples.md) | 这些示例演示 Office 脚本的基本构建基块。 |
| [了解有关在 Office 脚本中使用 Range 对象的基础知识](range-basics.md) | 本文介绍使用 Range 对象及其 API 的基础知识。 这是一个基础主题，将在所有其他项目中使用。 |

## <a name="beyond-the-basics"></a>除基础知识外

请查看以下端到端项目，该项目可自动执行示例方案以及完整脚本、使用的示例 Excel 文件 [和视频](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)。

| Project | 详细信息 |
|---------|---------|
| [在 Excel 中添加注释](add-excel-comments.md) | 本示例演示如何向单元格添加注释，包括@mentioning添加注释。 |
| [计算特定工作表或所有工作表中的空行数](count-blank-rows.md) | 此示例检测工作表中是否有预计存在数据的空白行，然后报告空白行计数以用于 Power Automate 流。 |
| [交叉引用和格式化 Excel 文件](excel-cross-reference.md) | 此解决方案演示如何使用 Office 脚本和 Power Automate 交叉引用两个 Excel 文件并设置其格式。 |
| [电子邮件图表和表格图像](email-images-chart-table.md) | 此示例使用 Office 脚本和 Power Automate 操作创建图表，并通过电子邮件将图表作为图像发送。 |
| [筛选 Excel 表并获取可见区域](filter-table-get-visible-range.md) | 此示例筛选 Excel 表，并返回可见区域作为 JSON 对象。 此 JSON 可以作为较大解决方案的一部分提供给 Power Automate 流。 |
| [在工作簿中生成唯一标识符](document-number-generator.md) | 此方案可帮助用户生成具有特定格式的唯一文档编号，并添加一个范围或表中的条目。 |
| [在 Excel 中管理计算模式](excel-calculation.md) | 此示例演示如何使用 Office 脚本在 Excel 网页中使用计算模式和计算方法。 |
| [将多个 Excel 表合并到单个表中](copy-tables-combine.md) | 此示例将多个 Excel 表中的数据合并到一个包含所有行的表中。 |
| [跨表移动行](move-rows-across-tables.md) | 此示例演示如何通过保存筛选器，然后处理和重新应用筛选器来跨表移动行。 |
| [将 Excel 数据输出为 JSON](get-table-data.md) | 此解决方案演示如何将 Excel 表数据输出为 JSON 以在 Power Automate 中使用。 |
| [从 Excel 工作表的每个单元格中删除超链接](remove-hyperlinks-from-cells.md) | 本示例清除当前工作表的所有超链接。 |
| [对文件夹中的所有 Excel 文件运行脚本](automate-tasks-on-all-excel-files-in-folder.md) | 此项目对位于 OneDrive for Business (上的文件夹中的所有文件执行一组自动化任务，也可以用于 SharePoint 文件夹) 。 该代码对 Excel 文件执行计算，添加格式，并插入一个@mentions批注。 |
| [从 Excel 数据发送 Teams 会议](send-teams-invite-from-excel-data.md) | 此解决方案演示如何使用 Office 脚本和 Power Automate 操作从 Excel 文件选择行，并使用它发送 Teams 会议邀请，然后更新 Excel。 |

## <a name="scenarios"></a>应用场景

Office 脚本可以自动执行日常例程的某些部分。 这些日常任务通常存在于独特的生态系统中，其中以特定方式设置的 Excel 工作簿。 这些较大的方案示例演示了此类实际用例。 它们同时包括 Office 脚本和工作簿，因此您可从头到尾查看方案。

| 方案 | 详细信息 |
|---------|---------|
| [分析 Web 下载项](../scenarios/analyze-web-downloads.md) | 此方案具有分析 Web 流量记录的脚本，以确定用户的来源国家/地区。 它展示文本分析、在脚本中使用子功能、应用条件格式和使用表的技能。 |
| [从 NOAA 中提取图形水级别的数据](../scenarios/noaa-data-fetch.md) | 此方案使用 Office 脚本从 [NOAA](https://tidesandcurrents.noaa.gov/) (当前数据库的外部源提取数据) 并绘制结果信息的图形。 它重点介绍了使用 `fetch` 获取数据和使用图表的技能。 |
| [成绩计算器](../scenarios/grade-calculator.md) | 此方案具有一个脚本，用于验证教师的课堂成绩记录。 它展示错误检查、单元格格式设置和正则表达式的技能。 |
| [任务提醒](../scenarios/task-reminders.md) | 此方案在 Power Automate 流中使用 Office 脚本向同事发送提醒，以更新项目状态。 它重点介绍了 Power Automate 集成以及与脚本之间传输数据的技能。 |

## <a name="community-contributions"></a>社区贡献

欢迎来自[](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) Office 脚本社区的贡献！ 可随意创建拉取请求进行审阅。

| Project | 详细信息 |
|---------|---------|
| [四年问候语动画](community-seasons-greetings.md) | 此脚本由 [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) 在假日假日的快乐中贡献！ 这是一个有趣脚本，它使用 Office 脚本在 Excel 网页中显示一个百叶树。 |

## <a name="try-it-out"></a>试用

这些示例是开源的。 尝试一下。 你需要从工作或学校获得 Microsoft 工作或学校帐户，并拥有使用 E3 或 (Microsoft 365 订阅) 。 只需前往登录 https://office.com 帐户并开始使用。

## <a name="leave-a-comment"></a>留下注释

使用特定示例的文档页面底部的"反馈"部分，可以随意留下评论、提出建议或记录问题。
