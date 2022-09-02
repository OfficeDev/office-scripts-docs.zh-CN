---
title: Office 脚本示例
description: 可用的 Office 脚本示例和方案。
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5798da37bd4166d18b41c005c4d8cc8a4b6c401d
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572484"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office 脚本示例和方案

本部分包含基于 [Office 脚本的](../../overview/excel.md) 自动化解决方案，可帮助最终用户实现日常任务的自动化。 它包含业务用户面临的现实方案，并提供详细的解决方案以及分步说明性视频链接。

有关 [基础](#basics) 知识和 [基础知识之外](#beyond-the-basics)的每个项目，请查看源代码、分步 [**YouTube 视频**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)等。

在 [方案中](#scenarios)，我们包含了一些更大的方案示例，这些示例演示了真实的用例。

我们也欢迎 [社区的贡献](#community-contributions-and-fun-samples)。 这些示例开放源代码。

> [!IMPORTANT]
> 在尝试示例之前，请确保满足 Office 脚本的先决条件。 Microsoft 365 订阅和帐户的要求位于 [Office Scripts for Excel 概述“要求”部分](../../overview/excel.md#requirements)下。

## <a name="basics"></a>基本信息

| Microsoft Project | 详细信息 |
|---------|---------|
| [脚本基础知识](excel-samples.md) | 这些示例演示了 Office 脚本的基本构建基块。 |
| [在 Excel 中添加注释](add-excel-comments.md) | 此示例将注释添加到包含@mentioning同事的单元格。 |
| [向工作簿添加图像](add-image-to-workbook.md) | 此示例将图像添加到工作簿，并跨工作表复制图像。|
| [将多个 Excel 表复制到单个表中](copy-tables-combine.md) | 此示例将多个 Excel 表中的数据合并到包含所有行的单个表中。 |
| [创建工作簿目录](table-of-contents.md) | 此示例创建一个包含指向每个工作表链接的内容表。 |
| [删除表格列筛选器](clear-table-filter-for-active-cell.md) | 此示例清除表列中的所有筛选器。 |
| [在 Excel 中记录日常更改，并使用 Power Automate 流报告这些更改](report-day-to-day-changes.md) | 此示例使用计划的 Power Automate 流来记录每日读数并报告更改。 |

## <a name="beyond-the-basics"></a>超越基础设置

请查看以下端到端项目，该项目可自动执行示例方案以及完整的脚本、使用的示例 Excel 文件以及 [托管在 YouTube) 上的视频 (](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)。

| 项目 | 详细信息 |
|---------|---------|
| [将工作表合并到单个工作簿中](combine-worksheets-into-single-workbook.md) | 此示例使用 Office 脚本和 Power Automate 将数据从其他工作簿拉取到单个工作簿中。 |
| [将 CSV 文件转换为 Excel 工作簿](convert-csv.md) | 此示例使用 Office 脚本和 Power Automate 从.csv文件创建.xlsx文件。 |
| [跨引用工作簿](excel-cross-reference.md) | 此示例使用 Office 脚本和 Power Automate 交叉引用和验证不同工作簿中的信息。 |
| [对特定工作表或所有工作表中的空白行进行计数](count-blank-rows.md) | 此示例检测工作表中是否有任何空白行，你预计数据存在，然后报告 Power Automate 流中的使用情况的空白行计数。 |
| [Email图表和表格图像](email-images-chart-table.md) | 此示例使用 Office 脚本和 Power Automate 操作创建图表，并通过电子邮件将该图表作为图像发送。 |
| [外部提取调用](external-fetch-calls.md) | 此示例用于 `fetch` 从 GitHub 获取脚本的信息。 |
| [在 Excel 中管理计算模式](excel-calculation.md) | 此示例演示如何使用 Office 脚本在Excel web 版中使用计算模式和计算方法。 |
| [跨表移动行](move-rows-across-tables.md) | 此示例演示如何通过保存筛选器，然后处理和重新应用筛选器来跨表移动行。 |
| [将 Excel 数据输出为 JSON](get-table-data.md) | 此解决方案演示如何将 Excel 表数据输出为要在 Power Automate 中使用的 JSON。 |
| [从 Excel 工作表中的每个单元格中删除超链接](remove-hyperlinks-from-cells.md) | 此示例清除当前工作表中的所有超链接。 |
| [对文件夹中的所有 Excel 文件运行脚本](automate-tasks-on-all-excel-files-in-folder.md) | 此项目对位于文件夹中的所有文件执行一组自动化任务OneDrive for Business (也可用于 SharePoint 文件夹) 。 它对 Excel 文件执行计算，添加格式，并插入@mentions同事的注释。 |
| [编写大型数据集](write-large-dataset.md) | 此示例演示如何将大范围作为较小的子范围发送。 |

## <a name="scenarios"></a>应用场景

Office 脚本可以自动执行日常程序的各个部分。 这些日常任务通常存在于独特的生态系统中，Excel 工作簿是以特定方式设置的。 这些较大的方案示例演示了这种真实的用例。 它们包括 Office 脚本和工作簿，因此你可以从端到端查看方案。

| 应用场景 | 详细信息 |
|---------|---------|
| [分析 Web 下载项](../scenarios/analyze-web-downloads.md) | 此方案具有一个脚本，用于分析 Web 流量记录以确定用户的原产国。 它展示了文本分析、在脚本中使用子功能、应用条件格式和处理表的技能。 |
| [从 NOAA 中提取图形水级别的数据](../scenarios/noaa-data-fetch.md) | 此方案使用 Office 脚本从外部源提取数据， ([NOAA Tides 和 Currents 数据库](https://tidesandcurrents.noaa.gov/)) 并绘制生成的信息图。 它突出显示了用于 `fetch` 获取数据和使用图表的技能。 |
| [成绩计算器](../scenarios/grade-calculator.md) | 此方案具有一个脚本，用于验证讲师的班级成绩记录。 它展示了错误检查、单元格格式和正则表达式的技能。 |
| [在 Teams 中安排面试](../scenarios/schedule-interviews-in-teams.md) | 此方案演示如何使用 Excel 电子表格来管理面试会议时间，以及如何在 Teams 中安排会议。 |
| [任务提醒](../scenarios/task-reminders.md) | 此方案使用 Power Automate 流中的 Office 脚本向同事发送提醒以更新项目的状态。 它突出显示了 Power Automate 集成和脚本数据传输的技能。 |

## <a name="community-contributions-and-fun-samples"></a>社区贡献和有趣的示例

我们欢迎来自 Office 脚本社区的 [捐款](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) ！ 可以随意创建拉取请求以供审阅。

| 项目 | 详细信息 |
|---------|---------|
| [生活游戏](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | 黄玉涛在 Excel 技术社区的“就绪玩家零”博客包括一个脚本，用于为约翰·康威的 [*《生活游戏》*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life)建模。 |
| [打卡时钟按钮](../scenarios/punch-clock.md) | 这个剧本是由 [布赖恩·冈萨雷斯](https://github.com/b-gonzalez)贡献的。 该方案具有一个脚本和一个脚本按钮，用于记录当前时间。 |
| [季节问候动画](community-seasons-greetings.md) | 这个剧本是由 [莱斯利·布莱克](https://www.linkedin.com/in/lesblackconsultant/) 本着节日的精神贡献的！ 这是一个有趣的脚本，显示一个唱歌的圣诞树在Excel web 版使用 Office 脚本。 |

## <a name="leave-a-comment"></a>保留注释

可以使用特定示例文档页面底部的反 **馈** 部分随意留下批注、提出建议或记录问题。
