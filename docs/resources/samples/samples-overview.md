---
title: Office脚本示例
description: 可用于Office脚本示例和方案。
ms.date: 01/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: ca8ed15983c2171c2e9eb2291cc78d7e4d536ac8
ms.sourcegitcommit: 161229492c85f3519c899573cf5022140026e7b8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/26/2022
ms.locfileid: "62220405"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office脚本示例和方案

本节包含[Office](../../overview/excel.md)脚本的自动化解决方案，帮助最终用户实现日常任务的自动化。 它包含业务用户面临的实际方案，并提供详细的解决方案以及分步说明视频链接。

对于 Basics 和 [Beyond the basics](#beyond-the-basics)中的每个项目，请查看源代码、分步 [**YouTube 视频**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)等。 [](#basics)

在 ["应用](#scenarios)场景"中，我们包含了几个演示实际用例的较大方案示例。

我们还欢迎 [来自社区的贡献](#community-contributions-and-fun-samples)。

## <a name="basics"></a>基本信息

| 项目 | 详细信息 |
|---------|---------|
| [脚本基础知识](../excel-samples.md) | 这些示例演示了脚本的基本Office构建基块。 |
| [在外接程序中添加Excel](add-excel-comments.md) | 本示例向单元格添加注释，@mentioning同事。 |
| [向工作簿添加图像](add-image-to-workbook.md) | 本示例向工作簿添加一个图像，并跨工作表复制一个图像。|
| [将多个Excel表复制到单个表中](copy-tables-combine.md) | 此示例将来自多个Excel的数据组合到一个包含所有行的表中。 |
| [创建工作簿目录](table-of-contents.md) | 此示例创建包含指向每个工作表的链接的目录。 |

## <a name="beyond-the-basics"></a>超越基础设置

请查看以下端到端项目，该项目可自动执行示例方案以及[YouTube ](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)) 上承载的完整脚本、Excel 文件示例 (文件。

| Project | 详细信息 |
|---------|---------|
| [将工作表合并到单个工作簿中](combine-worksheets-into-single-workbook.md) | 此示例使用Office脚本Power Automate将其他工作簿的数据拉入单个工作簿。 |
| [将 CSV 文件转换为Excel工作簿](convert-csv.md) | 此示例使用Office脚本Power Automate从.xlsx创建.csv文件。 |
| [交叉引用工作簿](excel-cross-reference.md) | 此示例使用Office脚本Power Automate来交叉引用和验证不同工作簿中的信息。 |
| [计算特定工作表或所有工作表中的空行数](count-blank-rows.md) | 本示例检测工作表中是否有预计存在数据的空白行，然后报告空白行计数，以用于Power Automate流。 |
| [电子邮件图表和表格图像](email-images-chart-table.md) | 此示例使用Office脚本Power Automate操作来创建图表，并通过电子邮件将图表作为图像发送。 |
| [外部提取调用](external-fetch-calls.md) | 此示例使用 `fetch` 从脚本的 GitHub获取信息。 |
| [筛选Excel表并获取可见区域](filter-table-get-visible-range.md) | 此示例筛选一Excel，并返回可见区域作为 JSON 对象。 此 JSON 可以作为较大解决方案的一Power Automate流提供给一个流。 |
| [在计算模式下管理Excel](excel-calculation.md) | 此示例演示如何在脚本中使用计算模式和Excel web 版Office方法。 |
| [跨表移动行](move-rows-across-tables.md) | 此示例演示如何通过保存筛选器，然后处理和重新应用筛选器来跨表移动行。 |
| [输出Excel JSON](get-table-data.md) | 此解决方案演示如何将Excel数据输出为 JSON，以用于Power Automate。 |
| [从工作表的每个单元格中删除Excel超链接](remove-hyperlinks-from-cells.md) | 本示例清除当前工作表的所有超链接。 |
| [对文件夹中的所有 Excel 文件运行脚本](automate-tasks-on-all-excel-files-in-folder.md) | 此项目对位于 OneDrive for Business (上的文件夹中的所有文件执行一组自动化任务，SharePoint文件夹) 。 该代码对Excel执行计算，添加格式，并插入一个@mentions注释。 |
| [编写大型数据集](write-large-dataset.md) | 此示例演示如何将较大区域作为较小的子范围发送。 |

## <a name="scenarios"></a>应用场景

Office脚本可以自动执行日常例程的某些部分。 这些日常任务通常存在于独特的生态系统中，其中Excel以特定方式设置的特定工作簿。 这些较大的方案示例演示了此类实际用例。 它们包括Office脚本和工作簿，因此您可以端到端查看方案。

| 应用场景 | 详细信息 |
|---------|---------|
| [分析 Web 下载项](../scenarios/analyze-web-downloads.md) | 此方案具有分析 Web 流量记录的脚本，以确定用户的来源国家/地区。 它展示文本分析、在脚本中使用子功能、应用条件格式和使用表的技能。 |
| [从 NOAA 中提取图形水级别的数据](../scenarios/noaa-data-fetch.md) | 此方案使用Office脚本从[NOAA](https://tidesandcurrents.noaa.gov/) (当前数据库的外部源提取数据) 并绘制结果信息的图形。 它重点介绍了使用 `fetch` 获取数据和使用图表的技能。 |
| [成绩计算器](../scenarios/grade-calculator.md) | 此方案具有一个脚本，用于验证教师的课堂成绩记录。 它展示错误检查、单元格格式设置和正则表达式的技能。 |
| [在 Teams 中安排面试](../scenarios/schedule-interviews-in-teams.md) | 此方案演示如何使用 Excel 电子表格管理访谈式会议时间，并创建一个流来安排Teams。 |
| [任务提醒](../scenarios/task-reminders.md) | 此方案在Office流Power Automate脚本，向同事发送更新项目状态的提醒。 它重点介绍了Power Automate和从脚本传输数据的技术。 |

## <a name="community-contributions-and-fun-samples"></a>Community贡献和有趣的示例

欢迎来自[我们的](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md)脚本Office贡献！ 可随意创建拉取请求进行审阅。

| 项目 | 详细信息 |
|---------|---------|
| [游戏生活](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | Tech Community 上的 Yutao 为"Ready Player Zero"（准备玩家零）的Excel包括一个脚本，用于为 John Conway 的《生活游戏》[*建模。*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life) |
| [四年问候语动画](community-seasons-greetings.md) | 此脚本由 [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) 在假日假日的快乐中贡献！ 这是一个有趣脚本，它使用脚本在Excel web 版中Office一个树。 |

## <a name="try-it-out"></a>试用

这些示例是开源的。 尝试一下。 你需要从工作或学校获得 Microsoft 工作或学校帐户，并拥有许可证才能Microsoft 365 E3 (或) 。 只需前往登录 https://office.com 帐户并开始使用。

## <a name="leave-a-comment"></a>留下注释

使用特定示例的文档页面底部的"反馈"部分，可以随意留下评论、提出建议或记录问题。
