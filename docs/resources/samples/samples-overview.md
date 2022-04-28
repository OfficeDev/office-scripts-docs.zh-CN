---
title: Office脚本示例
description: 可用Office脚本示例和方案。
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c9bbe9b6f7eb8abad2995dac72ccf636d585d69
ms.sourcegitcommit: e6428a5214fa38aef036a952a0e3c09dbf6e4d3e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/28/2022
ms.locfileid: "65109153"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office脚本示例和方案

本部分包含[基于脚本的Office](../../overview/excel.md)自动化解决方案，可帮助最终用户实现日常任务的自动化。 它包含业务用户面临的现实方案，并提供详细的解决方案以及分步说明性视频链接。

有关 [基础](#basics) 知识和 [基础知识之外](#beyond-the-basics)的每个项目，请查看源代码、分步 [**YouTube 视频**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)等。

在 [方案中](#scenarios)，我们包含了一些更大的方案示例，这些示例演示了真实的用例。

我们也欢迎 [社区的贡献](#community-contributions-and-fun-samples)。

## <a name="basics"></a>基本信息

| Project | 详细信息 |
|---------|---------|
| [脚本基础知识](../excel-samples.md) | 这些示例演示了Office脚本的基本构建基块。 |
| [在Excel中添加注释](add-excel-comments.md) | 此示例将注释添加到包含@mentioning同事的单元格。 |
| [向工作簿添加图像](add-image-to-workbook.md) | 此示例将图像添加到工作簿，并跨工作表复制图像。|
| [将多个Excel表复制到单个表中](copy-tables-combine.md) | 此示例将多个Excel表中的数据合并到包含所有行的单个表中。 |
| [创建工作簿目录](table-of-contents.md) | 此示例创建一个包含指向每个工作表链接的内容表。 |

## <a name="beyond-the-basics"></a>超越基础设置

请查看以下端到端项目，该项目可自动执行示例方案以及完整的脚本、所使用的示例Excel文件以及[托管在 YouTube) 上的视频 (](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)。

| Project | 详细信息 |
|---------|---------|
| [将工作表合并到单个工作簿中](combine-worksheets-into-single-workbook.md) | 此示例使用Office脚本和Power Automate将数据从其他工作簿拉取到单个工作簿中。 |
| [将 CSV 文件转换为Excel工作簿](convert-csv.md) | 此示例使用Office脚本和Power Automate从.csv文件创建.xlsx文件。 |
| [跨引用工作簿](excel-cross-reference.md) | 此示例使用Office脚本和Power Automate来交叉引用和验证不同工作簿中的信息。 |
| [对特定工作表或所有工作表中的空白行进行计数](count-blank-rows.md) | 此示例检测工作表中是否有任何空白行，你预计存在数据，然后报告Power Automate流中的使用情况的空白行计数。 |
| [电子邮件图表和表格图像](email-images-chart-table.md) | 此示例使用Office脚本和Power Automate操作创建图表，并通过电子邮件将该图表作为图像发送。 |
| [外部提取调用](external-fetch-calls.md) | 此示例用于`fetch`从脚本的GitHub获取信息。 |
| [筛选Excel表并获取可见范围](filter-table-get-visible-range.md) | 此示例筛选Excel表，并将可见范围作为 JSON 对象返回。 此 JSON 可作为更大解决方案的一部分提供给Power Automate流。 |
| [在Excel中管理计算模式](excel-calculation.md) | 此示例演示如何使用Office脚本在Excel web 版中使用计算模式和计算方法。 |
| [跨表移动行](move-rows-across-tables.md) | 此示例演示如何通过保存筛选器，然后处理和重新应用筛选器来跨表移动行。 |
| [将数据输出Excel为 JSON](get-table-data.md) | 此解决方案演示如何将Excel表数据输出为要在Power Automate中使用的 JSON。 |
| [从Excel工作表中的每个单元格中删除超链接](remove-hyperlinks-from-cells.md) | 此示例清除当前工作表中的所有超链接。 |
| [对文件夹中的所有 Excel 文件运行脚本](automate-tasks-on-all-excel-files-in-folder.md) | 此项目对位于文件夹中的所有文件执行一组自动化任务OneDrive for Business (也可用于SharePoint文件夹) 。 它对Excel文件执行计算，添加格式，并插入@mentions同事的注释。 |
| [编写大型数据集](write-large-dataset.md) | 此示例演示如何将大范围作为较小的子范围发送。 |

## <a name="scenarios"></a>应用场景

Office脚本可以自动执行日常程序的各个部分。 这些日常任务通常存在于独特的生态系统中，其中Excel工作簿是以特定方式设置的。 这些较大的方案示例演示了这种真实的用例。 它们包括Office脚本和工作簿，因此你可以从端到端查看方案。

| 应用场景 | 详细信息 |
|---------|---------|
| [分析 Web 下载项](../scenarios/analyze-web-downloads.md) | 此方案具有一个脚本，用于分析 Web 流量记录以确定用户的原产国。 它展示了文本分析、在脚本中使用子功能、应用条件格式和处理表的技能。 |
| [从 NOAA 中提取图形水级别的数据](../scenarios/noaa-data-fetch.md) | 此方案使用Office脚本从外部源提取数据， ([NOAA Tides 和 Currents 数据库](https://tidesandcurrents.noaa.gov/)) 并绘制生成的信息图。 它突出显示了用于 `fetch` 获取数据和使用图表的技能。 |
| [成绩计算器](../scenarios/grade-calculator.md) | 此方案具有一个脚本，用于验证讲师的班级成绩记录。 它展示了错误检查、单元格格式和正则表达式的技能。 |
| [在 Teams 中安排面试](../scenarios/schedule-interviews-in-teams.md) | 此方案演示如何使用Excel电子表格来管理面试会议时间，以及如何在Teams中安排会议。 |
| [任务提醒](../scenarios/task-reminders.md) | 此方案使用Power Automate流中的Office脚本向同事发送提醒以更新项目的状态。 它突出显示了Power Automate集成和数据传输到脚本和从脚本传输数据的技能。 |

## <a name="community-contributions-and-fun-samples"></a>Community贡献和有趣的示例

我们欢迎来自Office脚本社区[的贡献](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md)！ 可以随意创建拉取请求以供审阅。

| Project | 详细信息 |
|---------|---------|
| [生活游戏](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | 黄玉涛在Excel科技Community上的“准备玩家零”博客包括一个脚本，为约翰·康威的 [*《生活游戏》*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life)建模。 |
| [“打孔时钟”按钮](../scenarios/punch-clock.md) | 这个剧本是由 [布赖恩·冈萨雷斯](https://github.com/b-gonzalez)贡献的。 该方案具有一个脚本和一个脚本按钮，用于记录当前时间。 |
| [季节问候动画](community-seasons-greetings.md) | 这个剧本是由 [莱斯利·布莱克](https://www.linkedin.com/in/lesblackconsultant/) 本着节日的精神贡献的！ 这是一个有趣的脚本，显示一个唱歌的圣诞树在Excel web 版使用Office脚本。 |

## <a name="try-it-out"></a>试用

这些示例开放源代码。 亲自试用它们。 你将需要一个 Microsoft 工作或学校帐户，从工作或学校具有Microsoft 365订阅 (E3 或更高) 的许可证。 只需转到一 https://office.com 下即可登录到帐户并开始使用。

## <a name="leave-a-comment"></a>保留注释

可以使用特定示例文档页面底部的反 **馈** 部分随意留下批注、提出建议或记录问题。
