---
title: Office 脚本示例方案：年级计算器
description: 一个用于确定一类学生的百分比和信函等级的示例。
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 0db6f7c116594f7655bfc0adc8f5a79dbbf2a0af
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700129"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Office 脚本示例方案：年级计算器

在这种情况下，您是每位学生的长期成绩的指导员计数。 你已为工作分配和测试输入分数。 现在，我们来确定学生的 fates。

您将开发一个用于汇总每个点类别的成绩的脚本。 然后，它将根据总数向每个学生分配一个信函等级。 为了帮助确保准确性，您将添加两个检查，以查看是否有任何单个分数太低或过高。 如果学生的分数小于零或大于可能的磅值，则该脚本将使用红色填充标记单元格，而不是学生的分数的总和。 这将明确指出需要进行双重检查的记录。 您还将向成绩添加一些基本格式，以便您可以快速查看课程的顶部和底部。

## <a name="scripting-skills-covered"></a>涵盖的脚本技能

- 单元格格式
- 错误检查
- 正则表达式

## <a name="setup-instructions"></a>设置说明

1. 将<a href="grade-calculator.xlsx">grade-calculator</a>下载到你的 OneDrive。

2. 使用适用于 web 的 Excel 打开工作簿。

3. 在 "**自动化**" 选项卡上，打开**代码编辑器**。

4. 在 "**代码编辑器**" 任务窗格中，按 "**新建脚本**"，并将以下脚本粘贴到编辑器中。

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the number of student record rows.
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let studentsRange = sheet.getUsedRange().load("values, rowCount");
      await context.sync();
      console.log("Total students: " + (studentsRange.rowCount - 1));

      // Clean up any formatting from previous runs of the script.
      studentsRange.clear(Excel.ClearApplyTo.formats);
      studentsRange.getColumn(4).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      await context.sync();

      // Parse the headers for the maximum possible scores for each category.
      // The format is `category (score)`.
      let assignmentsMax = studentsRange.values[0][1].match(/\d+/)[0];
      let midTermMax = studentsRange.values[0][2].match(/\d+/)[0];
      let finalsMax = studentsRange.values[0][3].match(/\d+/)[0];
      console.log("Assignments max score:" + assignmentsMax);
      console.log("Mid-term max score: " + midTermMax);
      console.log("Final max score: " + finalsMax);

      // Look at every student row.
      for (let i = 1; i < studentsRange.values.length; i++) {
        let row = studentsRange.values[i];
        let total = row[1] + row[2] + row[3];
        let valid = true;

        // Look for any records that are too low or too high.
        if (row[1] < 0 || row[1] > assignmentsMax) {
          studentsRange.getCell(i, 1).format.fill.color = "Red";
          valid = false;
        }
        if (row[2] < 0 || row[2] > midTermMax) {
          studentsRange.getCell(i, 2).format.fill.color = "Red";
          valid = false;
        }
        if (row[3] < 0 || row[3] > finalsMax) {
          studentsRange.getCell(i, 3).format.fill.color = "Red";
          valid = false;
        }

        // If the scores are valid, total that student's points and assign them a letter grade.
        if (valid) {
          let grade: string;
          switch (true) {
            case total < 60:
              grade = "E";
              break;
            case total < 70:
              grade = "D";
              break;
            case total < 80:
              grade = "C";
              break;
            case total < 90:
              grade = "B";
              break;
            default:
              grade = "A";
              break;
          }

          studentsRange.getCell(i, 4).values = [[total]];
          studentsRange.getCell(i, 5).values = [[grade]];

          // Highlight excellent students and those in need of attention.
          if (grade === "A") {
            studentsRange.getCell(i, 5).format.fill.color = "Green";
          } else if (grade === "E" || grade === "D") {
            studentsRange.getCell(i, 5).format.fill.color = "Orange";
          }
        }
      }

      studentsRange.getColumn(5).format.horizontalAlignment = "Center";
    }
    ```

5. 将脚本重命名为**评分计算器**并保存它。

## <a name="running-the-script"></a>运行脚本

在唯一的工作表上运行**年级计算器**脚本。 该脚本将对分数进行合计并为每个学生分配一个字母等级。 如果任何一年级的分数多于工作分配或测试的数量，则会将有问题的等级标记为红色，并且不计算总计。

### <a name="before-running-the-script"></a>运行脚本之前

![显示学生的分数行的工作表。](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a>运行脚本后

![在有效学生行中显示带有红色总计的无效单元格的学生分数数据的工作表。](../../images/scenario-grade-calculator-after.png)
