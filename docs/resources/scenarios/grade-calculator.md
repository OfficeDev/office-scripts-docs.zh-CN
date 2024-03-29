---
title: Office 脚本示例方案：等级计算器
description: 一个示例，用于确定学生班的百分比和字母成绩。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7dda3ebe84dc3edd10998cbe2c4cd0806da11411
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572526"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Office 脚本示例方案：等级计算器

在此方案中，你是一名教师，用于计算每个学生的学期末成绩。 你一直在输入他们的作业和测试的分数，你去。 现在，是时候确定学生的命运了。

你将开发一个脚本，该脚本将每个分数类别的分数总计。 然后，它将根据总数为每个学生分配一个字母成绩。 为了帮助确保准确性，你将添加一些检查，以查看任何单个分数是否太低或过高。 如果学生的分数小于零或大于可能的点值，则脚本将用红色填充标记单元格，而不是该学生的分数总数。 这将清楚地指示需要仔细检查哪些记录。 你还将向成绩添加一些基本格式，以便快速查看类的顶部和底部。

## <a name="scripting-skills-covered"></a>所涵盖的脚本技能

- 单元格格式
- 错误检查
- 正则表达式
- 条件格式

## <a name="setup-instructions"></a>设置说明

1. 将 [grade-calculator.xlsx](grade-calculator.xlsx) 下载到 OneDrive。

1. 使用Excel 网页版打开工作簿。

1. 在 **“自动执行”** 选项卡下，选择 **“新建脚本** ”并将以下脚本粘贴到编辑器中。

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the worksheet and validate the data.
      let studentsRange = workbook.getActiveWorksheet().getUsedRange();
      if (studentsRange.getColumnCount() !== 6) {
        throw new Error(`The required columns are not present. Expected column headers: "Student ID | Assignment score | Mid-term | Final | Total | Grade"`);
      }

      let studentData = studentsRange.getValues();

      // Clear the total and grade columns.
      studentsRange.getColumn(4).getCell(1, 0).getAbsoluteResizedRange(studentData.length - 1, 2).clear();

      // Clear all conditional formatting.
      workbook.getActiveWorksheet().getUsedRange().clearAllConditionalFormats();

      // Use regular expressions to read the max score from the assignment, mid-term, and final scores columns.
      let maxScores: string[] = [];
      const assignmentMaxMatches = (studentData[0][1] as string).match(/\d+/);
      const midtermMaxMatches = (studentData[0][2] as string).match(/\d+/);
      const finalMaxMatches = (studentData[0][3] as string).match(/\d+/);

      // Check the matches happened before proceeding.
      if (!(assignmentMaxMatches && midtermMaxMatches && finalMaxMatches)) {
        throw new Error(`The scores are not present in the column headers. Expected format: "Assignments (n)|Mid-term (n)|Final (n)"`);
      }

      // Use the first (and only) match from the regular expressions as the max scores.
      maxScores = [assignmentMaxMatches[0], midtermMaxMatches[0], finalMaxMatches[0]];

      // Set conditional formatting for each of the assignment, mid-term, and final scores columns.
      maxScores.forEach((score, i) => {
        let range = studentsRange.getColumn(i + 1).getCell(0, 0).getRowsBelow(studentData.length - 1);
        setCellValueConditionalFormatting(
          score,
          range,
          "#9C0006",
          "#FFC7CE",
          ExcelScript.ConditionalCellValueOperator.greaterThan
        )
      });

      // Store the current range information to avoid calling the workbook in the loop.
      let studentsRangeFormulas = studentsRange.getColumn(4).getFormulasR1C1();
      let studentsRangeValues = studentsRange.getColumn(5).getValues();

      /* Iterate over each of the student rows and compute the total score and letter grade.
      * Note that iterator starts at index 1 to skip first (header) row.
      */
      for (let i = 1; i < studentData.length; i++) {
        // If any of the scores are invalid, skip processing it.
        if (studentData[i][1] > maxScores[0] ||
          studentData[i][2] > maxScores[1] ||
          studentData[i][3] > maxScores[2]) {
          continue;
        }
        const total = (studentData[i][1] as number) + (studentData[i][2] as number) + (studentData[i][3] as number);
        let grade: string;
        switch (true) {
          case total < 60:
            grade = "F";
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
    
        // Set total score formula.
        studentsRangeFormulas[i][0] = '=RC[-2]+RC[-1]';
        // Set grade cell.
        studentsRangeValues[i][0] = grade;
      }

      // Set the formulas and values outside the loop.
      studentsRange.getColumn(4).setFormulasR1C1(studentsRangeFormulas);
      studentsRange.getColumn(5).setValues(studentsRangeValues);

      // Put a conditional formatting on the grade column.
      let totalRange = studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentData.length - 1);
      setCellValueConditionalFormatting(
        "A",
        totalRange,
        "#001600",
        "#C6EFCE",
        ExcelScript.ConditionalCellValueOperator.equalTo
      );
      ["D", "F"].forEach((grade) => {
        setCellValueConditionalFormatting(
          grade,
          totalRange,
          "#443300",
          "#FFEE22",
          ExcelScript.ConditionalCellValueOperator.equalTo
        );
      })
      // Center the grade column.
      studentsRange.getColumn(5).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    }

    /**
     * Helper function to apply conditional formatting.
     * @param value Cell value to use in conditional formatting formula1.
     * @param range Target range.
     * @param fontColor Font color to use.
     * @param fillColor Fill color to use.
     * @param operator Operator to use in conditional formatting.
     */
    function setCellValueConditionalFormatting(
      value: string,
      range: ExcelScript.Range,
      fontColor: string,
      fillColor: string,
      operator: ExcelScript.ConditionalCellValueOperator) {
      // Determine the formula1 based on the type of value parameter.
      let formula1: string;
      if (isNaN(Number(value))) {
        // For cell value equalTo rule, use this format: formula1: "=\"A\"",
        formula1 = `=\"${value}\"`;
      } else {
        // For number input (greater-than or less-than rules), just append '='.
        formula1 = `=${value}`;
      }

      // Apply conditional formatting.
      let conditionalFormatting: ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({ formula1, operator });
    }
    ```

1. 将脚本重命名为 **“等级计算器”** 并保存它。

## <a name="running-the-script"></a>运行脚本

在唯一的工作表上运行 **“等级计算器** ”脚本。 该脚本将汇总成绩，并为每个学生分配一个字母等级。 如果任何单个成绩的分数超过作业或测试值，则将违规成绩标记为红色，并且不计算总分。 此外，任何“A”成绩都以绿色突出显示，而“D”和“F”成绩以黄色突出显示。

### <a name="before-running-the-script"></a>运行脚本之前

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="显示学生分数行的工作表。":::

### <a name="after-running-the-script"></a>运行脚本后

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="显示有效学生行的无效单元格为红色总计的学生评分数据的工作表。":::
