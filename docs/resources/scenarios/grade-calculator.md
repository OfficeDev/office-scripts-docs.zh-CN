---
title: Office脚本示例方案：成绩计算器
description: 确定一类学生成绩的百分比和字母成绩的示例。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 2d98e68f37418ade238a707cb74cc7ccf47e8f59
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313790"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="b1e42-103">Office脚本示例方案：成绩计算器</span><span class="sxs-lookup"><span data-stu-id="b1e42-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="b1e42-104">在此方案中，你是一名教师，对每位学生的学期结束成绩进行评分。</span><span class="sxs-lookup"><span data-stu-id="b1e42-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="b1e42-105">你一直在输入他们的工作分配和测试的分数。</span><span class="sxs-lookup"><span data-stu-id="b1e42-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="b1e42-106">现在，是时候确定学生了。</span><span class="sxs-lookup"><span data-stu-id="b1e42-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="b1e42-107">您将开发一个脚本，该脚本将针对每个分数类别对成绩进行总计。</span><span class="sxs-lookup"><span data-stu-id="b1e42-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="b1e42-108">然后，它将基于总数为每个学生分配一个信函成绩。</span><span class="sxs-lookup"><span data-stu-id="b1e42-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="b1e42-109">为了帮助确保准确性，你将添加一些检查，以查看个别分数是否太低或太高。</span><span class="sxs-lookup"><span data-stu-id="b1e42-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="b1e42-110">如果学生的分数小于零或大于可能的分数值，该脚本将用红色填充标记该单元格，而不是该学生的总分。</span><span class="sxs-lookup"><span data-stu-id="b1e42-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="b1e42-111">这将清楚地指示需要仔细检查哪些记录。</span><span class="sxs-lookup"><span data-stu-id="b1e42-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="b1e42-112">你还将向成绩添加一些基本格式，以便快速查看课程的顶部和底部。</span><span class="sxs-lookup"><span data-stu-id="b1e42-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="b1e42-113">涵盖的脚本编写技能</span><span class="sxs-lookup"><span data-stu-id="b1e42-113">Scripting skills covered</span></span>

- <span data-ttu-id="b1e42-114">单元格格式</span><span class="sxs-lookup"><span data-stu-id="b1e42-114">Cell formatting</span></span>
- <span data-ttu-id="b1e42-115">错误检查</span><span class="sxs-lookup"><span data-stu-id="b1e42-115">Error checking</span></span>
- <span data-ttu-id="b1e42-116">正则表达式</span><span class="sxs-lookup"><span data-stu-id="b1e42-116">Regular expressions</span></span>
- <span data-ttu-id="b1e42-117">条件格式</span><span class="sxs-lookup"><span data-stu-id="b1e42-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="b1e42-118">设置说明</span><span class="sxs-lookup"><span data-stu-id="b1e42-118">Setup instructions</span></span>

1. <span data-ttu-id="b1e42-119">将<a href="grade-calculator.xlsx">grade-calculator.xlsx</a>下载到OneDrive。</span><span class="sxs-lookup"><span data-stu-id="b1e42-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="b1e42-120">打开工作簿，Excel 网页版。</span><span class="sxs-lookup"><span data-stu-id="b1e42-120">Open the workbook with Excel for the web.</span></span>

1. <span data-ttu-id="b1e42-121">在" **自动化"** 选项卡下，选择 **"新建脚本** "，然后将以下脚本粘贴到编辑器中。</span><span class="sxs-lookup"><span data-stu-id="b1e42-121">Under the **Automate** tab, select **New Script** and paste the following script into the editor.</span></span>

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

1. <span data-ttu-id="b1e42-122">将脚本重命名 **为成绩计算器** 并保存它。</span><span class="sxs-lookup"><span data-stu-id="b1e42-122">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="b1e42-123">运行脚本</span><span class="sxs-lookup"><span data-stu-id="b1e42-123">Running the script</span></span>

<span data-ttu-id="b1e42-124">在唯 **一的工作表** 上运行成绩计算器脚本。</span><span class="sxs-lookup"><span data-stu-id="b1e42-124">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="b1e42-125">该脚本将总计成绩，并为每个学生分配一个信函成绩。</span><span class="sxs-lookup"><span data-stu-id="b1e42-125">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="b1e42-126">如果任何单个成绩的分数大于作业或测试的分数，则有问题的成绩将标记为红色，不计算总分。</span><span class="sxs-lookup"><span data-stu-id="b1e42-126">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span> <span data-ttu-id="b1e42-127">此外，任何"A"成绩都用绿色突出显示，而"D"和"F"成绩用黄色突出显示。</span><span class="sxs-lookup"><span data-stu-id="b1e42-127">Also, any 'A' grades are highlighted in green, while 'D' and 'F' grades are highlighted in yellow.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="b1e42-128">运行脚本之前</span><span class="sxs-lookup"><span data-stu-id="b1e42-128">Before running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="显示学生分数行的工作表。":::

### <a name="after-running-the-script"></a><span data-ttu-id="b1e42-130">运行脚本后</span><span class="sxs-lookup"><span data-stu-id="b1e42-130">After running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="一个工作表，用于显示有效学生行的学生分数数据，并包含红色总计中的无效单元格。":::
