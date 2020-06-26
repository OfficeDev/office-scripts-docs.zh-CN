---
title: Office 脚本示例方案：年级计算器
description: 一个用于确定一类学生的百分比和信函等级的示例。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 6f8e3db756c72cf1d0e2f774ccd819c041f0c42d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878638"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="928f6-103">Office 脚本示例方案：年级计算器</span><span class="sxs-lookup"><span data-stu-id="928f6-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="928f6-104">在这种情况下，您是每位学生的长期成绩的指导员计数。</span><span class="sxs-lookup"><span data-stu-id="928f6-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="928f6-105">你已为工作分配和测试输入分数。</span><span class="sxs-lookup"><span data-stu-id="928f6-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="928f6-106">现在，我们来确定学生的 fates。</span><span class="sxs-lookup"><span data-stu-id="928f6-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="928f6-107">您将开发一个用于汇总每个点类别的成绩的脚本。</span><span class="sxs-lookup"><span data-stu-id="928f6-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="928f6-108">然后，它将根据总数向每个学生分配一个信函等级。</span><span class="sxs-lookup"><span data-stu-id="928f6-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="928f6-109">为了帮助确保准确性，您将添加两个检查，以查看是否有任何单个分数太低或过高。</span><span class="sxs-lookup"><span data-stu-id="928f6-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="928f6-110">如果学生的分数小于零或大于可能的磅值，则该脚本将使用红色填充标记单元格，而不是学生的分数的总和。</span><span class="sxs-lookup"><span data-stu-id="928f6-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="928f6-111">这将明确指出需要进行双重检查的记录。</span><span class="sxs-lookup"><span data-stu-id="928f6-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="928f6-112">您还将向成绩添加一些基本格式，以便您可以快速查看课程的顶部和底部。</span><span class="sxs-lookup"><span data-stu-id="928f6-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="928f6-113">涵盖的脚本技能</span><span class="sxs-lookup"><span data-stu-id="928f6-113">Scripting skills covered</span></span>

- <span data-ttu-id="928f6-114">单元格格式</span><span class="sxs-lookup"><span data-stu-id="928f6-114">Cell formatting</span></span>
- <span data-ttu-id="928f6-115">错误检查</span><span class="sxs-lookup"><span data-stu-id="928f6-115">Error checking</span></span>
- <span data-ttu-id="928f6-116">正则表达式</span><span class="sxs-lookup"><span data-stu-id="928f6-116">Regular expressions</span></span>
- <span data-ttu-id="928f6-117">条件格式</span><span class="sxs-lookup"><span data-stu-id="928f6-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="928f6-118">设置说明</span><span class="sxs-lookup"><span data-stu-id="928f6-118">Setup instructions</span></span>

1. <span data-ttu-id="928f6-119">将<a href="grade-calculator.xlsx">grade-calculator.xlsx</a>下载到你的 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="928f6-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="928f6-120">使用适用于 web 的 Excel 打开工作簿。</span><span class="sxs-lookup"><span data-stu-id="928f6-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="928f6-121">在 "**自动化**" 选项卡上，打开**代码编辑器**。</span><span class="sxs-lookup"><span data-stu-id="928f6-121">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="928f6-122">在 "**代码编辑器**" 任务窗格中，按 "**新建脚本**"，并将以下脚本粘贴到编辑器中。</span><span class="sxs-lookup"><span data-stu-id="928f6-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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
      const assignmentMaxMatches = studentData[0][1].match(/\d+/);
      const midtermMaxMatches = studentData[0][2].match(/\d+/);
      const finalMaxMatches = studentData[0][3].match(/\d+/);

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
        const total = studentData[i][1] + studentData[i][2] + studentData[i][3];
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
          "#9C0006",
          "#FFC7CE",
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
      let conditionalFormatting : ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({formula1, operator});
    }
    ```

5. <span data-ttu-id="928f6-123">将脚本重命名为**评分计算器**并保存它。</span><span class="sxs-lookup"><span data-stu-id="928f6-123">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="928f6-124">运行脚本</span><span class="sxs-lookup"><span data-stu-id="928f6-124">Running the script</span></span>

<span data-ttu-id="928f6-125">在唯一的工作表上运行**年级计算器**脚本。</span><span class="sxs-lookup"><span data-stu-id="928f6-125">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="928f6-126">该脚本将对分数进行合计并为每个学生分配一个字母等级。</span><span class="sxs-lookup"><span data-stu-id="928f6-126">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="928f6-127">如果任何一年级的分数多于工作分配或测试的数量，则会将有问题的等级标记为红色，并且不计算总计。</span><span class="sxs-lookup"><span data-stu-id="928f6-127">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="928f6-128">运行脚本之前</span><span class="sxs-lookup"><span data-stu-id="928f6-128">Before running the script</span></span>

![显示学生的分数行的工作表。](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="928f6-130">运行脚本后</span><span class="sxs-lookup"><span data-stu-id="928f6-130">After running the script</span></span>

![在有效学生行中显示带有红色总计的无效单元格的学生分数数据的工作表。](../../images/scenario-grade-calculator-after.png)
