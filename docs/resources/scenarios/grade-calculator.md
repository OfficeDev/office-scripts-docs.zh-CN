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
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="96eb7-103">Office 脚本示例方案：年级计算器</span><span class="sxs-lookup"><span data-stu-id="96eb7-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="96eb7-104">在这种情况下，您是每位学生的长期成绩的指导员计数。</span><span class="sxs-lookup"><span data-stu-id="96eb7-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="96eb7-105">你已为工作分配和测试输入分数。</span><span class="sxs-lookup"><span data-stu-id="96eb7-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="96eb7-106">现在，我们来确定学生的 fates。</span><span class="sxs-lookup"><span data-stu-id="96eb7-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="96eb7-107">您将开发一个用于汇总每个点类别的成绩的脚本。</span><span class="sxs-lookup"><span data-stu-id="96eb7-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="96eb7-108">然后，它将根据总数向每个学生分配一个信函等级。</span><span class="sxs-lookup"><span data-stu-id="96eb7-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="96eb7-109">为了帮助确保准确性，您将添加两个检查，以查看是否有任何单个分数太低或过高。</span><span class="sxs-lookup"><span data-stu-id="96eb7-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="96eb7-110">如果学生的分数小于零或大于可能的磅值，则该脚本将使用红色填充标记单元格，而不是学生的分数的总和。</span><span class="sxs-lookup"><span data-stu-id="96eb7-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="96eb7-111">这将明确指出需要进行双重检查的记录。</span><span class="sxs-lookup"><span data-stu-id="96eb7-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="96eb7-112">您还将向成绩添加一些基本格式，以便您可以快速查看课程的顶部和底部。</span><span class="sxs-lookup"><span data-stu-id="96eb7-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="96eb7-113">涵盖的脚本技能</span><span class="sxs-lookup"><span data-stu-id="96eb7-113">Scripting skills covered</span></span>

- <span data-ttu-id="96eb7-114">单元格格式</span><span class="sxs-lookup"><span data-stu-id="96eb7-114">Cell formatting</span></span>
- <span data-ttu-id="96eb7-115">错误检查</span><span class="sxs-lookup"><span data-stu-id="96eb7-115">Error checking</span></span>
- <span data-ttu-id="96eb7-116">正则表达式</span><span class="sxs-lookup"><span data-stu-id="96eb7-116">Regular expressions</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="96eb7-117">设置说明</span><span class="sxs-lookup"><span data-stu-id="96eb7-117">Setup instructions</span></span>

1. <span data-ttu-id="96eb7-118">将<a href="grade-calculator.xlsx">grade-calculator</a>下载到你的 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="96eb7-118">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="96eb7-119">使用适用于 web 的 Excel 打开工作簿。</span><span class="sxs-lookup"><span data-stu-id="96eb7-119">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="96eb7-120">在 "**自动化**" 选项卡上，打开**代码编辑器**。</span><span class="sxs-lookup"><span data-stu-id="96eb7-120">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="96eb7-121">在 "**代码编辑器**" 任务窗格中，按 "**新建脚本**"，并将以下脚本粘贴到编辑器中。</span><span class="sxs-lookup"><span data-stu-id="96eb7-121">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="96eb7-122">将脚本重命名为**评分计算器**并保存它。</span><span class="sxs-lookup"><span data-stu-id="96eb7-122">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="96eb7-123">运行脚本</span><span class="sxs-lookup"><span data-stu-id="96eb7-123">Running the script</span></span>

<span data-ttu-id="96eb7-124">在唯一的工作表上运行**年级计算器**脚本。</span><span class="sxs-lookup"><span data-stu-id="96eb7-124">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="96eb7-125">该脚本将对分数进行合计并为每个学生分配一个字母等级。</span><span class="sxs-lookup"><span data-stu-id="96eb7-125">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="96eb7-126">如果任何一年级的分数多于工作分配或测试的数量，则会将有问题的等级标记为红色，并且不计算总计。</span><span class="sxs-lookup"><span data-stu-id="96eb7-126">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="96eb7-127">运行脚本之前</span><span class="sxs-lookup"><span data-stu-id="96eb7-127">Before running the script</span></span>

![显示学生的分数行的工作表。](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="96eb7-129">运行脚本后</span><span class="sxs-lookup"><span data-stu-id="96eb7-129">After running the script</span></span>

![在有效学生行中显示带有红色总计的无效单元格的学生分数数据的工作表。](../../images/scenario-grade-calculator-after.png)
