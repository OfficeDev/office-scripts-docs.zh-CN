---
title: Office 脚本示例方案：分析 web 下载
description: 在将该信息组织到表中之前，获取 Excel 工作簿中的原始 internet 流量数据并确定原始位置的示例。
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 9ee12c8d4ca7c191168e3734d7cd9eadc333c165
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700132"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="cf543-103">Office 脚本示例方案：分析 web 下载</span><span class="sxs-lookup"><span data-stu-id="cf543-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="cf543-104">在这种情况下，您将负责分析来自公司网站的下载报告。</span><span class="sxs-lookup"><span data-stu-id="cf543-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="cf543-105">此分析的目标是确定 web 流量是来自世界各地还是其他地方。</span><span class="sxs-lookup"><span data-stu-id="cf543-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="cf543-106">您的同事将原始数据上传到工作簿。</span><span class="sxs-lookup"><span data-stu-id="cf543-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="cf543-107">每周的数据集都有自己的工作表。</span><span class="sxs-lookup"><span data-stu-id="cf543-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="cf543-108">还有一个显示每周趋势的表格和图表的**摘要**工作表。</span><span class="sxs-lookup"><span data-stu-id="cf543-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="cf543-109">您将开发一个脚本，用于分析活动工作表中每周的下载数据。</span><span class="sxs-lookup"><span data-stu-id="cf543-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="cf543-110">它将分析与每个下载关联的 IP 地址，并确定是否来自美国。</span><span class="sxs-lookup"><span data-stu-id="cf543-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="cf543-111">该答案将作为布尔值（"TRUE" 或 "FALSE"）插入到工作表中，条件格式将应用于这些单元格。</span><span class="sxs-lookup"><span data-stu-id="cf543-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="cf543-112">将在工作表上汇总 IP 地址位置结果，并将其复制到摘要表。</span><span class="sxs-lookup"><span data-stu-id="cf543-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="cf543-113">涵盖的脚本技能</span><span class="sxs-lookup"><span data-stu-id="cf543-113">Scripting skills covered</span></span>

- <span data-ttu-id="cf543-114">文本解析</span><span class="sxs-lookup"><span data-stu-id="cf543-114">Text parsing</span></span>
- <span data-ttu-id="cf543-115">脚本中的 Subfunctions</span><span class="sxs-lookup"><span data-stu-id="cf543-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="cf543-116">条件格式</span><span class="sxs-lookup"><span data-stu-id="cf543-116">Conditional formatting</span></span>
- <span data-ttu-id="cf543-117">表</span><span class="sxs-lookup"><span data-stu-id="cf543-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="cf543-118">演示视频</span><span class="sxs-lookup"><span data-stu-id="cf543-118">Demo video</span></span>

<span data-ttu-id="cf543-119">此示例是 demoed 年2月2020的 Office 外接程序开发人员社区呼叫的一部分。</span><span class="sxs-lookup"><span data-stu-id="cf543-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a><span data-ttu-id="cf543-120">设置说明</span><span class="sxs-lookup"><span data-stu-id="cf543-120">Setup instructions</span></span>

1. <span data-ttu-id="cf543-121">将<a href="analyze-web-downloads.xlsx">analyze-web-downloads</a>下载到你的 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="cf543-121">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="cf543-122">使用适用于 web 的 Excel 打开工作簿。</span><span class="sxs-lookup"><span data-stu-id="cf543-122">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="cf543-123">在 "**自动化**" 选项卡上，打开**代码编辑器**。</span><span class="sxs-lookup"><span data-stu-id="cf543-123">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="cf543-124">在 "**代码编辑器**" 任务窗格中，按 "**新建脚本**"，并将以下脚本粘贴到编辑器中。</span><span class="sxs-lookup"><span data-stu-id="cf543-124">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
      async function main(context: Excel.RequestContext) {
        let currentWorksheet = context.workbook.worksheets
          .getActiveWorksheet();
        // Get the values of the active range of the active worksheet.
        let logRange = currentWorksheet.getUsedRange().load("values");

        // Get the Summary worksheet and table.
        let summaryWorksheet = context.workbook.worksheets.getItem("Summary");
        let summaryTable = context.workbook.tables.getItem("Table1");

        // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
        let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1)
          .load("address");

        // Get the values of all the US IP addresses.
        let ipRange = context.workbook.worksheets
          .getItem("USIPAddresses")
          .getUsedRange()
          .load("values");
        await context.sync();

        // Remove the first row.
        let topRow = logRange.values.shift();

        // Create a new array to contain the boolean representing if this is a US IP address.
        let newCol = [[]];

        // Go through each row in worksheet and add Boolean.
        for (let i = 0; i < logRange.values.length; i++) {
          let curRowIP = logRange.values[i][1];
          if (findIP(ipRange.values, ipAddressToInteger(curRowIP)) > 0) {
            newCol.push([true]);
          } else {
            newCol.push([false]);
          }
        }

        // Remove the empty column header and add proper heading.
        newCol.shift();
        newCol.unshift(["Is US IP"]);

        // Write the result to the spreadsheet.
        isUSColumn.values = newCol;
        addSummaryData();
        applyConditionalFormatting();
        currentWorksheet.getUsedRange().format.autofitColumns();

        // Get the calculated summary data.
        let summaryRange = currentWorksheet.getRange("J2:M2").load("values");
        await context.sync();

        // Add the corresponding row to the summary table.
        summaryTable.rows.add(null, summaryRange.values);

        // Function to apply conditional formatting to the new column.
        function applyConditionalFormatting() {
          // Add conditional formatting to the new column.
          let conditionalFormatTrue = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          let conditionalFormatFalse = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          // Set TRUE to light blue and FALSE to light orange.
          conditionalFormatTrue.cellValue.format.fill.color = "#8FA8DB";
          conditionalFormatTrue.cellValue.rule = {
            formula1: "=TRUE",
            operator: "EqualTo"
          };
          conditionalFormatFalse.cellValue.format.fill.color = "#F8CCAD";
          conditionalFormatFalse.cellValue.rule = {
            formula1: "=FALSE",
            operator: "EqualTo"
          };
        }

        // Adds the summary data to the current sheet and to the summary table.
        function addSummaryData() {
          // Add a summary row and table.
          let summaryHeader = [["Year", "Week", "US", "Other"]];
          let countTrueFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=TRUE")/' + (newCol.length - 1);
          let countFalseFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=FALSE")/' + (newCol.length - 1);

          let summaryContent = [
            [
              '=TEXT(A2,"YYYY")',
              '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
              countTrueFormula,
              countFalseFormula
            ]
          ];
          let summaryHeaderRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J1:M1");
          let summaryContentRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J2:M2");
          summaryHeaderRow.values = summaryHeader;
          summaryContentRow.values = summaryContent;
          let formats = [[".000", ".000"]];
          summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).numberFormat = formats;
        }
      }

      // Translate an IP address into an integer.
      function ipAddressToInteger(ipAddress: string) {
        // Split the IP address into octets.
        let octets = ipAddress.split(".");

        // Create a number for each octet and do the math to create the integer value of the IP address.
        let fullNum =
          // Define an arbitrary number for the last octet.
          111 +
          parseInt(octets[2]) * 256 +
          parseInt(octets[1]) * 65536 +
          parseInt(octets[0]) * 16777216;
        return fullNum;
      }

      // Return the row number where the ip address is found.
      function findIP(ipLookupTable: number[][], n: number) {
        for (let i = 0; i < ipLookupTable.length; i++) {
          if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
            return i;
          }
        }
        return -1;
      }
    ```

5. <span data-ttu-id="cf543-125">重命名脚本以**分析 Web 下载**并保存它。</span><span class="sxs-lookup"><span data-stu-id="cf543-125">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="cf543-126">运行脚本</span><span class="sxs-lookup"><span data-stu-id="cf543-126">Running the script</span></span>

<span data-ttu-id="cf543-127">导航到任意**一周\* **工作表，并运行**分析 Web 下载**脚本。</span><span class="sxs-lookup"><span data-stu-id="cf543-127">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="cf543-128">该脚本将应用在当前工作表上标签的条件格式和位置。</span><span class="sxs-lookup"><span data-stu-id="cf543-128">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="cf543-129">它还将更新**摘要**工作表。</span><span class="sxs-lookup"><span data-stu-id="cf543-129">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="cf543-130">运行脚本之前</span><span class="sxs-lookup"><span data-stu-id="cf543-130">Before running the script</span></span>

![显示原始 web 流量数据的工作表。](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="cf543-132">运行脚本后</span><span class="sxs-lookup"><span data-stu-id="cf543-132">After running the script</span></span>

![显示以前的 web 流量行的格式化 IP 位置信息的工作表。](../../images/scenario-analyze-web-downloads-after.png)

![汇总表和图表，其中汇总了运行脚本的工作表。](../../images/scenario-analyze-web-downloads-table.png)
