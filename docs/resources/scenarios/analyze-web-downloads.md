---
title: Office 脚本示例方案：分析 web 下载
description: 在将该信息组织到表中之前，获取 Excel 工作簿中的原始 internet 流量数据并确定原始位置的示例。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: adc2cb401830b66b245c0dfcc4441b7ac9c8c61f
ms.sourcegitcommit: 009935c5773761c5833e5857491af47e2c95d851
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/25/2020
ms.locfileid: "49408964"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="aba3a-103">Office 脚本示例方案：分析 web 下载</span><span class="sxs-lookup"><span data-stu-id="aba3a-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="aba3a-104">在这种情况下，您将负责分析来自公司网站的下载报告。</span><span class="sxs-lookup"><span data-stu-id="aba3a-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="aba3a-105">此分析的目标是确定 web 流量是来自世界各地还是其他地方。</span><span class="sxs-lookup"><span data-stu-id="aba3a-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="aba3a-106">您的同事将原始数据上传到工作簿。</span><span class="sxs-lookup"><span data-stu-id="aba3a-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="aba3a-107">每周的数据集都有自己的工作表。</span><span class="sxs-lookup"><span data-stu-id="aba3a-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="aba3a-108">还有一个显示每周趋势的表格和图表的 **摘要** 工作表。</span><span class="sxs-lookup"><span data-stu-id="aba3a-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="aba3a-109">您将开发一个脚本，用于分析活动工作表中每周的下载数据。</span><span class="sxs-lookup"><span data-stu-id="aba3a-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="aba3a-110">它将分析与每个下载关联的 IP 地址，并确定是否来自美国。</span><span class="sxs-lookup"><span data-stu-id="aba3a-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="aba3a-111">该答案将作为布尔值插入到工作表中， ( 为 "TRUE" 或 "FALSE" ) 并且条件格式将应用于这些单元格。</span><span class="sxs-lookup"><span data-stu-id="aba3a-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="aba3a-112">将在工作表上汇总 IP 地址位置结果，并将其复制到摘要表。</span><span class="sxs-lookup"><span data-stu-id="aba3a-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="aba3a-113">涵盖的脚本技能</span><span class="sxs-lookup"><span data-stu-id="aba3a-113">Scripting skills covered</span></span>

- <span data-ttu-id="aba3a-114">文本解析</span><span class="sxs-lookup"><span data-stu-id="aba3a-114">Text parsing</span></span>
- <span data-ttu-id="aba3a-115">脚本中的 Subfunctions</span><span class="sxs-lookup"><span data-stu-id="aba3a-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="aba3a-116">条件格式</span><span class="sxs-lookup"><span data-stu-id="aba3a-116">Conditional formatting</span></span>
- <span data-ttu-id="aba3a-117">表格</span><span class="sxs-lookup"><span data-stu-id="aba3a-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="aba3a-118">演示视频</span><span class="sxs-lookup"><span data-stu-id="aba3a-118">Demo video</span></span>

<span data-ttu-id="aba3a-119">此示例是 demoed 年2月2020的 Office 外接程序开发人员社区呼叫的一部分。</span><span class="sxs-lookup"><span data-stu-id="aba3a-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

> [!NOTE]
> <span data-ttu-id="aba3a-120">此视频中显示的代码使用较旧的 API 模型 ([Office 脚本的异步 api](../../develop/excel-async-model.md)) 。</span><span class="sxs-lookup"><span data-stu-id="aba3a-120">The code shown in this video uses an older API model (the [Office Scripts Async APIs](../../develop/excel-async-model.md)).</span></span> <span data-ttu-id="aba3a-121">此页面上显示的示例已更新，但代码看上去与录制略有不同。</span><span class="sxs-lookup"><span data-stu-id="aba3a-121">The sample presented on this page has been updated, but the code looks a little different from the recording.</span></span> <span data-ttu-id="aba3a-122">这些更改不会影响脚本的行为或演示者演示中的其他内容。</span><span class="sxs-lookup"><span data-stu-id="aba3a-122">The changes don't affect the behavior of the script or the other content in the presenter's demo.</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="aba3a-123">设置说明</span><span class="sxs-lookup"><span data-stu-id="aba3a-123">Setup instructions</span></span>

1. <span data-ttu-id="aba3a-124">将 <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> 下载到你的 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="aba3a-124">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="aba3a-125">使用适用于 web 的 Excel 打开工作簿。</span><span class="sxs-lookup"><span data-stu-id="aba3a-125">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="aba3a-126">在 " **自动化** " 选项卡上，打开 **代码编辑器**。</span><span class="sxs-lookup"><span data-stu-id="aba3a-126">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="aba3a-127">在 " **代码编辑器** " 任务窗格中，按 " **新建脚本** "，并将以下脚本粘贴到编辑器中。</span><span class="sxs-lookup"><span data-stu-id="aba3a-127">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      /* Get the Summary worksheet and table.
       * End the script early if either object is not in the workbook.
       */
      let summaryWorksheet = workbook.getWorksheet("Summary");
      if (!summaryWorksheet) {
        console.log("The script expects a worksheet named \"Summary\". Please download the correct template and try again.");
        return;
      }
      let summaryTable = summaryWorksheet.getTable("Table1");
      if (!summaryTable) {
        console.log("The script expects a summary table named \"Table1\". Please download the correct template and try again.");
        return;
      }

      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (currentWorksheet.getName().toLocaleLowerCase().indexOf("week") !== 0) {
        console.log("Please switch worksheet to one of the weekly data sheets and try again.")
        return;
      }

      // Get the values of the active range of the active worksheet.
      let logRange = currentWorksheet.getUsedRange();

      if (logRange.getColumnCount() !== 8) {
        console.log(`Verify that you are on the correct worksheet. Either the week's data has been already processed or the content is incorrect. The following columns are expected: ${[
          "Time Stamp", "IP Address", "kilobytes", "user agent code", "milliseconds", "Request", "Results", "Referrer"
        ]}`);
        return;
      }
      // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
      let isUSColumn = logRange
        .getLastColumn()
        .getOffsetRange(0, 1);

      // Get the values of all the US IP addresses.
      let ipRange = workbook.getWorksheet("USIPAddresses").getUsedRange();
      let ipRangeValues = ipRange.getValues();
      let logRangeValues = logRange.getValues();
      // Remove the first row.
      let topRow = logRangeValues.shift();
      console.log(`Analyzing ${logRangeValues.length} entries.`);

      // Create a new array to contain the boolean representing if this is a US IP address.
      let newCol = [];

      // Go through each row in worksheet and add Boolean.
      for (let i = 0; i < logRangeValues.length; i++) {
        let curRowIP = logRangeValues[i][1];
        if (findIP(ipRangeValues, ipAddressToInteger(curRowIP)) > 0) {
          newCol.push([true]);
        } else {
          newCol.push([false]);
        }
      }

      // Remove the empty column header and add proper heading.
      newCol = [["Is US IP"], ...newCol];

      // Write the result to the spreadsheet.
      console.log(`Adding column to indicate whether IP belongs to US region or not at address: ${isUSColumn.getAddress()}`);
      console.log(newCol.length);
      console.log(newCol);
      isUSColumn.setValues(newCol);

      // Call the local function to add summary data to the worksheet.
      addSummaryData();

      // Call the local function to apply conditional formatting.
      applyConditionalFormatting(isUSColumn);

      // Autofit columns.
      currentWorksheet.getUsedRange().getFormat().autofitColumns();

      // Get the calculated summary data.
      let summaryRangeValues = currentWorksheet.getRange("J2:M2").getValues();

      // Add the corresponding row to the summary table.
      summaryTable.addRow(null, summaryRangeValues[0]);
      console.log("Complete.");
      return;

      /**
       * A function to add summary data on the worksheet.
       */
      function addSummaryData() {
        // Add a summary row and table.
        let summaryHeader = [["Year", "Week", "US", "Other"]];
        let countTrueFormula =
          "=COUNTIF(" + isUSColumn.getAddress() + ', "=TRUE")/' + (newCol.length - 1);
        let countFalseFormula =
          "=COUNTIF(" + isUSColumn.getAddress() + ', "=FALSE")/' + (newCol.length - 1);

        let summaryContent = [
          [
            '=TEXT(A2,"YYYY")',
            '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
            countTrueFormula,
            countFalseFormula
          ]
        ];
        let summaryHeaderRow = currentWorksheet
          .getRange("J1:M1");
        let summaryContentRow = currentWorksheet
          .getRange("J2:M2");
        console.log("2");

        summaryHeaderRow.setValues(summaryHeader);
        console.log("3");

        summaryContentRow.setValues(summaryContent);
        console.log("4");

        let formats = [[".000", ".000"]];
        summaryContentRow
          .getOffsetRange(0, 2)
          .getResizedRange(0, -2).setNumberFormats(formats);
      }
    }
    /**
     * Apply conditional formatting based on TRUE/FALSE values of the Is US IP column.
     */
    function applyConditionalFormatting(isUSColumn: ExcelScript.Range) {
      // Add conditional formatting to the new column.
      let conditionalFormatTrue = isUSColumn.addConditionalFormat(
        ExcelScript.ConditionalFormatType.cellValue
      );
      let conditionalFormatFalse = isUSColumn.addConditionalFormat(
        ExcelScript.ConditionalFormatType.cellValue
      );
      // Set TRUE to light blue and FALSE to light orange.
      conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#8FA8DB");
      conditionalFormatTrue.getCellValue().setRule({
        formula1: "=TRUE",
        operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
      conditionalFormatFalse.getCellValue().getFormat().getFill().setColor("#F8CCAD");
      conditionalFormatFalse.getCellValue().setRule({
        formula1: "=FALSE",
        operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
    }
    /**
     * Translate an IP address into an integer.
     * @param ipAddress: IP address to verify.
     */
    function ipAddressToInteger(ipAddress: string): number {
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
    /**
     * Return the row number where the ip address is found.
     * @param ipLookupTable IP look-up table.
     * @param n IP address to number value.  
     */
    function findIP(ipLookupTable: number[][], n: number): number {
      for (let i = 0; i < ipLookupTable.length; i++) {
        if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
          return i;
        }
      }
      return -1;
    }
    ```

5. <span data-ttu-id="aba3a-128">重命名脚本以 **分析 Web 下载** 并保存它。</span><span class="sxs-lookup"><span data-stu-id="aba3a-128">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="aba3a-129">运行脚本</span><span class="sxs-lookup"><span data-stu-id="aba3a-129">Running the script</span></span>

<span data-ttu-id="aba3a-130">导航到任意 **一周 \* \*** 工作表，并运行 **分析 Web 下载** 脚本。</span><span class="sxs-lookup"><span data-stu-id="aba3a-130">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="aba3a-131">该脚本将应用在当前工作表上标签的条件格式和位置。</span><span class="sxs-lookup"><span data-stu-id="aba3a-131">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="aba3a-132">它还将更新 **摘要** 工作表。</span><span class="sxs-lookup"><span data-stu-id="aba3a-132">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="aba3a-133">运行脚本之前</span><span class="sxs-lookup"><span data-stu-id="aba3a-133">Before running the script</span></span>

![显示原始 web 流量数据的工作表。](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="aba3a-135">运行脚本后</span><span class="sxs-lookup"><span data-stu-id="aba3a-135">After running the script</span></span>

![显示以前的 web 流量行的格式化 IP 位置信息的工作表。](../../images/scenario-analyze-web-downloads-after.png)

![汇总表和图表，其中汇总了运行脚本的工作表。](../../images/scenario-analyze-web-downloads-table.png)
