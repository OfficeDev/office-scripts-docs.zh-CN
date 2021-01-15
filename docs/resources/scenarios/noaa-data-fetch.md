---
title: Office 脚本示例方案：绘制 NOAA 中的水级数据
description: 从 NOAA 数据库提取 JSON 数据并使用它创建图表的示例。
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: 5b0b4e3675cbe053368f63123d819f0dab626e60
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867875"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a><span data-ttu-id="f175a-103">Office 脚本示例方案：从 NOAA 提取和绘制水级数据</span><span class="sxs-lookup"><span data-stu-id="f175a-103">Office Scripts sample scenario: Fetch and graph water-level data from NOAA</span></span>

<span data-ttu-id="f175a-104">在此方案中，你需要绘制国家远洋和水管理局的 [西雅图站的水位](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)。</span><span class="sxs-lookup"><span data-stu-id="f175a-104">In this scenario, you need to plot the water level at the [National Oceanic and Atmospheric Administration's Seattle station](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130).</span></span> <span data-ttu-id="f175a-105">你将使用外部数据填充电子表格并创建图表。</span><span class="sxs-lookup"><span data-stu-id="f175a-105">You'll use external data to populate a spreadsheet and create a chart.</span></span>

<span data-ttu-id="f175a-106">您将开发一个脚本，该脚本使用 `fetch` 命令查询 [NOAA"更新和当前"数据库](https://tidesandcurrents.noaa.gov/)。</span><span class="sxs-lookup"><span data-stu-id="f175a-106">You'll develop a script that uses the `fetch` command to query the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/).</span></span> <span data-ttu-id="f175a-107">这将获取在给定时间跨度中记录的水位。</span><span class="sxs-lookup"><span data-stu-id="f175a-107">That will get the water level recorded across a given time span.</span></span> <span data-ttu-id="f175a-108">该信息将返回为 JSON，因此脚本的一部分将转换为区域值。</span><span class="sxs-lookup"><span data-stu-id="f175a-108">The information will be returned as JSON, so part of the script will translate that into range values.</span></span> <span data-ttu-id="f175a-109">数据位于电子表格中后，它将用于制作图表。</span><span class="sxs-lookup"><span data-stu-id="f175a-109">Once the data is in the spreadsheet, it will be used to make a chart.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="f175a-110">涵盖的脚本编写技能</span><span class="sxs-lookup"><span data-stu-id="f175a-110">Scripting skills covered</span></span>

- <span data-ttu-id="f175a-111">外部 API 调用 `fetch` () </span><span class="sxs-lookup"><span data-stu-id="f175a-111">External API calls (`fetch`)</span></span>
- <span data-ttu-id="f175a-112">JSON 分析</span><span class="sxs-lookup"><span data-stu-id="f175a-112">JSON parsing</span></span>
- <span data-ttu-id="f175a-113">图表</span><span class="sxs-lookup"><span data-stu-id="f175a-113">Charts</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="f175a-114">安装说明</span><span class="sxs-lookup"><span data-stu-id="f175a-114">Setup instructions</span></span>

1. <span data-ttu-id="f175a-115">使用 Excel 网页打开工作簿。</span><span class="sxs-lookup"><span data-stu-id="f175a-115">Open the workbook with Excel on the web.</span></span>

1. <span data-ttu-id="f175a-116">在"**自动化"** 选项卡下，选择 **"所有脚本"。**</span><span class="sxs-lookup"><span data-stu-id="f175a-116">Under the **Automate** tab, select **All Scripts**.</span></span>

1. <span data-ttu-id="f175a-117">在" **代码编辑器"** 任务窗格中，选择 **"新建脚本** "，然后将以下脚本粘贴到编辑器中。</span><span class="sxs-lookup"><span data-stu-id="f175a-117">In the **Code Editor** task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```typescript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook): Promise<void> {
      // Get the current sheet.
      let currentSheet = workbook.getActiveWorksheet();
    
      // Create selection of parameters for the fetch URL.
      // More information on the NOAA APIs is found here: 
      // https://api.tidesandcurrents.noaa.gov/api/prod/
      const option = "water_level";
      const startDate = "20201225"; /* yyyymmdd date format */
      const endDate = "20201227";
      const station = "9447130"; /* Seattle */
    
      // Construct the URL for the fetch call.
      const strQuery = `https://api.tidesandcurrents.noaa.gov/api/prod/datagetter?product=${option}&begin_date=${startDate}&end_date=${endDate}&datum=MLLW&station=${station}&units=english&time_zone=gmt&application=NOS.COOPS.TAC.WL&format=json`;
    
      console.log(strQuery);
    
      // Resolve the Promises returned by the fetch operation.
      const response = await fetch(strQuery);
      const rawJson = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
      const noaaData = JSON.parse(stringifiedJson);
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.data.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData.data[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
      
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth,dataRange);
      chart.getTitle().setText("Water Level - Seattle");
      chart.setTop(0);
      chart.setLeft(300);
      chart.setWidth(500);
      chart.setHeight(300);
      chart.getAxes().getValueAxis().setShowDisplayUnitLabel(false);
      chart.getAxes().getCategoryAxis().setTextOrientation(60);
      chart.getLegend().setVisible(false);

      // Add a comment with the data attribution.
      currentSheet.addComment(
        "A1", 
        `This data was taken from the National Oceanic and Atmospheric Administration's Tides and Currents database on ${new Date(Date.now())}.`
      );
    }
    ```

1. <span data-ttu-id="f175a-118">将该脚本重命名为 **NOAA 水级别图表** 并保存它。</span><span class="sxs-lookup"><span data-stu-id="f175a-118">Rename the script to **NOAA Water Level Chart** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="f175a-119">运行脚本</span><span class="sxs-lookup"><span data-stu-id="f175a-119">Running the script</span></span>

<span data-ttu-id="f175a-120">在任何工作表上，运行 **NOAA 水位图表** 脚本。</span><span class="sxs-lookup"><span data-stu-id="f175a-120">On any worksheet, run the **NOAA Water Level Chart** script.</span></span> <span data-ttu-id="f175a-121">该脚本提取从 2020 年 12 月 25 日到 2020 年 12 月 27 日的级别数据。</span><span class="sxs-lookup"><span data-stu-id="f175a-121">The script fetches the water level data from December 25, 2020 to December 27, 2020.</span></span> <span data-ttu-id="f175a-122">可以将 `const` 脚本开头的变量更改为使用不同的日期或获取不同的工作站信息。</span><span class="sxs-lookup"><span data-stu-id="f175a-122">The `const` variables at the beginning of the script can be changed to use different dates or get different station information.</span></span> <span data-ttu-id="f175a-123">[用于数据检索的 CO-OPS API](https://api.tidesandcurrents.noaa.gov/api/prod/)介绍如何获取所有这些数据。</span><span class="sxs-lookup"><span data-stu-id="f175a-123">The [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) describes how to get all this data.</span></span>

### <a name="after-running-the-script"></a><span data-ttu-id="f175a-124">运行脚本后</span><span class="sxs-lookup"><span data-stu-id="f175a-124">After running the script</span></span>

![运行脚本后的工作表显示一些水位数据和图表。](../../images/scenario-noaa-water-level-after.png)
