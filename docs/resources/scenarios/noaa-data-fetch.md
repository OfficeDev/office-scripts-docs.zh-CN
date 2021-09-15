---
title: Office脚本示例方案：Graph NOAA 中的水级数据
description: 从 NOAA 数据库提取 JSON 数据并使用它创建图表的示例。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: f0492c79b9fc2d7d98f4433611fd8589cf52054a
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/15/2021
ms.locfileid: "59327887"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Office脚本示例方案：从 NOAA 提取和绘制水级数据

在此方案中，你需要绘制"国家/地区"和"城市管理" [的西雅图站的水底](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)。 您将使用外部数据填充电子表格并创建图表。

您将开发一个脚本，该脚本使用 `fetch` 命令查询 [NOAA 的"产品"和"当前"数据库](https://tidesandcurrents.noaa.gov/)。 这将获取在给定时间跨度中记录的水位。 信息将作为 JSON 返回，因此脚本的一部分会将其转换为范围值。 数据位于电子表格中后，它将用于制作图表。

## <a name="scripting-skills-covered"></a>涵盖的脚本编写技能

- 外部 API 调用 `fetch` () 
- JSON 分析
- 图表

## <a name="setup-instructions"></a>设置说明

1. 打开工作簿，Excel web 版。

1. 在" **自动化"** 选项卡下，选择 **"新建脚本** "，然后将以下脚本粘贴到编辑器中。

    ```TypeScript
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
      const rawJson: string = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
    
      // Note that we're only taking the data part of the JSON and excluding the metadata.
      const noaaData: NOAAData[] = JSON.parse(stringifiedJson).data;
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
    
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth, dataRange);
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
    
      /**
       * An interface to wrap the parts of the JSON we need.
       * These properties must match the names used in the JSON.
       */ 
      interface NOAAData {
        t: string; // Time
        v: number; // Level
      }
    }
    ```

1. 将该脚本重命名为 **NOAA 水级图** 并保存它。

## <a name="running-the-script"></a>运行脚本

在任何工作表上，运行 **NOAA 水级图表** 脚本。 该脚本提取从 2020 年 12 月 25 日到 2020 年 12 月 27 日的水级数据。 可以将 `const` 脚本开头的变量更改为使用不同的日期或获取不同的工作站信息。 [用于数据检索的 CO-OPS API](https://api.tidesandcurrents.noaa.gov/api/prod/)介绍如何获取所有这些数据。

### <a name="after-running-the-script"></a>运行脚本后

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="运行脚本后的工作表显示一些水级数据和图表。":::
