---
title: Office脚本示例方案：Graph来自 NOAA 的水位数据
description: 从 NOAA 数据库提取 JSON 数据并使用它创建图表的示例。
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4181edae7d8a46ae381ddfb1a2893b03faffd9b
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088097"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Office脚本示例方案：从 NOAA 提取和绘制水位数据

在这种情况下，你需要绘制 [国家海洋和大气管理局西雅图站的](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)水位。 你将使用外部数据填充电子表格并创建图表。

你将开发一个脚本，该脚本使用该 `fetch` 命令查询 [NOAA Tides 和 Currents 数据库](https://tidesandcurrents.noaa.gov/)。 这将在给定时间范围内记录水位。 信息将作为 [JSON](https://www.w3schools.com/whatis/whatis_json.asp) 返回，因此脚本的一部分会将其转换为范围值。 数据在电子表格中后，将用于制作图表。

有关使用 JSON 的详细信息，请阅读[使用 JSON 向Office脚本传递数据](../../develop/use-json.md)。

## <a name="scripting-skills-covered"></a>所涵盖的脚本技能

- 外部 API 调用 (`fetch`) 
- JSON 分析
- 图表

## <a name="setup-instructions"></a>设置说明

1. 使用Excel web 版打开工作簿。

1. 在 **“自动执行”** 选项卡下，选择 **“新建脚本** ”并将以下脚本粘贴到编辑器中。

    ```TypeScript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook) {
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

1. 将脚本重命名为 **NOAA 水位图** 并保存它。

## <a name="running-the-script"></a>运行脚本

在任何工作表上，运行 **NOAA 水位图** 脚本。 该脚本提取 2020 年 12 月 25 日至 2020 年 12 月 27 日的水位数据。 `const`脚本开头的变量可以更改为使用不同的日期或获取不同的站信息。 [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) 介绍了如何获取所有这些数据。

### <a name="after-running-the-script"></a>运行脚本后

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="运行脚本后的工作表显示一些水位数据和图表。":::
