---
title: Office脚本示例方案：“打孔时钟”按钮
description: 此示例添加一个打孔时钟按钮，并允许用户使用当前时间打卡和打卡。
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: de56fb651d6f6088620678cfd72ce662875eafa7
ms.sourcegitcommit: e6428a5214fa38aef036a952a0e3c09dbf6e4d3e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/28/2022
ms.locfileid: "65109287"
---
# <a name="office-scripts-sample-scenario-punch-clock-button"></a>Office脚本示例方案：“打孔时钟”按钮

本示例中使用的方案构想和脚本由Office脚本社区成员 [Brian Gonzalez](https://github.com/b-gonzalez) 提供。

在此方案中，你将为员工创建一个时间表，允许他们使用 [按钮](../../develop/script-buttons.md)记录开始和结束时间。 根据之前记录的内容，按下按钮将在) 开始他们的一天 (时钟或结束他们的一天 (时钟) 。 该示例适用于Excel web 版和Windows。

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="工作簿中包含三列 (“Clock In”、“Clock Out”和“Duration”) 和标记为“打孔时钟”的按钮的表。":::

## <a name="setup-instructions"></a>设置说明

1. <a href="punch-clock-sample.xlsx"> 将punch-clock-sample.xlsx</a>下载到OneDrive。

    :::image type="content" source="../../images/punch-clock-sample-1.png" alt-text="包含三列的表：“Clock In”、“Clock Out”和“Duration”。":::

1. 在Excel web 版中打开工作簿。

1. 在 **“自动执行”** 选项卡下，选择 **“新建脚本** ”并将以下脚本粘贴到编辑器中。

    ```typescript
    /**
     * This script records either the start or end time of a shift, 
     * depending on what is filled out in the table. 
     * It is intended to be used with a Script Button.
     */
    function main(workbook: ExcelScript.Workbook) {
      // Get the first table in the timesheet.
      const timeSheet = workbook.getWorksheet("MyTimeSheet");
      const timeTable = timeSheet.getTables()[0];
    
      // Get the appropriate table columns.
      const clockInColumn = timeTable.getColumnByName("Clock In");
      const clockOutColumn = timeTable.getColumnByName("Clock Out");
      const durationColumn = timeTable.getColumnByName("Duration");
    
      // Get the last rows for the Clock In and Clock Out columns.
      let clockInLastRow = clockInColumn.getRangeBetweenHeaderAndTotal().getLastRow();
      let clockOutLastRow = clockOutColumn.getRangeBetweenHeaderAndTotal().getLastRow();
    
      // Get the current date to use as the start or end time.
      let date: Date = new Date();
    
      // Add the current time to a column based on the state of the table.
      if (clockInLastRow.getValue() as string === "") {
        // If the Clock In column has an empty value in the table, add a start time.
        clockInLastRow.setValue(date.toLocaleString());
      } else if (clockOutLastRow.getValue() as string === "") {
        // If the Clock Out column has an empty value in the table, 
        // add an end time and calculate the shift duration.
        clockOutLastRow.setValue(date.toLocaleString());
        const clockInTime = new Date(clockInLastRow.getValue() as string);
        const clockOutTime  = new Date(clockOutLastRow.getValue() as string);
        const clockDuration = Math.abs((clockOutTime.getTime() - clockInTime.getTime()));
    
        let durationString = getDurationMessage(clockDuration);
        durationColumn.getRangeBetweenHeaderAndTotal().getLastRow().setValue(durationString);
      } else {
        // If both columns are full, add a new row, then add a start time.
        timeTable.addRow()
        clockInLastRow.getOffsetRange(1, 0).setValue(date.toLocaleString());
      }
    }
    
    /**
     * A function to write a time duration as a string.
     */
    function getDurationMessage(delta: number) {
      // Adapted from here:
      // https://stackoverflow.com/questions/13903897/javascript-return-number-of-days-hours-minutes-seconds-between-two-dates
    
      delta = delta / 1000;
      let durationString = "";
    
      let days = Math.floor(delta / 86400);
      delta -= days * 86400;
    
      let hours = Math.floor(delta / 3600) % 24;
      delta -= hours * 3600;
    
      let minutes = Math.floor(delta / 60) % 60;
    
      if (days >= 1) {
        durationString += days;
        durationString += (days > 1 ? " days" : " day");
    
        if (hours >= 1 && minutes >= 1) {
          durationString += ", ";
        }
        else if (hours >= 1 || minutes > 1) {
          durationString += " and ";
        }
      }
    
      if (hours >= 1) {
        durationString += hours;
        durationString += (hours > 1 ? " hours" : " hour");
        if (minutes >= 1) {
          durationString += " and ";
        }
      }
    
      if (minutes >= 1) {
        durationString += minutes;
        durationString += (minutes > 1 ? " minutes" : " minute");
      }
    
      return durationString;
    }
    ```

1. 将脚本重命名为“打孔时钟”。

1. 保存脚本。

1. 在工作簿中，选择 **单元格 E2**。

1. 添加脚本按钮。 转到“**脚本详细信息**”页 **中的“更多”选项 (...)** 菜单，然后选择 **“添加”按钮**。

    :::image type="content" source="../../images/punch-clock-sample-2.png" alt-text="“更多选项”菜单和“添加按钮”按钮。":::

1. 保存工作簿。

## <a name="run-the-script"></a>运行脚本

按 **“打孔时钟”** 按钮运行脚本。 它根据之前输入的时间，在“Clock In”或“Clock Out”下记录当前时间。

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="工作簿中的表和“打孔时钟”按钮。":::

> [!NOTE]
> 仅当持续时间超过一分钟时，才会记录持续时间。 手动编辑“时钟输入”时间以测试更长的持续时间。
