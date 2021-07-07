---
title: 'Office Пример сценария: калькулятор оценки'
description: Пример, определяя процент и оценки букв для класса учащихся.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 2d98e68f37418ade238a707cb74cc7ccf47e8f59
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313794"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Office Пример сценария: калькулятор оценки

В этом сценарии вы будете инструктором, который подытвет оценки каждого учащегося по окончании срока обучения. Вы вводя оценки для их назначений и тестов, как вы идете. Теперь настало время определить судьбы учащихся.

Вы разработает сценарий, который суммит оценки для каждой категории точки. Затем он назначает каждому учащемуся оценку буквы в зависимости от общей суммы. Чтобы обеспечить точность, вы добавим несколько проверок, чтобы убедиться, что какие-либо отдельные оценки слишком низкие или высокие. Если оценка учащегося меньше нуля или больше возможного значения точки, скрипт будет пометить ячейку с красной заливкой, а не суммой точек этого студента. Это будет четким указанием, какие записи необходимо перепросмотрить. Вы также добавим некоторые базовые форматирования в оценки, чтобы можно было быстро просмотреть верхнюю и нижнюю части класса.

## <a name="scripting-skills-covered"></a>Навыки скриптов, охватываемых

- Форматирование ячейки
- Проверка ошибок
- Регулярные выражения
- Условное форматирование

## <a name="setup-instructions"></a>Инструкции по настройке

1. Скачайте <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> в OneDrive.

1. Откройте книгу с помощью Excel для Интернета.

1. В **вкладке Automate** выберите **Новый скрипт** и вклеите следующий скрипт в редактор.

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

1. Переименуй сценарий в **калькулятор класса и** сохраните его.

## <a name="running-the-script"></a>Выполнение скрипта

Запустите **сценарий калькулятора класса** на единственной таблице. Сценарий будет общую оценку и назначить каждому учащемуся оценку буквы. Если в отдельных классах имеется больше баллов, чем стоит назначение или тест, то класс обижающих будет отмечен красным, а общее число не вычисляется. Кроме того, все оценки "A" выделены зеленым цветом, а оценки "D" и "F" выделены желтым цветом.

### <a name="before-running-the-script"></a>Перед запуском сценария

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="Таблица, в которую показаны строки баллов для учащихся.":::

### <a name="after-running-the-script"></a>После запуска скрипта

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="Таблица, в которую показаны данные о оценках учащихся с недействительными ячейками в красных итоговых числах для допустимых студенческих строк.":::
