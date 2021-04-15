---
title: 'Пример сценария Office Scripts: калькулятор оценки'
description: Пример, определяя процент и оценки букв для класса учащихся.
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: b8c45ad405c06a943c75e76391c1160ecb1bd18e
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755030"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="13722-103">Пример сценария Office Scripts: калькулятор оценки</span><span class="sxs-lookup"><span data-stu-id="13722-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="13722-104">В этом сценарии вы будете инструктором, который подытвет оценки каждого учащегося по окончании срока обучения.</span><span class="sxs-lookup"><span data-stu-id="13722-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="13722-105">Вы вводя оценки для их назначений и тестов, как вы идете.</span><span class="sxs-lookup"><span data-stu-id="13722-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="13722-106">Теперь настало время определить судьбы учащихся.</span><span class="sxs-lookup"><span data-stu-id="13722-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="13722-107">Вы разработает сценарий, который суммит оценки для каждой категории точки.</span><span class="sxs-lookup"><span data-stu-id="13722-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="13722-108">Затем он назначает каждому учащемуся оценку буквы в зависимости от общей суммы.</span><span class="sxs-lookup"><span data-stu-id="13722-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="13722-109">Чтобы обеспечить точность, вы добавим несколько проверок, чтобы убедиться, что какие-либо отдельные оценки слишком низкие или высокие.</span><span class="sxs-lookup"><span data-stu-id="13722-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="13722-110">Если оценка учащегося меньше нуля или больше возможного значения точки, скрипт будет пометить ячейку с красной заливкой, а не суммой точек этого студента.</span><span class="sxs-lookup"><span data-stu-id="13722-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="13722-111">Это будет четким указанием, какие записи необходимо перепросмотрить.</span><span class="sxs-lookup"><span data-stu-id="13722-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="13722-112">Вы также добавим некоторые базовые форматирования в оценки, чтобы можно было быстро просмотреть верхнюю и нижнюю части класса.</span><span class="sxs-lookup"><span data-stu-id="13722-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="13722-113">Навыки скриптов, охватываемых</span><span class="sxs-lookup"><span data-stu-id="13722-113">Scripting skills covered</span></span>

- <span data-ttu-id="13722-114">Форматирование ячейки</span><span class="sxs-lookup"><span data-stu-id="13722-114">Cell formatting</span></span>
- <span data-ttu-id="13722-115">Проверка ошибок</span><span class="sxs-lookup"><span data-stu-id="13722-115">Error checking</span></span>
- <span data-ttu-id="13722-116">Регулярные выражения</span><span class="sxs-lookup"><span data-stu-id="13722-116">Regular expressions</span></span>
- <span data-ttu-id="13722-117">Условное форматирование</span><span class="sxs-lookup"><span data-stu-id="13722-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="13722-118">Инструкции по настройке</span><span class="sxs-lookup"><span data-stu-id="13722-118">Setup instructions</span></span>

1. <span data-ttu-id="13722-119">Скачайте <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="13722-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="13722-120">Откройте книгу с Excel для интернета.</span><span class="sxs-lookup"><span data-stu-id="13722-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="13722-121">В **вкладке Automate** откройте **все скрипты.**</span><span class="sxs-lookup"><span data-stu-id="13722-121">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="13722-122">В области **задач редактора** кода нажмите **кнопку Новый скрипт** и вклеите следующий скрипт в редактор.</span><span class="sxs-lookup"><span data-stu-id="13722-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="13722-123">Переименуй сценарий в **калькулятор класса и** сохраните его.</span><span class="sxs-lookup"><span data-stu-id="13722-123">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="13722-124">Выполнение скрипта</span><span class="sxs-lookup"><span data-stu-id="13722-124">Running the script</span></span>

<span data-ttu-id="13722-125">Запустите **сценарий калькулятора класса** на единственной таблице.</span><span class="sxs-lookup"><span data-stu-id="13722-125">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="13722-126">Сценарий будет общую оценку и назначить каждому учащемуся оценку буквы.</span><span class="sxs-lookup"><span data-stu-id="13722-126">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="13722-127">Если в отдельных классах имеется больше баллов, чем стоит назначение или тест, то класс обижающих будет отмечен красным, а общее число не вычисляется.</span><span class="sxs-lookup"><span data-stu-id="13722-127">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span> <span data-ttu-id="13722-128">Кроме того, все оценки "A" выделены зеленым цветом, а оценки "D" и "F" выделены желтым цветом.</span><span class="sxs-lookup"><span data-stu-id="13722-128">Also, any 'A' grades are highlighted in green, while 'D' and 'F' grades are highlighted in yellow.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="13722-129">Перед запуском сценария</span><span class="sxs-lookup"><span data-stu-id="13722-129">Before running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="Таблица, в которую показаны строки баллов для учащихся.":::

### <a name="after-running-the-script"></a><span data-ttu-id="13722-131">После запуска скрипта</span><span class="sxs-lookup"><span data-stu-id="13722-131">After running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="Таблица, в которую показаны данные о оценках учащихся с недействительными ячейками в красных итоговых числах для допустимых студенческих строк.":::
