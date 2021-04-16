---
title: Чтение данных книги с помощью сценариев Office в Excel в Интернете
description: Учебник по сценариям Office о чтении данных из книг и их оценке в сценарии.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: d6321cb91a425da3fd45329d5171f1d5694b2b99
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754860"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="38985-103">Чтение данных книги с помощью сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="38985-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="38985-104">В этом учебнике объясняется, как читать данные из книги с помощью сценария Office для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="38985-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="38985-105">Вы напишите новый сценарий, форматирующий банковскую выписку и нормализующий данные в ней.</span><span class="sxs-lookup"><span data-stu-id="38985-105">You'll be writing a new script that formats a bank statement and normalizes the data in that statement.</span></span> <span data-ttu-id="38985-106">В рамках этой очистки данных ваш сценарий будет считывать значения из ячеек транзакций, применять простую формулу к каждому значению и записывать полученный ответ в книгу.</span><span class="sxs-lookup"><span data-stu-id="38985-106">As part of that data clean-up, your script will read values from the transaction cells, apply a simple formula to each value, and write the resulting answer to the workbook.</span></span> <span data-ttu-id="38985-107">Чтение данных из книги позволяет вам автоматизировать некоторые процессы принятия решений в сценарии.</span><span class="sxs-lookup"><span data-stu-id="38985-107">Reading data from the workbook lets you automate some of your decision making processes in the script.</span></span>

> [!TIP]
> <span data-ttu-id="38985-108">Если вы только приступили к работе со сценариями Office, рекомендуем начать с учебника [Запись, редактирование и создание сценариев Office в Excel в Интернете](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="38985-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="38985-109">[Сценарии Office используют TypeScript](../overview/code-editor-environment.md), и этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="38985-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="38985-110">Если вы впервые работаете с JavaScript, рекомендуем начать с [учебника Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="38985-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="38985-111">Предварительные условия</span><span class="sxs-lookup"><span data-stu-id="38985-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## <a name="read-a-cell"></a><span data-ttu-id="38985-112">Чтение ячейки</span><span class="sxs-lookup"><span data-stu-id="38985-112">Read a cell</span></span>

<span data-ttu-id="38985-113">Сценарии, созданные с помощью средства записи действий, могут только записывать информацию в книгу.</span><span class="sxs-lookup"><span data-stu-id="38985-113">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="38985-114">С помощью редактора кода можно редактировать и создавать сценарии, которые также читают данные из книги.</span><span class="sxs-lookup"><span data-stu-id="38985-114">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="38985-115">Давайте создадим сценарий, читающий данные и действующий на основе прочитанного.</span><span class="sxs-lookup"><span data-stu-id="38985-115">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="38985-116">Мы будем работать с примером банковской выписки.</span><span class="sxs-lookup"><span data-stu-id="38985-116">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="38985-117">Эта выписка объединяет чековую выписку и выписку по кредиту.</span><span class="sxs-lookup"><span data-stu-id="38985-117">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="38985-118">К сожалению, изменения баланса в них указываются по-разному.</span><span class="sxs-lookup"><span data-stu-id="38985-118">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="38985-119">В чековой выписке доходы указываются как положительный кредит, а расходы — в виде отрицательного дебета.</span><span class="sxs-lookup"><span data-stu-id="38985-119">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="38985-120">В выписке по кредиту все наоборот.</span><span class="sxs-lookup"><span data-stu-id="38985-120">The credit statement does the opposite.</span></span>

<span data-ttu-id="38985-121">В остальной части учебника мы нормализуем эти данные с помощью сценария.</span><span class="sxs-lookup"><span data-stu-id="38985-121">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="38985-122">Сначала научимся читать данные из книги.</span><span class="sxs-lookup"><span data-stu-id="38985-122">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="38985-123">Создайте лист в книге, которую вы использовали в остальной части учебника.</span><span class="sxs-lookup"><span data-stu-id="38985-123">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="38985-124">Скопируйте следующие данные и вставьте их на новый лист, начиная с ячейки **A1**.</span><span class="sxs-lookup"><span data-stu-id="38985-124">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="38985-125">Дата</span><span class="sxs-lookup"><span data-stu-id="38985-125">Date</span></span> |<span data-ttu-id="38985-126">Счет</span><span class="sxs-lookup"><span data-stu-id="38985-126">Account</span></span> |<span data-ttu-id="38985-127">Описание</span><span class="sxs-lookup"><span data-stu-id="38985-127">Description</span></span> |<span data-ttu-id="38985-128">Дебет</span><span class="sxs-lookup"><span data-stu-id="38985-128">Debit</span></span> |<span data-ttu-id="38985-129">Кредит</span><span class="sxs-lookup"><span data-stu-id="38985-129">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="38985-130">10.10.2019</span><span class="sxs-lookup"><span data-stu-id="38985-130">10/10/2019</span></span> |<span data-ttu-id="38985-131">Чековый</span><span class="sxs-lookup"><span data-stu-id="38985-131">Checking</span></span> |<span data-ttu-id="38985-132">Виноградник Coho</span><span class="sxs-lookup"><span data-stu-id="38985-132">Coho Vineyard</span></span> |<span data-ttu-id="38985-133">–20,05</span><span class="sxs-lookup"><span data-stu-id="38985-133">-20.05</span></span> | |
    |<span data-ttu-id="38985-134">11.10.2019</span><span class="sxs-lookup"><span data-stu-id="38985-134">10/11/2019</span></span> |<span data-ttu-id="38985-135">Кредитный</span><span class="sxs-lookup"><span data-stu-id="38985-135">Credit</span></span> |<span data-ttu-id="38985-136">Телефонная компания</span><span class="sxs-lookup"><span data-stu-id="38985-136">The Phone Company</span></span> |<span data-ttu-id="38985-137">99,95</span><span class="sxs-lookup"><span data-stu-id="38985-137">99.95</span></span> | |
    |<span data-ttu-id="38985-138">13.10.2019</span><span class="sxs-lookup"><span data-stu-id="38985-138">10/13/2019</span></span> |<span data-ttu-id="38985-139">Кредитный</span><span class="sxs-lookup"><span data-stu-id="38985-139">Credit</span></span> |<span data-ttu-id="38985-140">Виноградник Coho</span><span class="sxs-lookup"><span data-stu-id="38985-140">Coho Vineyard</span></span> |<span data-ttu-id="38985-141">154,43</span><span class="sxs-lookup"><span data-stu-id="38985-141">154.43</span></span> | |
    |<span data-ttu-id="38985-142">15.10.2019</span><span class="sxs-lookup"><span data-stu-id="38985-142">10/15/2019</span></span> |<span data-ttu-id="38985-143">Чековый</span><span class="sxs-lookup"><span data-stu-id="38985-143">Checking</span></span> |<span data-ttu-id="38985-144">Внешний депозит</span><span class="sxs-lookup"><span data-stu-id="38985-144">External Deposit</span></span> | |<span data-ttu-id="38985-145">1000</span><span class="sxs-lookup"><span data-stu-id="38985-145">1000</span></span> |
    |<span data-ttu-id="38985-146">20.10.2019</span><span class="sxs-lookup"><span data-stu-id="38985-146">10/20/2019</span></span> |<span data-ttu-id="38985-147">Кредитный</span><span class="sxs-lookup"><span data-stu-id="38985-147">Credit</span></span> |<span data-ttu-id="38985-148">Виноградник Coho — возмещение</span><span class="sxs-lookup"><span data-stu-id="38985-148">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="38985-149">–35,45</span><span class="sxs-lookup"><span data-stu-id="38985-149">-35.45</span></span> |
    |<span data-ttu-id="38985-150">25.10.2019</span><span class="sxs-lookup"><span data-stu-id="38985-150">10/25/2019</span></span> |<span data-ttu-id="38985-151">Чековый</span><span class="sxs-lookup"><span data-stu-id="38985-151">Checking</span></span> |<span data-ttu-id="38985-152">Органическая компания "Лучшее для вас"</span><span class="sxs-lookup"><span data-stu-id="38985-152">Best For You Organics Company</span></span> | <span data-ttu-id="38985-153">–85,64</span><span class="sxs-lookup"><span data-stu-id="38985-153">-85.64</span></span> | |
    |<span data-ttu-id="38985-154">01.11.2019</span><span class="sxs-lookup"><span data-stu-id="38985-154">11/01/2019</span></span> |<span data-ttu-id="38985-155">Чековый</span><span class="sxs-lookup"><span data-stu-id="38985-155">Checking</span></span> |<span data-ttu-id="38985-156">Внешний депозит</span><span class="sxs-lookup"><span data-stu-id="38985-156">External Deposit</span></span> | |<span data-ttu-id="38985-157">1000</span><span class="sxs-lookup"><span data-stu-id="38985-157">1000</span></span> |

3. <span data-ttu-id="38985-158">Откройте **Все сценарии** и выберите **Создать сценарий**.</span><span class="sxs-lookup"><span data-stu-id="38985-158">Open **All Scripts** and select **New Script**.</span></span>
4. <span data-ttu-id="38985-159">Давайте очистим форматирование.</span><span class="sxs-lookup"><span data-stu-id="38985-159">Let's clean up the formatting.</span></span> <span data-ttu-id="38985-160">Это финансовый документ, поэтому изменим числовой формат в столбцах **Дебет** и **Кредит**, чтобы отобразить значения в долларах.</span><span class="sxs-lookup"><span data-stu-id="38985-160">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="38985-161">Также настроим ширину столбца по данным.</span><span class="sxs-lookup"><span data-stu-id="38985-161">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="38985-162">Замените содержимое сценария следующим кодом:</span><span class="sxs-lookup"><span data-stu-id="38985-162">Replace the script contents with the following code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. <span data-ttu-id="38985-163">Теперь прочитаем значение в одном из числовых столбцов.</span><span class="sxs-lookup"><span data-stu-id="38985-163">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="38985-164">Добавьте следующий код в конце сценария (перед закрывающей скобкой `}`):</span><span class="sxs-lookup"><span data-stu-id="38985-164">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="38985-165">Запустите сценарий.</span><span class="sxs-lookup"><span data-stu-id="38985-165">Run the script.</span></span>
7. <span data-ttu-id="38985-166">В консоли должно отображаться следующее: `[Array[1]]`.</span><span class="sxs-lookup"><span data-stu-id="38985-166">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="38985-167">Это не число, так как диапазоны являются двухмерными массивами данных.</span><span class="sxs-lookup"><span data-stu-id="38985-167">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="38985-168">Этот двухмерный диапазон напрямую регистрируется в консоли.</span><span class="sxs-lookup"><span data-stu-id="38985-168">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="38985-169">К счастью, редактор кода позволяет просмотреть содержимое массива.</span><span class="sxs-lookup"><span data-stu-id="38985-169">Luckily, the Code Editor lets you see the contents of the array.</span></span>
8. <span data-ttu-id="38985-170">Когда двухмерный массив регистрируется в консоли, она группирует значения столбцов под каждой строкой.</span><span class="sxs-lookup"><span data-stu-id="38985-170">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="38985-171">Разверните журнал массива, нажав синий треугольник.</span><span class="sxs-lookup"><span data-stu-id="38985-171">Expand the array log by pressing the blue triangle.</span></span>
9. <span data-ttu-id="38985-172">Разверните второй уровень массива, нажав появившийся синий треугольник.</span><span class="sxs-lookup"><span data-stu-id="38985-172">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="38985-173">Должно отобразиться следующее:</span><span class="sxs-lookup"><span data-stu-id="38985-173">You should now see this:</span></span>

    :::image type="content" source="../images/tutorial-4.png" alt-text="Журнал консоли, отображающий результат &quot;–20,05&quot;, размещенный под двумя массивами":::

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="38985-175">Изменение значения ячейки</span><span class="sxs-lookup"><span data-stu-id="38985-175">Modify the value of a cell</span></span>

<span data-ttu-id="38985-176">Теперь, когда мы можем читать данные, воспользуемся ими, чтобы изменить книгу.</span><span class="sxs-lookup"><span data-stu-id="38985-176">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="38985-177">Мы сделаем значение ячейки **D2** положительным с помощью функции `Math.abs`.</span><span class="sxs-lookup"><span data-stu-id="38985-177">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="38985-178">Объект [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) содержит множество функций, к которым имеют доступ сценарии.</span><span class="sxs-lookup"><span data-stu-id="38985-178">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="38985-179">Дополнительные сведения о `Math` и других встроенных объектах можно найти в статье [Использование встроенных объектов JavaScript в сценариях Office](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="38985-179">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="38985-180">Мы будем использовать методы `getValue` и `setValue`, чтобы изменить значение ячейки.</span><span class="sxs-lookup"><span data-stu-id="38985-180">We'll use `getValue` and `setValue` methods to change the value of the cell.</span></span> <span data-ttu-id="38985-181">Эти методы применимы к одной ячейке.</span><span class="sxs-lookup"><span data-stu-id="38985-181">These methods work on a single cell.</span></span> <span data-ttu-id="38985-182">При обработке диапазонов, включающих несколько ячеек, нужно использовать `getValues` и `setValues`.</span><span class="sxs-lookup"><span data-stu-id="38985-182">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span> <span data-ttu-id="38985-183">Добавьте следующий код в конце сценария:</span><span class="sxs-lookup"><span data-stu-id="38985-183">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue() as number);
    range.setValue(positiveValue);
    ```

    > [!NOTE]
    > <span data-ttu-id="38985-184">Мы [приводим](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) возвращаемое значение `range.getValue()` к `number`, используя ключевое слово `as`.</span><span class="sxs-lookup"><span data-stu-id="38985-184">We are [casting](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) the returned value of `range.getValue()` to a `number` by using the `as` keyword.</span></span> <span data-ttu-id="38985-185">Это необходимо, так как диапазон может включать строки, числа или логические значения.</span><span class="sxs-lookup"><span data-stu-id="38985-185">This is necessary because a range could be strings, numbers, or booleans.</span></span> <span data-ttu-id="38985-186">В данном случае нам явно нужно число.</span><span class="sxs-lookup"><span data-stu-id="38985-186">In this instance, we explicitly need a number.</span></span>

2. <span data-ttu-id="38985-187">Значение ячейки **D2** теперь должно быть положительным.</span><span class="sxs-lookup"><span data-stu-id="38985-187">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="38985-188">Изменение значений столбца</span><span class="sxs-lookup"><span data-stu-id="38985-188">Modify the values of a column</span></span>

<span data-ttu-id="38985-189">Теперь, когда вы знаете, как читать и записывать данные в одной ячейке, давайте обобщим сценарий для работы со всеми столбцами **Дебет** и **Кредит**.</span><span class="sxs-lookup"><span data-stu-id="38985-189">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="38985-190">Удалите код, влияющий только на одну ячейку (предыдущий код с абсолютным значением), чтобы ваш сценарий выглядел следующим образом:</span><span class="sxs-lookup"><span data-stu-id="38985-190">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. <span data-ttu-id="38985-191">Добавьте цикл в конце сценария, выполняющий итерацию в строках двух последних столбцов.</span><span class="sxs-lookup"><span data-stu-id="38985-191">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="38985-192">Для каждой ячейки сценарий устанавливает текущее абсолютное значение.</span><span class="sxs-lookup"><span data-stu-id="38985-192">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="38985-193">Обратите внимание, что массив, определяющий расположения ячеек, отсчитывается от нуля.</span><span class="sxs-lookup"><span data-stu-id="38985-193">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="38985-194">Это означает, что ячейка **A1** имеет значение `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="38985-194">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3] as number);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4] as number);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    <span data-ttu-id="38985-195">Эта часть сценария выполняет несколько важных задач.</span><span class="sxs-lookup"><span data-stu-id="38985-195">This portion of the script does several important tasks.</span></span> <span data-ttu-id="38985-196">Сначала она получает значения и количество строк используемого диапазона.</span><span class="sxs-lookup"><span data-stu-id="38985-196">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="38985-197">Это позволяет просматривать значения и определять момент остановки.</span><span class="sxs-lookup"><span data-stu-id="38985-197">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="38985-198">Затем выполняется итерация в используемом диапазоне с проверкой каждой ячейки в столбцах **Дебет** или **Кредит**.</span><span class="sxs-lookup"><span data-stu-id="38985-198">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="38985-199">Наконец, если значение в ячейке не равно 0, оно заменяется абсолютным значением.</span><span class="sxs-lookup"><span data-stu-id="38985-199">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="38985-200">Мы избегаем использования нулей, поэтому можно оставить пустые ячейки неизменными.</span><span class="sxs-lookup"><span data-stu-id="38985-200">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="38985-201">Запустите сценарий.</span><span class="sxs-lookup"><span data-stu-id="38985-201">Run the script.</span></span>

    <span data-ttu-id="38985-202">Теперь банковская выписка должна выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="38985-202">Your banking statement should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-5.png" alt-text="Таблица, показывающая банковскую выписку в виде отформатированной таблицы только с положительными значениями.":::

## <a name="next-steps"></a><span data-ttu-id="38985-204">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="38985-204">Next steps</span></span>

<span data-ttu-id="38985-205">Откройте редактор кода и попробуйте некоторые [примеры сценариев Office в Excel в Интернете](../resources/excel-samples.md).</span><span class="sxs-lookup"><span data-stu-id="38985-205">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="38985-206">Дополнительные сведения о создании сценариев Office доступны в статье [Основные сведения о сценариях Office в Excel в Интернете](../develop/scripting-fundamentals.md).</span><span class="sxs-lookup"><span data-stu-id="38985-206">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>

<span data-ttu-id="38985-207">В следующем наборе учебников по сценариям Office рассматривается использование сценариев Office с помощью Power Automate.</span><span class="sxs-lookup"><span data-stu-id="38985-207">The next series of Office Scripts tutorials focus on using Office Scripts with Power Automate.</span></span> <span data-ttu-id="38985-208">Узнайте больше о преимуществах сочетания двух платформ в статье [Выполнение сценариев Office с помощью Power Automate](../develop/power-automate-integration.md) или ознакомьтесь с учебником [Вызов сценариев из активированного вручную потока Power Automate](excel-power-automate-manual.md), чтобы создать поток Power Automate, использующий сценарий Office.</span><span class="sxs-lookup"><span data-stu-id="38985-208">Learn more about the advantages combining the two platforms in [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) or try the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial to create a Power Automate flow that uses an Office Script.</span></span>
