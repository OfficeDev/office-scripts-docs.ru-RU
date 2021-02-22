---
title: Запись, редактирование и создание сценариев Office в Excel в Интернете
description: Учебник с основными сведениями о сценариях Office, включая запись сценариев с помощью средства записи действий и запись данных в книгу.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: fa2fec54cdb0289590868ee6848a942401f21665
ms.sourcegitcommit: 9df67e007ddbfec79a7360df9f4ea5ac6c86fb08
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49772980"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="c826b-103">Запись, редактирование и создание сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="c826b-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="c826b-104">В этом учебнике вы ознакомитесь с основами записи, редактирования и создания сценария Office для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="c826b-104">This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span> <span data-ttu-id="c826b-105">Вы запишите сценарий, применяющий форматирование к листу продаж.</span><span class="sxs-lookup"><span data-stu-id="c826b-105">You'll record a script that applies some formatting to a sales record worksheet.</span></span> <span data-ttu-id="c826b-106">После этого вы измените записанный сценарий, чтобы применить дополнительное форматирование, создать таблицу и отсортировать ее.</span><span class="sxs-lookup"><span data-stu-id="c826b-106">You'll then edit the recorded script to apply more formatting, create a table, and sort that table.</span></span> <span data-ttu-id="c826b-107">Эта шаблон записи с последующим изменением является важным инструментом для просмотра ваших действий Excel в виде кода.</span><span class="sxs-lookup"><span data-stu-id="c826b-107">This record-then-edit pattern is an important tool to see what your Excel actions look like as code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c826b-108">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="c826b-108">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="c826b-109">Этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="c826b-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="c826b-110">Если вы впервые работаете с JavaScript, рекомендуем начать с [учебника Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="c826b-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="c826b-111">Чтобы получить дополнительные сведения о среде сценариев, ознакомьтесь со статьей [Среда редактора кода сценариев Office](../overview/code-editor-environment.md).</span><span class="sxs-lookup"><span data-stu-id="c826b-111">Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="c826b-112">Добавление данных и запись простого сценария</span><span class="sxs-lookup"><span data-stu-id="c826b-112">Add data and record a basic script</span></span>

<span data-ttu-id="c826b-113">Сначала нам потребуются некоторые данные и небольшой начальный сценарий.</span><span class="sxs-lookup"><span data-stu-id="c826b-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="c826b-114">Создайте книгу в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="c826b-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="c826b-115">Скопируйте следующие данные о продаже фруктов и вставьте их на лист, начиная с ячейки **A1**.</span><span class="sxs-lookup"><span data-stu-id="c826b-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="c826b-116">Фрукты</span><span class="sxs-lookup"><span data-stu-id="c826b-116">Fruit</span></span> |<span data-ttu-id="c826b-117">2018</span><span class="sxs-lookup"><span data-stu-id="c826b-117">2018</span></span> |<span data-ttu-id="c826b-118">2019</span><span class="sxs-lookup"><span data-stu-id="c826b-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="c826b-119">Апельсины</span><span class="sxs-lookup"><span data-stu-id="c826b-119">Oranges</span></span> |<span data-ttu-id="c826b-120">1000</span><span class="sxs-lookup"><span data-stu-id="c826b-120">1000</span></span> |<span data-ttu-id="c826b-121">1200</span><span class="sxs-lookup"><span data-stu-id="c826b-121">1200</span></span> |
    |<span data-ttu-id="c826b-122">Лимоны</span><span class="sxs-lookup"><span data-stu-id="c826b-122">Lemons</span></span> |<span data-ttu-id="c826b-123">800</span><span class="sxs-lookup"><span data-stu-id="c826b-123">800</span></span> |<span data-ttu-id="c826b-124">900</span><span class="sxs-lookup"><span data-stu-id="c826b-124">900</span></span> |
    |<span data-ttu-id="c826b-125">Лаймы</span><span class="sxs-lookup"><span data-stu-id="c826b-125">Limes</span></span> |<span data-ttu-id="c826b-126">600</span><span class="sxs-lookup"><span data-stu-id="c826b-126">600</span></span> |<span data-ttu-id="c826b-127">500</span><span class="sxs-lookup"><span data-stu-id="c826b-127">500</span></span> |
    |<span data-ttu-id="c826b-128">Грейпфруты</span><span class="sxs-lookup"><span data-stu-id="c826b-128">Grapefruits</span></span> |<span data-ttu-id="c826b-129">900</span><span class="sxs-lookup"><span data-stu-id="c826b-129">900</span></span> |<span data-ttu-id="c826b-130">700</span><span class="sxs-lookup"><span data-stu-id="c826b-130">700</span></span> |

3. <span data-ttu-id="c826b-131">Откройте вкладку **Автоматизировать**. Если вы не видите вкладку **Автоматизировать**, проверьте переполнение ленты, нажав стрелку раскрывающегося списка.</span><span class="sxs-lookup"><span data-stu-id="c826b-131">Open the **Automate** tab. If you do not see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.</span></span>
4. <span data-ttu-id="c826b-132">Нажмите кнопку **Записать действия**.</span><span class="sxs-lookup"><span data-stu-id="c826b-132">Press the **Record Actions** button.</span></span>
5. <span data-ttu-id="c826b-133">Выделите ячейки **A2:C2** (строка "Апельсины") и установите оранжевый цвет заливки.</span><span class="sxs-lookup"><span data-stu-id="c826b-133">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="c826b-134">Чтобы остановить запись, нажмите кнопку **Остановить**.</span><span class="sxs-lookup"><span data-stu-id="c826b-134">Stop the recording by pressing the **Stop** button.</span></span>
7. <span data-ttu-id="c826b-135">Введите в поле **Имя сценария** запоминающееся имя.</span><span class="sxs-lookup"><span data-stu-id="c826b-135">Fill in the **Script Name** field with a memorable name.</span></span>
8. <span data-ttu-id="c826b-136">*Необязательно.* Введите в поле **Описание** понятное описание.</span><span class="sxs-lookup"><span data-stu-id="c826b-136">*Optional:* Fill in the **Description** field with a meaningful description.</span></span> <span data-ttu-id="c826b-137">Оно используется для предоставления контекста в отношении действий сценария.</span><span class="sxs-lookup"><span data-stu-id="c826b-137">This is used to provide context as to what the script does.</span></span> <span data-ttu-id="c826b-138">Для этого учебника можно использовать описание "Цветовая кодировка строк таблицы".</span><span class="sxs-lookup"><span data-stu-id="c826b-138">For the tutorial, you can use "Color-codes rows of a table".</span></span>

   > [!TIP]
   > <span data-ttu-id="c826b-139">Вы можете изменить описание сценария позже в области **Сведения о сценарии**, расположенной в меню **...** редактора кода.</span><span class="sxs-lookup"><span data-stu-id="c826b-139">You can edit a script's description later from the **Script Details** pane, which is located under the Code Editor's **...** menu.</span></span>

9. <span data-ttu-id="c826b-140">Сохраните сценарий, нажав кнопку **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="c826b-140">Save the script by pressing the **Save** button.</span></span>

    <span data-ttu-id="c826b-141">Ваш лист должен выглядеть, как показано ниже (не волнуйтесь, если цвет отличается):</span><span class="sxs-lookup"><span data-stu-id="c826b-141">Your worksheet should look like this (don't worry if the color is different):</span></span>

    ![Строка данных о продажах фруктов с выделенной оранжевым цветом строкой "Апельсины"](../images/tutorial-1.png)

## <a name="edit-an-existing-script"></a><span data-ttu-id="c826b-143">Редактирование существующего сценария</span><span class="sxs-lookup"><span data-stu-id="c826b-143">Edit an existing script</span></span>

<span data-ttu-id="c826b-144">Предыдущий сценарий окрасил строку "Апельсины" в оранжевый цвет.</span><span class="sxs-lookup"><span data-stu-id="c826b-144">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="c826b-145">Давайте добавим желтый цвет для строки "Лимоны".</span><span class="sxs-lookup"><span data-stu-id="c826b-145">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="c826b-146">В открывшейся области **Сведения** нажмите кнопку **Изменить**.</span><span class="sxs-lookup"><span data-stu-id="c826b-146">From the now-open **Details** pane, press the **Edit** button.</span></span>
2. <span data-ttu-id="c826b-147">Должен отобразиться примерно такой код:</span><span class="sxs-lookup"><span data-stu-id="c826b-147">You should see something similar to this code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    <span data-ttu-id="c826b-148">Этот код получает текущий лист из книги.</span><span class="sxs-lookup"><span data-stu-id="c826b-148">This code gets the current worksheet from the workbook.</span></span> <span data-ttu-id="c826b-149">Затем он настраивает цвет заливки диапазона **A2:C2**.</span><span class="sxs-lookup"><span data-stu-id="c826b-149">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="c826b-150">Диапазоны — это фундаментальная часть сценариев Office в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="c826b-150">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="c826b-151">Диапазон — это непрерывный прямоугольный блок ячеек, содержащий значения, формулы и форматирование.</span><span class="sxs-lookup"><span data-stu-id="c826b-151">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="c826b-152">Они представляют собой базовую структуру ячеек, в которой можно выполнять большинство задач сценариев.</span><span class="sxs-lookup"><span data-stu-id="c826b-152">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

3. <span data-ttu-id="c826b-153">Добавьте следующую строку в конце сценария (между местом настройки значения `color` и закрывающей скобкой `}`):</span><span class="sxs-lookup"><span data-stu-id="c826b-153">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. <span data-ttu-id="c826b-154">Протестируйте сценарий, нажав **Запустить**.</span><span class="sxs-lookup"><span data-stu-id="c826b-154">Test the script by pressing **Run**.</span></span> <span data-ttu-id="c826b-155">Книга должна выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="c826b-155">Your workbook should now look like this:</span></span>

    ![Строка данных о продажах фруктов с выделенной оранжевым цветом строкой "Апельсины" и выделенной желтым цветом строкой "Лимоны"](../images/tutorial-2.png)

## <a name="create-a-table"></a><span data-ttu-id="c826b-157">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="c826b-157">Create a table</span></span>

<span data-ttu-id="c826b-158">Давайте преобразуем эти данные продаж фруктов в таблицу.</span><span class="sxs-lookup"><span data-stu-id="c826b-158">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="c826b-159">Мы воспользуемся собственным сценарием для всего процесса.</span><span class="sxs-lookup"><span data-stu-id="c826b-159">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="c826b-160">Добавьте следующую строку в конце сценария (перед закрывающей скобкой `}`):</span><span class="sxs-lookup"><span data-stu-id="c826b-160">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. <span data-ttu-id="c826b-161">Этот вызов возвращает объект `Table`.</span><span class="sxs-lookup"><span data-stu-id="c826b-161">That call returns a `Table` object.</span></span> <span data-ttu-id="c826b-162">Воспользуемся этой таблицей, чтобы отсортировать данные.</span><span class="sxs-lookup"><span data-stu-id="c826b-162">Let's use that table to sort the data.</span></span> <span data-ttu-id="c826b-163">Отсортируем данные по возрастанию на основе значений в столбце "Фрукты".</span><span class="sxs-lookup"><span data-stu-id="c826b-163">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="c826b-164">Добавьте следующую строку после создания таблицы:</span><span class="sxs-lookup"><span data-stu-id="c826b-164">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="c826b-165">Ваш сценарий должен выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="c826b-165">Your script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    <span data-ttu-id="c826b-166">В таблицах есть объект `TableSort`, доступный с помощью метода `Table.getSort`.</span><span class="sxs-lookup"><span data-stu-id="c826b-166">Tables have a `TableSort` object, accessed through the `Table.getSort` method.</span></span> <span data-ttu-id="c826b-167">Вы можете применить условия сортировки к этому объекту.</span><span class="sxs-lookup"><span data-stu-id="c826b-167">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="c826b-168">Метод `apply` использует массив объектов `SortField`.</span><span class="sxs-lookup"><span data-stu-id="c826b-168">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="c826b-169">В этом случае у нас есть только одно условие сортировки, поэтому мы используем только один параметр `SortField`.</span><span class="sxs-lookup"><span data-stu-id="c826b-169">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="c826b-170">`key: 0` присваивает столбцу со значениями, определяющими сортировку, значение "0" (это первый столбец в таблице, в данном случае: **A**).</span><span class="sxs-lookup"><span data-stu-id="c826b-170">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="c826b-171">`ascending: true` сортирует данные по возрастанию (вместо порядка по убыванию).</span><span class="sxs-lookup"><span data-stu-id="c826b-171">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="c826b-172">Запустите сценарий.</span><span class="sxs-lookup"><span data-stu-id="c826b-172">Run the script.</span></span> <span data-ttu-id="c826b-173">Вы увидите следующую таблицу:</span><span class="sxs-lookup"><span data-stu-id="c826b-173">You should see a table like this:</span></span>

    ![Отсортированная таблица продаж фруктов](../images/tutorial-3.png)

    > [!NOTE]
    > <span data-ttu-id="c826b-175">При повторном запуске сценария возникнет ошибка.</span><span class="sxs-lookup"><span data-stu-id="c826b-175">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="c826b-176">Это связано с тем, что вы не можете создать таблицу поверх другой таблицы.</span><span class="sxs-lookup"><span data-stu-id="c826b-176">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="c826b-177">Однако вы можете запустить этот сценарий на другом листе или в другой книге.</span><span class="sxs-lookup"><span data-stu-id="c826b-177">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="c826b-178">Повторный запуск сценария</span><span class="sxs-lookup"><span data-stu-id="c826b-178">Re-run the script</span></span>

1. <span data-ttu-id="c826b-179">Создайте лист в текущей книге.</span><span class="sxs-lookup"><span data-stu-id="c826b-179">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="c826b-180">Скопируйте данные фруктов из начала учебника и вставьте их на новый лист, начиная с ячейки **A1**.</span><span class="sxs-lookup"><span data-stu-id="c826b-180">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="c826b-181">Запустите сценарий.</span><span class="sxs-lookup"><span data-stu-id="c826b-181">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="c826b-182">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="c826b-182">Next steps</span></span>

<span data-ttu-id="c826b-183">Выполните инструкции учебника [Чтение данных книги с помощью сценариев Office в Excel в Интернете](excel-read-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="c826b-183">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="c826b-184">С его помощью вы научитесь читать данные из книги с помощью сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="c826b-184">It teaches you how to read data from a workbook with an Office Script.</span></span>
