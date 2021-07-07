---
title: Запись, редактирование и создание сценариев Office в Excel в Интернете
description: Учебник с основными сведениями о сценариях Office, включая запись сценариев с помощью средства записи действий и запись данных в книгу.
ms.date: 05/23/2021
localization_priority: Priority
ms.openlocfilehash: 6bcf603211aa07920e99178c35c6f405224c29bd
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313927"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="4d1be-103">Запись, редактирование и создание сценариев Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="4d1be-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="4d1be-104">В этом учебнике вы ознакомитесь с основами записи, редактирования и создания сценария Office для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="4d1be-104">This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span> <span data-ttu-id="4d1be-105">Вы запишите сценарий, применяющий форматирование к листу продаж.</span><span class="sxs-lookup"><span data-stu-id="4d1be-105">You'll record a script that applies some formatting to a sales record worksheet.</span></span> <span data-ttu-id="4d1be-106">После этого вы измените записанный сценарий, чтобы применить дополнительное форматирование, создать таблицу и отсортировать ее.</span><span class="sxs-lookup"><span data-stu-id="4d1be-106">You'll then edit the recorded script to apply more formatting, create a table, and sort that table.</span></span> <span data-ttu-id="4d1be-107">Эта шаблон записи с последующим изменением является важным инструментом для просмотра ваших действий Excel в виде кода.</span><span class="sxs-lookup"><span data-stu-id="4d1be-107">This record-then-edit pattern is an important tool to see what your Excel actions look like as code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="4d1be-108">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="4d1be-108">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="4d1be-109">Этот учебник предназначен для пользователей с начальным и средним уровнем знаний по JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="4d1be-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="4d1be-110">Если вы впервые работаете с JavaScript, рекомендуем начать с [учебника Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="4d1be-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="4d1be-111">Чтобы получить дополнительные сведения о среде сценариев, ознакомьтесь со статьей [Среда редактора кода сценариев Office](../overview/code-editor-environment.md).</span><span class="sxs-lookup"><span data-stu-id="4d1be-111">Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="4d1be-112">Добавление данных и запись простого сценария</span><span class="sxs-lookup"><span data-stu-id="4d1be-112">Add data and record a basic script</span></span>

<span data-ttu-id="4d1be-113">Сначала нам потребуются некоторые данные и небольшой начальный сценарий.</span><span class="sxs-lookup"><span data-stu-id="4d1be-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="4d1be-114">Создайте книгу в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="4d1be-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="4d1be-115">Скопируйте следующие данные о продаже фруктов и вставьте их на лист, начиная с ячейки **A1**.</span><span class="sxs-lookup"><span data-stu-id="4d1be-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="4d1be-116">Фрукты</span><span class="sxs-lookup"><span data-stu-id="4d1be-116">Fruit</span></span> |<span data-ttu-id="4d1be-117">2018</span><span class="sxs-lookup"><span data-stu-id="4d1be-117">2018</span></span> |<span data-ttu-id="4d1be-118">2019</span><span class="sxs-lookup"><span data-stu-id="4d1be-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="4d1be-119">Апельсины</span><span class="sxs-lookup"><span data-stu-id="4d1be-119">Oranges</span></span> |<span data-ttu-id="4d1be-120">1000</span><span class="sxs-lookup"><span data-stu-id="4d1be-120">1000</span></span> |<span data-ttu-id="4d1be-121">1200</span><span class="sxs-lookup"><span data-stu-id="4d1be-121">1200</span></span> |
    |<span data-ttu-id="4d1be-122">Лимоны</span><span class="sxs-lookup"><span data-stu-id="4d1be-122">Lemons</span></span> |<span data-ttu-id="4d1be-123">800</span><span class="sxs-lookup"><span data-stu-id="4d1be-123">800</span></span> |<span data-ttu-id="4d1be-124">900</span><span class="sxs-lookup"><span data-stu-id="4d1be-124">900</span></span> |
    |<span data-ttu-id="4d1be-125">Лаймы</span><span class="sxs-lookup"><span data-stu-id="4d1be-125">Limes</span></span> |<span data-ttu-id="4d1be-126">600</span><span class="sxs-lookup"><span data-stu-id="4d1be-126">600</span></span> |<span data-ttu-id="4d1be-127">500</span><span class="sxs-lookup"><span data-stu-id="4d1be-127">500</span></span> |
    |<span data-ttu-id="4d1be-128">Грейпфруты</span><span class="sxs-lookup"><span data-stu-id="4d1be-128">Grapefruits</span></span> |<span data-ttu-id="4d1be-129">900</span><span class="sxs-lookup"><span data-stu-id="4d1be-129">900</span></span> |<span data-ttu-id="4d1be-130">700</span><span class="sxs-lookup"><span data-stu-id="4d1be-130">700</span></span> |

3. <span data-ttu-id="4d1be-131">Откройте вкладку **Автоматизация**. Если вы не видите вкладку **Автоматизация**, проверьте переполнение ленты, нажав стрелку раскрывающегося списка.</span><span class="sxs-lookup"><span data-stu-id="4d1be-131">Open the **Automate** tab. If you don't see the **Automate** tab, check the ribbon overflow by selecting the drop-down arrow.</span></span> <span data-ttu-id="4d1be-132">Если нужного элемента по-прежнему нет, выполните рекомендации из статьи [Устранение неполадок в сценариях Office](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span><span class="sxs-lookup"><span data-stu-id="4d1be-132">If it's still not there, follow the advice in the article [Troubleshoot Office Scripts](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span></span>
4. <span data-ttu-id="4d1be-133">Нажмите кнопку **Записать действия**.</span><span class="sxs-lookup"><span data-stu-id="4d1be-133">Select the **Record Actions** button.</span></span>
5. <span data-ttu-id="4d1be-134">Выделите ячейки **A2:C2** (строка "Апельсины") и установите оранжевый цвет заливки.</span><span class="sxs-lookup"><span data-stu-id="4d1be-134">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="4d1be-135">Чтобы остановить запись, нажмите кнопку **Остановить**.</span><span class="sxs-lookup"><span data-stu-id="4d1be-135">Stop the recording by selecting the **Stop** button.</span></span>

    <span data-ttu-id="4d1be-136">Ваш лист должен выглядеть, как показано ниже (не волнуйтесь, если цвет отличается):</span><span class="sxs-lookup"><span data-stu-id="4d1be-136">Your worksheet should look like this (don't worry if the color is different):</span></span>

    :::image type="content" source="../images/tutorial-1.png" alt-text="Лист, показывающий строку данных о продажах фруктов, причем строка &quot;Апельсины&quot; выделена оранжевым цветом.":::

## <a name="edit-an-existing-script"></a><span data-ttu-id="4d1be-138">Редактирование существующего сценария</span><span class="sxs-lookup"><span data-stu-id="4d1be-138">Edit an existing script</span></span>

<span data-ttu-id="4d1be-139">Предыдущий сценарий окрасил строку "Апельсины" в оранжевый цвет.</span><span class="sxs-lookup"><span data-stu-id="4d1be-139">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="4d1be-140">Давайте добавим желтый цвет для строки "Лимоны".</span><span class="sxs-lookup"><span data-stu-id="4d1be-140">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="4d1be-141">В открывшейся области **Сведения** нажмите кнопку **Изменить**.</span><span class="sxs-lookup"><span data-stu-id="4d1be-141">From the now-open **Details** pane, select the **Edit** button.</span></span>
2. <span data-ttu-id="4d1be-142">Должен отобразиться примерно такой код:</span><span class="sxs-lookup"><span data-stu-id="4d1be-142">You should see something similar to this code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    <span data-ttu-id="4d1be-143">Этот код получает текущий лист из книги.</span><span class="sxs-lookup"><span data-stu-id="4d1be-143">This code gets the current worksheet from the workbook.</span></span> <span data-ttu-id="4d1be-144">Затем он настраивает цвет заливки диапазона **A2:C2**.</span><span class="sxs-lookup"><span data-stu-id="4d1be-144">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="4d1be-145">Диапазоны — это фундаментальная часть сценариев Office в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="4d1be-145">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="4d1be-146">Диапазон — это непрерывный прямоугольный блок ячеек, содержащий значения, формулы и форматирование.</span><span class="sxs-lookup"><span data-stu-id="4d1be-146">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="4d1be-147">Они представляют собой базовую структуру ячеек, в которой можно выполнять большинство задач сценариев.</span><span class="sxs-lookup"><span data-stu-id="4d1be-147">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

3. <span data-ttu-id="4d1be-148">Добавьте следующую строку в конце сценария (между местом настройки значения `color` и закрывающей скобкой `}`):</span><span class="sxs-lookup"><span data-stu-id="4d1be-148">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. <span data-ttu-id="4d1be-149">Протестируйте сценарий, нажав **Запустить**.</span><span class="sxs-lookup"><span data-stu-id="4d1be-149">Test the script by selecting **Run**.</span></span> <span data-ttu-id="4d1be-150">Книга должна выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="4d1be-150">Your workbook should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-2.png" alt-text="Лист, показывающий строку данных о продажах фруктов, в которой строка &quot;Апельсины&quot; выделена оранжевым цветом, а строка &quot;Лимоны&quot; — желтым цветом.":::

## <a name="create-a-table"></a><span data-ttu-id="4d1be-152">Создание таблицы</span><span class="sxs-lookup"><span data-stu-id="4d1be-152">Create a table</span></span>

<span data-ttu-id="4d1be-153">Давайте преобразуем эти данные продаж фруктов в таблицу.</span><span class="sxs-lookup"><span data-stu-id="4d1be-153">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="4d1be-154">Мы воспользуемся собственным сценарием для всего процесса.</span><span class="sxs-lookup"><span data-stu-id="4d1be-154">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="4d1be-155">Добавьте следующую строку в конце сценария (перед закрывающей скобкой `}`):</span><span class="sxs-lookup"><span data-stu-id="4d1be-155">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. <span data-ttu-id="4d1be-156">Этот вызов возвращает объект `Table`.</span><span class="sxs-lookup"><span data-stu-id="4d1be-156">That call returns a `Table` object.</span></span> <span data-ttu-id="4d1be-157">Воспользуемся этой таблицей, чтобы отсортировать данные.</span><span class="sxs-lookup"><span data-stu-id="4d1be-157">Let's use that table to sort the data.</span></span> <span data-ttu-id="4d1be-158">Отсортируем данные по возрастанию на основе значений в столбце "Фрукты".</span><span class="sxs-lookup"><span data-stu-id="4d1be-158">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="4d1be-159">Добавьте следующую строку после создания таблицы:</span><span class="sxs-lookup"><span data-stu-id="4d1be-159">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="4d1be-160">Ваш сценарий должен выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="4d1be-160">Your script should look like this:</span></span>

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

    <span data-ttu-id="4d1be-161">В таблицах есть объект `TableSort`, доступный с помощью метода `Table.getSort`.</span><span class="sxs-lookup"><span data-stu-id="4d1be-161">Tables have a `TableSort` object, accessed through the `Table.getSort` method.</span></span> <span data-ttu-id="4d1be-162">Вы можете применить условия сортировки к этому объекту.</span><span class="sxs-lookup"><span data-stu-id="4d1be-162">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="4d1be-163">Метод `apply` использует массив объектов `SortField`.</span><span class="sxs-lookup"><span data-stu-id="4d1be-163">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="4d1be-164">В этом случае у нас есть только одно условие сортировки, поэтому мы используем только один параметр `SortField`.</span><span class="sxs-lookup"><span data-stu-id="4d1be-164">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="4d1be-165">`key: 0` присваивает столбцу со значениями, определяющими сортировку, значение "0" (это первый столбец в таблице, в данном случае: **A**).</span><span class="sxs-lookup"><span data-stu-id="4d1be-165">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="4d1be-166">`ascending: true` сортирует данные по возрастанию (вместо порядка по убыванию).</span><span class="sxs-lookup"><span data-stu-id="4d1be-166">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="4d1be-p111">Запустите сценарий. Вы увидите следующую таблицу:</span><span class="sxs-lookup"><span data-stu-id="4d1be-p111">Run the script. You should see a table like this:</span></span>

    :::image type="content" source="../images/tutorial-3.png" alt-text="лист с таблицей продаж отсортированных фруктов.":::

    > [!NOTE]
    > <span data-ttu-id="4d1be-170">При повторном запуске сценария возникнет ошибка.</span><span class="sxs-lookup"><span data-stu-id="4d1be-170">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="4d1be-171">Это связано с тем, что вы не можете создать таблицу поверх другой таблицы.</span><span class="sxs-lookup"><span data-stu-id="4d1be-171">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="4d1be-172">Однако вы можете запустить этот сценарий на другом листе или в другой книге.</span><span class="sxs-lookup"><span data-stu-id="4d1be-172">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="4d1be-173">Повторный запуск сценария</span><span class="sxs-lookup"><span data-stu-id="4d1be-173">Re-run the script</span></span>

1. <span data-ttu-id="4d1be-174">Создайте лист в текущей книге.</span><span class="sxs-lookup"><span data-stu-id="4d1be-174">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="4d1be-175">Скопируйте данные фруктов из начала учебника и вставьте их на новый лист, начиная с ячейки **A1**.</span><span class="sxs-lookup"><span data-stu-id="4d1be-175">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="4d1be-176">Запустите сценарий.</span><span class="sxs-lookup"><span data-stu-id="4d1be-176">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="4d1be-177">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="4d1be-177">Next steps</span></span>

<span data-ttu-id="4d1be-178">Выполните инструкции учебника [Чтение данных книги с помощью сценариев Office в Excel в Интернете](excel-read-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="4d1be-178">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="4d1be-179">С его помощью вы научитесь читать данные из книги с помощью сценариев Office.</span><span class="sxs-lookup"><span data-stu-id="4d1be-179">It teaches you how to read data from a workbook with an Office Script.</span></span>
