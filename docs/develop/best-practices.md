---
title: Рекомендации по сценариям Office
description: Как предотвратить общие проблемы и написать надежные Office, которые могут обрабатывать неожиданные входные данные или данные.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546033"
---
# <a name="best-practices-in-office-scripts"></a><span data-ttu-id="5b509-103">Рекомендации по сценариям Office</span><span class="sxs-lookup"><span data-stu-id="5b509-103">Best practices in Office Scripts</span></span>

<span data-ttu-id="5b509-104">Эти шаблоны и практики разработаны, чтобы помочь вашим сценариям успешно работать каждый раз.</span><span class="sxs-lookup"><span data-stu-id="5b509-104">These patterns and practices are designed to help your scripts run successfully every time.</span></span> <span data-ttu-id="5b509-105">Используйте их, чтобы избежать распространенных ловушек при запуске автоматизации рабочего Excel процесса.</span><span class="sxs-lookup"><span data-stu-id="5b509-105">Use them to avoid common pitfalls as you start automating your Excel workflow.</span></span>

## <a name="verify-an-object-is-present"></a><span data-ttu-id="5b509-106">Проверка присутствуют на объекте</span><span class="sxs-lookup"><span data-stu-id="5b509-106">Verify an object is present</span></span>

<span data-ttu-id="5b509-107">Сценарии часто полагаются на определенный лист или таблицу, присутствуют в рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="5b509-107">Scripts often rely on a certain worksheet or table being present in the workbook.</span></span> <span data-ttu-id="5b509-108">Тем не менее, они могут быть переименованы или удалены между запускается сценарий.</span><span class="sxs-lookup"><span data-stu-id="5b509-108">However, they might get renamed or removed between script runs.</span></span> <span data-ttu-id="5b509-109">Проверяя, существуют ли эти таблицы или листы перед вызовом методов на них, вы можете убедиться, что сценарий не заканчивается внезапно.</span><span class="sxs-lookup"><span data-stu-id="5b509-109">By checking if those tables or worksheets exist before calling methods on them, you can make sure the script doesn't end abruptly.</span></span>

<span data-ttu-id="5b509-110">Следующий пример кода проверяет, присутствует ли лист «Индекс» в рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="5b509-110">The following sample code checks if the "Index" worksheet is present in the workbook.</span></span> <span data-ttu-id="5b509-111">Если лист присутствует, скрипт получает диапазон и продолжается.</span><span class="sxs-lookup"><span data-stu-id="5b509-111">If the worksheet is present, the script gets a range and proceeds.</span></span> <span data-ttu-id="5b509-112">Если его нет, скрипт регистрирует пользовательское сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="5b509-112">If it isn't present, the script logs a custom error message.</span></span>

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

<span data-ttu-id="5b509-113">Оператор TypeScript `?` проверяет, существует ли объект перед вызовом метода.</span><span class="sxs-lookup"><span data-stu-id="5b509-113">The TypeScript `?` operator checks if the object exists before calling a method.</span></span> <span data-ttu-id="5b509-114">Это может сделать ваш код более упорядоченным, если вам не нужно делать ничего особенного, когда объект не существует.</span><span class="sxs-lookup"><span data-stu-id="5b509-114">This can make your code more streamlined if you don't need to do anything special when the object doesn't exist.</span></span>

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a><span data-ttu-id="5b509-115">Проверка данных и состояние рабочей книги в первую очередь</span><span class="sxs-lookup"><span data-stu-id="5b509-115">Validate data and workbook state first</span></span>

<span data-ttu-id="5b509-116">Убедитесь, что все ваши листы, таблицы, формы и другие объекты присутствуют перед работой над данными.</span><span class="sxs-lookup"><span data-stu-id="5b509-116">Make sure all your worksheets, tables, shapes, and other objects are present before working on the data.</span></span> <span data-ttu-id="5b509-117">Используя предыдущую схему, проверьте, все ли в рабочей книге и соответствует вашим ожиданиям.</span><span class="sxs-lookup"><span data-stu-id="5b509-117">Using the previous pattern, check to see if everything is in the workbook and matches your expectations.</span></span> <span data-ttu-id="5b509-118">Это делается до того, как будут написаны какие-либо данные, и ваш скрипт не оставит трудовую книжку в частичном состоянии.</span><span class="sxs-lookup"><span data-stu-id="5b509-118">Doing this before any data is written ensures your script doesn't leave the workbook in a partial state.</span></span>

<span data-ttu-id="5b509-119">Следующий скрипт требует, чтобы присутствовали две таблицы под названием "Table1" и "Table2".</span><span class="sxs-lookup"><span data-stu-id="5b509-119">The following script requires two tables named "Table1" and "Table2" to be present.</span></span> <span data-ttu-id="5b509-120">Скрипт сначала проверяет, присутствуют ли таблицы, а затем заканчивается `return` выпиской и соответствующим сообщением, если это не так.</span><span class="sxs-lookup"><span data-stu-id="5b509-120">The script first checks if the tables are present and then ends with the `return` statement and an appropriate message if they're not.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="5b509-121">Если проверка происходит в отдельной функции, вы все равно должны закончить сценарий, выдав `return` выписку из `main` функции.</span><span class="sxs-lookup"><span data-stu-id="5b509-121">If the verification is happening in a separate function, you still must end the script by issuing the `return` statement from the `main` function.</span></span> <span data-ttu-id="5b509-122">Возвращение из подфункции не заканчивается сценарием.</span><span class="sxs-lookup"><span data-stu-id="5b509-122">Returning from the subfunction doesn't end the script.</span></span>

<span data-ttu-id="5b509-123">Следующий скрипт имеет такое же поведение, как и предыдущий.</span><span class="sxs-lookup"><span data-stu-id="5b509-123">The following script has the same behavior as the previous one.</span></span> <span data-ttu-id="5b509-124">Разница в том, что `main` функция вызывает `inputPresent` функцию, чтобы проверить все.</span><span class="sxs-lookup"><span data-stu-id="5b509-124">The difference is that the `main` function calls the `inputPresent` function to verify everything.</span></span> <span data-ttu-id="5b509-125">`inputPresent` возвращает boolean (или `true` ) для того чтобы `false` указать присутствуют ли все необходимые входы.</span><span class="sxs-lookup"><span data-stu-id="5b509-125">`inputPresent` returns a boolean (`true` or `false`) to indicate whether all required inputs are present.</span></span> <span data-ttu-id="5b509-126">Функция `main` использует этот boolean, чтобы принять решение о продолжении или прекращении сценария.</span><span class="sxs-lookup"><span data-stu-id="5b509-126">The `main` function uses that boolean to decide on continuing or ending the script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a><span data-ttu-id="5b509-127">Когда использовать `throw` выписку</span><span class="sxs-lookup"><span data-stu-id="5b509-127">When to use a `throw` statement</span></span>

<span data-ttu-id="5b509-128">В [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) заявлении указывается, что произошла неожиданная ошибка.</span><span class="sxs-lookup"><span data-stu-id="5b509-128">A [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) statement indicates an unexpected error has occurred.</span></span> <span data-ttu-id="5b509-129">Он немедленно завершает код.</span><span class="sxs-lookup"><span data-stu-id="5b509-129">It ends the code immediately.</span></span> <span data-ttu-id="5b509-130">По большей части, вам не нужно из `throw` вашего сценария.</span><span class="sxs-lookup"><span data-stu-id="5b509-130">For the most part, you don't need to `throw` from your script.</span></span> <span data-ttu-id="5b509-131">Как правило, скрипт автоматически информирует пользователя о том, что скрипт не был запущен из-за проблемы.</span><span class="sxs-lookup"><span data-stu-id="5b509-131">Usually, the script automatically informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="5b509-132">В большинстве случаев достаточно закончить сценарий сообщением об ошибке и `return` выпиской из `main` функции.</span><span class="sxs-lookup"><span data-stu-id="5b509-132">In most cases, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="5b509-133">Однако, если скрипт работает как часть Power Automate потока, вы можете остановить поток от продолжения.</span><span class="sxs-lookup"><span data-stu-id="5b509-133">However, if your script is running as part of a Power Automate flow, you may want to stop the flow from continuing.</span></span> <span data-ttu-id="5b509-134">Заявление `throw` останавливает сценарий и говорит поток, чтобы остановить, а также.</span><span class="sxs-lookup"><span data-stu-id="5b509-134">A `throw` statement stops the script and tells the flow to stop as well.</span></span>

<span data-ttu-id="5b509-135">В следующем скрипте показано, как использовать `throw` выписку в примере проверки таблицы.</span><span class="sxs-lookup"><span data-stu-id="5b509-135">The following script shows how to use the `throw` statement in our table checking example.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a><span data-ttu-id="5b509-136">Когда использовать `try...catch` выписку</span><span class="sxs-lookup"><span data-stu-id="5b509-136">When to use a `try...catch` statement</span></span>

<span data-ttu-id="5b509-137">Заявление [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) — это способ определить, не удается ли вызов API, и продолжить запуск скрипта.</span><span class="sxs-lookup"><span data-stu-id="5b509-137">The [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statement is a way to detect if an API call fails and continue running the script.</span></span>

<span data-ttu-id="5b509-138">Рассмотрим следующий фрагмент, который выполняет большое обновление данных на диапазоне.</span><span class="sxs-lookup"><span data-stu-id="5b509-138">Consider the following snippet that performs a large data update on a range.</span></span>

```TypeScript
range.setValues(someLargeValues);
```

<span data-ttu-id="5b509-139">Если `someLargeValues` больше, чем Excel для интернета может обрабатывать, вызов `setValues()` не удается.</span><span class="sxs-lookup"><span data-stu-id="5b509-139">If `someLargeValues` is larger than Excel for the web can handle, the `setValues()` call fails.</span></span> <span data-ttu-id="5b509-140">Скрипт затем также не удается с [ошибкой времени выполнения](../testing/troubleshooting.md#runtime-errors).</span><span class="sxs-lookup"><span data-stu-id="5b509-140">The script then also fails with a [runtime error](../testing/troubleshooting.md#runtime-errors).</span></span> <span data-ttu-id="5b509-141">Заявление `try...catch` позволяет скрипту распознать это условие, без немедленной прекращения сценария и отображения ошибки по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5b509-141">The `try...catch` statement lets your script recognize this condition, without immediately ending the script and showing the default error.</span></span>

<span data-ttu-id="5b509-142">Один из подходов к предоставлению пользователю скрипта лучшего опыта заключается в том, чтобы представить им пользовательское сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="5b509-142">One approach for giving the script user a better experience is to present them a custom error message.</span></span> <span data-ttu-id="5b509-143">Следующий фрагмент показывает заявление `try...catch` регистрации больше информации об ошибках, чтобы лучше помочь читателю.</span><span class="sxs-lookup"><span data-stu-id="5b509-143">The following snippet shows a `try...catch` statement logging more error information to better help the reader.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

<span data-ttu-id="5b509-144">Другой подход к работе с ошибками заключается в том, чтобы иметь обратное поведение, которое обрабатывает случай ошибки.</span><span class="sxs-lookup"><span data-stu-id="5b509-144">Another approach to dealing with errors is to have fallback behavior that handles the error case.</span></span> <span data-ttu-id="5b509-145">Следующий фрагмент использует `catch` блок, чтобы попробовать альтернативный метод разбить обновление на более мелкие части и избежать ошибки.</span><span class="sxs-lookup"><span data-stu-id="5b509-145">The following snippet uses the `catch` block to try an alternate method break up the update into smaller pieces and avoid the error.</span></span>

> [!TIP]
> <span data-ttu-id="5b509-146">Полный пример обновления большого диапазона можно найти в большом [наборе данных.](../resources/samples/write-large-dataset.md)</span><span class="sxs-lookup"><span data-stu-id="5b509-146">For a full example on how to update a large range, see [Write a large dataset](../resources/samples/write-large-dataset.md).</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> <span data-ttu-id="5b509-147">Использование `try...catch` внутри или вокруг цикла замедляет работу скрипта.</span><span class="sxs-lookup"><span data-stu-id="5b509-147">Using `try...catch` inside or around a loop slows down your script.</span></span> <span data-ttu-id="5b509-148">Для получения дополнительной информации о производительности [см. `try...catch` ](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)</span><span class="sxs-lookup"><span data-stu-id="5b509-148">For more performance information, see [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span></span>

## <a name="see-also"></a><span data-ttu-id="5b509-149">См. также</span><span class="sxs-lookup"><span data-stu-id="5b509-149">See also</span></span>

- [<span data-ttu-id="5b509-150">Устранение неполадок в сценариях Office</span><span class="sxs-lookup"><span data-stu-id="5b509-150">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="5b509-151">Информация о устранении неполадок для Power Automate с помощью Office скриптов</span><span class="sxs-lookup"><span data-stu-id="5b509-151">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="5b509-152">Ограничения платформы с Office скриптами</span><span class="sxs-lookup"><span data-stu-id="5b509-152">Platform limits with Office Scripts</span></span>](../testing/platform-limits.md)
- [<span data-ttu-id="5b509-153">Улучшение производительности ваших Office скриптов</span><span class="sxs-lookup"><span data-stu-id="5b509-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
