---
title: Повышение производительности Office скриптов
description: Создайте более быстрые сценарии, понимая связь между Excel книгой и скриптом.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544993"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="72a15-103">Повышение производительности Office скриптов</span><span class="sxs-lookup"><span data-stu-id="72a15-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="72a15-104">Цель Office скриптов — автоматизировать часто выполняемые серии задач, чтобы сэкономить время.</span><span class="sxs-lookup"><span data-stu-id="72a15-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="72a15-105">При медленном сценарии может быть ощущение, что он не ускоряет рабочий процесс.</span><span class="sxs-lookup"><span data-stu-id="72a15-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="72a15-106">В большинстве своем сценарий будет работать в отличном порядке и работать, как и ожидалось.</span><span class="sxs-lookup"><span data-stu-id="72a15-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="72a15-107">Однако существует несколько сценариев, которые могут повлиять на производительность.</span><span class="sxs-lookup"><span data-stu-id="72a15-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="72a15-108">Наиболее частой причиной медленного сценария является чрезмерная связь с книгой.</span><span class="sxs-lookup"><span data-stu-id="72a15-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="72a15-109">Сценарий выполняется на локальном компьютере, а книга существует в облаке.</span><span class="sxs-lookup"><span data-stu-id="72a15-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="72a15-110">В определенное время сценарий синхронизирует локальные данные с данными книги.</span><span class="sxs-lookup"><span data-stu-id="72a15-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="72a15-111">Это означает, что любые операции записи (например) применяются к книге только тогда, когда происходит эта закулисье `workbook.addWorksheet()` синхронизация.</span><span class="sxs-lookup"><span data-stu-id="72a15-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="72a15-112">Кроме того, любые операции чтения (например) получают данные из книги для `myRange.getValues()` скрипта в это время.</span><span class="sxs-lookup"><span data-stu-id="72a15-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="72a15-113">В любом случае сценарий извлекает сведения, прежде чем он будет действовать на данных.</span><span class="sxs-lookup"><span data-stu-id="72a15-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="72a15-114">Например, в следующем коде будет точно входить число строк в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="72a15-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="72a15-115">Office API скриптов гарантируют, что любые данные в книге или скрипте точны и в случае необходимости устарели.</span><span class="sxs-lookup"><span data-stu-id="72a15-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="72a15-116">Вам не нужно беспокоиться об этих синхронизациях для правильного запуска скрипта.</span><span class="sxs-lookup"><span data-stu-id="72a15-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="72a15-117">Однако осведомленность об этом сообщении от скрипта к облаку поможет избежать нежелательных сетевых вызовов.</span><span class="sxs-lookup"><span data-stu-id="72a15-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="72a15-118">Оптимизация производительности</span><span class="sxs-lookup"><span data-stu-id="72a15-118">Performance optimizations</span></span>

<span data-ttu-id="72a15-119">Вы можете применить простые методы, чтобы уменьшить сообщение с облаком.</span><span class="sxs-lookup"><span data-stu-id="72a15-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="72a15-120">Следующие шаблоны помогают ускорить скрипты.</span><span class="sxs-lookup"><span data-stu-id="72a15-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="72a15-121">Чтение данных книг один раз, а не несколько раз в цикле.</span><span class="sxs-lookup"><span data-stu-id="72a15-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="72a15-122">Удаление `console.log` ненужных заявлений.</span><span class="sxs-lookup"><span data-stu-id="72a15-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="72a15-123">Избегайте использования блоков try/catch.</span><span class="sxs-lookup"><span data-stu-id="72a15-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="72a15-124">Чтение данных книг за пределами цикла</span><span class="sxs-lookup"><span data-stu-id="72a15-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="72a15-125">Любой метод, который получает данные из книги, может вызвать сетевой вызов.</span><span class="sxs-lookup"><span data-stu-id="72a15-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="72a15-126">Вместо того, чтобы повторять один и тот же вызов, необходимо сохранять данные локально по мере возможности.</span><span class="sxs-lookup"><span data-stu-id="72a15-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="72a15-127">Это особенно актуально при работе с циклами.</span><span class="sxs-lookup"><span data-stu-id="72a15-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="72a15-128">Рассмотрим сценарий, чтобы получить количество отрицательных чисел в используемом диапазоне таблицы.</span><span class="sxs-lookup"><span data-stu-id="72a15-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="72a15-129">Сценарию необходимо итерировать каждую ячейку используемого диапазона.</span><span class="sxs-lookup"><span data-stu-id="72a15-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="72a15-130">Для этого ему необходимы диапазон, количество строк и количество столбцов.</span><span class="sxs-lookup"><span data-stu-id="72a15-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="72a15-131">Перед запуском цикла следует хранить эти параметры в качестве локальных переменных.</span><span class="sxs-lookup"><span data-stu-id="72a15-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="72a15-132">В противном случае каждая итерация цикла заставит вернуться к книге.</span><span class="sxs-lookup"><span data-stu-id="72a15-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> <span data-ttu-id="72a15-133">В качестве эксперимента попробуйте заменить `usedRangeValues` в цикле `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="72a15-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="72a15-134">При работе с большими диапазонами скрипт может работать значительно дольше.</span><span class="sxs-lookup"><span data-stu-id="72a15-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a><span data-ttu-id="72a15-135">Избегайте `try...catch` использования блоков в или окружающих циклах</span><span class="sxs-lookup"><span data-stu-id="72a15-135">Avoid using `try...catch` blocks in or surrounding loops</span></span>

<span data-ttu-id="72a15-136">Мы не рекомендуем использовать заявления ни в циклах, ни [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) в окружающих циклах.</span><span class="sxs-lookup"><span data-stu-id="72a15-136">We don't recommend using [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statements either in loops or surrounding loops.</span></span> <span data-ttu-id="72a15-137">По той же причине следует избегать чтения данных в цикле: каждая итерация заставляет скрипт синхронизироваться с книгой, чтобы убедиться, что ошибка не была брошена.</span><span class="sxs-lookup"><span data-stu-id="72a15-137">This is for the same reason you should avoid reading data in a loop: each iteration forces the script to synchronize with the workbook to make sure no error has been thrown.</span></span> <span data-ttu-id="72a15-138">Большинство ошибок можно избежать, проверяя объекты, возвращенные из книги.</span><span class="sxs-lookup"><span data-stu-id="72a15-138">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="72a15-139">Например, следующий сценарий проверяет, что таблица, возвращаемая книгой, существует перед попыткой добавить строку.</span><span class="sxs-lookup"><span data-stu-id="72a15-139">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="72a15-140">Удаление `console.log` ненужных заявлений</span><span class="sxs-lookup"><span data-stu-id="72a15-140">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="72a15-141">Ведение журнала консоли — это жизненно важный инструмент для [отладки скриптов.](../testing/troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="72a15-141">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="72a15-142">Однако этот сценарий должен синхронизироваться с книгой, чтобы убедиться в том, что зарегистрированные сведения устарели.</span><span class="sxs-lookup"><span data-stu-id="72a15-142">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="72a15-143">Перед совместным использованием скрипта следует удалить ненужные отчеты о журнале (например, используемые для тестирования).</span><span class="sxs-lookup"><span data-stu-id="72a15-143">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="72a15-144">Обычно это не вызывает заметной проблемы с производительностью, если заявление `console.log()` не находится в цикле.</span><span class="sxs-lookup"><span data-stu-id="72a15-144">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

## <a name="case-by-case-help"></a><span data-ttu-id="72a15-145">Помощь в разных случаях</span><span class="sxs-lookup"><span data-stu-id="72a15-145">Case-by-case help</span></span>

<span data-ttu-id="72a15-146">По мере расширения платформы Office скриптов для работы с [Power Automate,](https://flow.microsoft.com/) [адаптивными](/adaptive-cards)картами и другими функциями кросс-продукта, сведения о связи скрипта и книги становятся более сложными.</span><span class="sxs-lookup"><span data-stu-id="72a15-146">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="72a15-147">Если вам нужна помощь по ускорению запуска сценария, обратитесь к [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span><span class="sxs-lookup"><span data-stu-id="72a15-147">If you need help making your script run faster, please reach out through [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span></span> <span data-ttu-id="72a15-148">Обязательно пометите свой вопрос с помощью "office-scripts-dev", чтобы эксперты могли найти его и помочь.</span><span class="sxs-lookup"><span data-stu-id="72a15-148">Be sure to tag your question with "office-scripts-dev" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="72a15-149">См. также</span><span class="sxs-lookup"><span data-stu-id="72a15-149">See also</span></span>

- [<span data-ttu-id="72a15-150">Основные сведения о сценариях Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="72a15-150">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="72a15-151">Веб-документы MDN: циклы и итерация</span><span class="sxs-lookup"><span data-stu-id="72a15-151">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
