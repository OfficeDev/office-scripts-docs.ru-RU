---
title: Улучшение производительности ваших Office скриптов
description: Создавайте более быстрые сценарии, понимая связь между Excel и вашим скриптом.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544993"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="77177-103">Улучшение производительности ваших Office скриптов</span><span class="sxs-lookup"><span data-stu-id="77177-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="77177-104">Целью этих Office является автоматизация обычно выполняемых ряд задач, чтобы сэкономить время.</span><span class="sxs-lookup"><span data-stu-id="77177-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="77177-105">Медленный сценарий может чувствовать, что он не ускоряет ваш рабочий процесс.</span><span class="sxs-lookup"><span data-stu-id="77177-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="77177-106">Большую часть времени, ваш сценарий будет прекрасно и работать, как ожидалось.</span><span class="sxs-lookup"><span data-stu-id="77177-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="77177-107">Тем не менее, есть несколько, предотвратимых сценариев, которые могут повлиять на производительность.</span><span class="sxs-lookup"><span data-stu-id="77177-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="77177-108">Наиболее распространенной причиной медленного сценария является чрезмерное общение с рабочей книгой.</span><span class="sxs-lookup"><span data-stu-id="77177-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="77177-109">Скрипт работает на локальной машине, в то время как рабочая книга существует в облаке.</span><span class="sxs-lookup"><span data-stu-id="77177-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="77177-110">В определенное время скрипт синхронизирует свои локальные данные с данными рабочей книги.</span><span class="sxs-lookup"><span data-stu-id="77177-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="77177-111">Это означает, что любые операции записи `workbook.addWorksheet()` (например), применяются к рабочей книге только тогда, когда происходит эта закулисная синхронизация.</span><span class="sxs-lookup"><span data-stu-id="77177-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="77177-112">Аналогичным образом, любые операции чтения `myRange.getValues()` (например), получают данные из рабочей книги только для сценария в то время.</span><span class="sxs-lookup"><span data-stu-id="77177-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="77177-113">В любом случае скрипт получает информацию, прежде чем он действует на данные.</span><span class="sxs-lookup"><span data-stu-id="77177-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="77177-114">Например, следующий код точно залогит количество строк в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="77177-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="77177-115">Office API-файлы скриптов гарантируют, что любые данные в рабочей книге или скрипте являются точными и точными, когда это необходимо.</span><span class="sxs-lookup"><span data-stu-id="77177-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="77177-116">Вам не нужно беспокоиться об этих синхронизациях для правильного запуска скрипта.</span><span class="sxs-lookup"><span data-stu-id="77177-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="77177-117">Тем не менее, осведомленность об этой связи между скриптами и облаками может помочь вам избежать ненужных сетевых звонков.</span><span class="sxs-lookup"><span data-stu-id="77177-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="77177-118">Оптимизация производительности</span><span class="sxs-lookup"><span data-stu-id="77177-118">Performance optimizations</span></span>

<span data-ttu-id="77177-119">Вы можете применить простые методы, чтобы помочь уменьшить связь в облаке.</span><span class="sxs-lookup"><span data-stu-id="77177-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="77177-120">Следующие шаблоны помогают ускорить ваши скрипты.</span><span class="sxs-lookup"><span data-stu-id="77177-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="77177-121">Прочитайте данные рабочей книги один раз, а не повторно в цикле.</span><span class="sxs-lookup"><span data-stu-id="77177-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="77177-122">Удалите ненужные `console.log` операторы.</span><span class="sxs-lookup"><span data-stu-id="77177-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="77177-123">Избегайте использования попробовать / поймать блоков.</span><span class="sxs-lookup"><span data-stu-id="77177-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="77177-124">Читать данные о работе вне цикла</span><span class="sxs-lookup"><span data-stu-id="77177-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="77177-125">Любой метод, который получает данные из рабочей книги, может вызвать сетевой звонок.</span><span class="sxs-lookup"><span data-stu-id="77177-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="77177-126">Вместо того, чтобы неоднократно делать один и тот же вызов, вы должны сохранить данные локально, когда это возможно.</span><span class="sxs-lookup"><span data-stu-id="77177-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="77177-127">Это особенно верно при работе с петлями.</span><span class="sxs-lookup"><span data-stu-id="77177-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="77177-128">Рассмотрим сценарий, чтобы получить подсчет отрицательных чисел в используемом диапазоне листа.</span><span class="sxs-lookup"><span data-stu-id="77177-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="77177-129">Скрипт должен итерировать над каждой ячейкой в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="77177-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="77177-130">Для этого ему нужен диапазон, количество строк и количество столбцов.</span><span class="sxs-lookup"><span data-stu-id="77177-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="77177-131">Вы должны хранить их в качестве локальных переменных перед началом цикла.</span><span class="sxs-lookup"><span data-stu-id="77177-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="77177-132">В противном случае каждая итерация цикла заставит вернуться к рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="77177-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="77177-133">В качестве эксперимента попробуйте `usedRangeValues` заменить в цикле `usedRange.getValues()` с .</span><span class="sxs-lookup"><span data-stu-id="77177-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="77177-134">Вы можете заметить, что запуск скрипта занимает значительно больше времени при работе с большими диапазонами.</span><span class="sxs-lookup"><span data-stu-id="77177-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a><span data-ttu-id="77177-135">Избегайте использования `try...catch` блоков в или окружающих петель</span><span class="sxs-lookup"><span data-stu-id="77177-135">Avoid using `try...catch` blocks in or surrounding loops</span></span>

<span data-ttu-id="77177-136">Мы не рекомендуем использовать операторы [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) в петлях или окружающих петлях.</span><span class="sxs-lookup"><span data-stu-id="77177-136">We don't recommend using [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statements either in loops or surrounding loops.</span></span> <span data-ttu-id="77177-137">Это по той же причине, по которой следует избегать чтения данных в цикле: каждая итерация заставляет скрипт синхронизироваться с рабочей книгой, чтобы убедиться, что ошибка не была брошена.</span><span class="sxs-lookup"><span data-stu-id="77177-137">This is for the same reason you should avoid reading data in a loop: each iteration forces the script to synchronize with the workbook to make sure no error has been thrown.</span></span> <span data-ttu-id="77177-138">Большинство ошибок можно избежать, проверяя объекты, возвращенные из рабочей книги.</span><span class="sxs-lookup"><span data-stu-id="77177-138">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="77177-139">Например, следующий скрипт проверяет, что таблица, возвращенная рабочей книгой, существует, прежде чем пытаться добавить строку.</span><span class="sxs-lookup"><span data-stu-id="77177-139">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="77177-140">Удалить ненужные `console.log` операторы</span><span class="sxs-lookup"><span data-stu-id="77177-140">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="77177-141">Консольный журнал является жизненно важным инструментом [для отладки скриптов.](../testing/troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="77177-141">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="77177-142">Тем не менее, это требует синхронизации скрипта с рабочей книгой, чтобы убедиться, что зарегистрированная информация является в курсе.</span><span class="sxs-lookup"><span data-stu-id="77177-142">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="77177-143">Перед просмотром сценария следует удалить ненужные операторы журналов (например, те, которые используются для тестирования).</span><span class="sxs-lookup"><span data-stu-id="77177-143">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="77177-144">Обычно это не вызывает заметной проблемы с производительностью, если только `console.log()` заявление не находится в цикле.</span><span class="sxs-lookup"><span data-stu-id="77177-144">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

## <a name="case-by-case-help"></a><span data-ttu-id="77177-145">Помощь в каждом конкретном случае</span><span class="sxs-lookup"><span data-stu-id="77177-145">Case-by-case help</span></span>

<span data-ttu-id="77177-146">По мере Office платформы Power Automate, [адаптивных карт](/adaptive-cards) [и](https://flow.microsoft.com/)других кросс-продуктов, детали общения скрипта и рабочей книги становятся все более запутанными.</span><span class="sxs-lookup"><span data-stu-id="77177-146">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="77177-147">Если вам нужна помощь в том, чтобы сделать ваш скрипт работать быстрее, пожалуйста, пройдите [через корпорацию Майкрософт&A.](/answers/topics/office-scripts-dev.html)</span><span class="sxs-lookup"><span data-stu-id="77177-147">If you need help making your script run faster, please reach out through [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span></span> <span data-ttu-id="77177-148">Не забудьте отметить ваш вопрос с "офис-скрипты-dev", чтобы эксперты могли найти его и помочь.</span><span class="sxs-lookup"><span data-stu-id="77177-148">Be sure to tag your question with "office-scripts-dev" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="77177-149">См. также</span><span class="sxs-lookup"><span data-stu-id="77177-149">See also</span></span>

- [<span data-ttu-id="77177-150">Основные сведения о сценариях Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="77177-150">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="77177-151">Веб-документы MDN: петли и итерация</span><span class="sxs-lookup"><span data-stu-id="77177-151">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
