---
title: Повышение производительности сценариев Office
description: Создавайте более быстрые сценарии, понимая связь между книгой Excel и сценарием.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: ce50a6fd7ad02ddcd2dd304be8b4dd8fa3d0acf3
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867872"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="c2d23-103">Повышение производительности сценариев Office</span><span class="sxs-lookup"><span data-stu-id="c2d23-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="c2d23-104">Целью сценариев Office является автоматизация часто выполняемых рядов задач, чтобы сэкономить время.</span><span class="sxs-lookup"><span data-stu-id="c2d23-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="c2d23-105">Медленный сценарий может выглядеть так, будто он не ускоряет рабочий процесс.</span><span class="sxs-lookup"><span data-stu-id="c2d23-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="c2d23-106">В большинстве своем сценарий будет работать безошибок.</span><span class="sxs-lookup"><span data-stu-id="c2d23-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="c2d23-107">Однако существует несколько сценариев, которые можно избежать, которые могут повлиять на производительность.</span><span class="sxs-lookup"><span data-stu-id="c2d23-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="c2d23-108">Наиболее распространенная причина медленного сценария — чрезмерное взаимодействие с книгой.</span><span class="sxs-lookup"><span data-stu-id="c2d23-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="c2d23-109">Сценарий выполняется на локальном компьютере, а книга существует в облаке.</span><span class="sxs-lookup"><span data-stu-id="c2d23-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="c2d23-110">В определенное время сценарий синхронизирует локальные данные с данными книги.</span><span class="sxs-lookup"><span data-stu-id="c2d23-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="c2d23-111">Это означает, что любые операции записи (например,) применяются к книге только при такой синхронизации за `workbook.addWorksheet()` кадром.</span><span class="sxs-lookup"><span data-stu-id="c2d23-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="c2d23-112">Аналогично, в такие моменты любые операции чтения (например,) получают данные только из книги для `myRange.getValues()` сценария.</span><span class="sxs-lookup"><span data-stu-id="c2d23-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="c2d23-113">В любом случае сценарий получает сведения, прежде чем он будет действовать с данными.</span><span class="sxs-lookup"><span data-stu-id="c2d23-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="c2d23-114">Например, в следующем коде точно занося в журнал количество строк в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="c2d23-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="c2d23-115">API сценариев Office обеспечивают точность и правильность любых данных в книге или сценарии при необходимости.</span><span class="sxs-lookup"><span data-stu-id="c2d23-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="c2d23-116">Вам не нужно беспокоиться об этих синхронизациях для правильного запуска скрипта.</span><span class="sxs-lookup"><span data-stu-id="c2d23-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="c2d23-117">Тем не менее, понимание этого сценария для облачного взаимодействия поможет избежать нежелательных сетевых вызовов.</span><span class="sxs-lookup"><span data-stu-id="c2d23-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="c2d23-118">Оптимизация производительности</span><span class="sxs-lookup"><span data-stu-id="c2d23-118">Performance optimizations</span></span>

<span data-ttu-id="c2d23-119">Вы можете применять простые методы, чтобы сократить объем взаимодействия с облаком.</span><span class="sxs-lookup"><span data-stu-id="c2d23-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="c2d23-120">Следующие шаблоны помогают ускорить ваши сценарии.</span><span class="sxs-lookup"><span data-stu-id="c2d23-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="c2d23-121">Чтение данных книги один раз, а не несколько раз в цикле.</span><span class="sxs-lookup"><span data-stu-id="c2d23-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="c2d23-122">Удалите `console.log` ненужные утверждения.</span><span class="sxs-lookup"><span data-stu-id="c2d23-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="c2d23-123">Избегайте использования блоков try/catch.</span><span class="sxs-lookup"><span data-stu-id="c2d23-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="c2d23-124">Чтение данных книги вне цикла</span><span class="sxs-lookup"><span data-stu-id="c2d23-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="c2d23-125">Любой метод, который получает данные из книги, может вызвать сетевой вызов.</span><span class="sxs-lookup"><span data-stu-id="c2d23-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="c2d23-126">Вместо того чтобы повторять один и тот же вызов, по возможности следует сохранять данные локально.</span><span class="sxs-lookup"><span data-stu-id="c2d23-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="c2d23-127">Это особенно актуально при работе с циклами.</span><span class="sxs-lookup"><span data-stu-id="c2d23-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="c2d23-128">Рассмотрим сценарий, чтобы получить количество отрицательных чисел в используемом диапазоне таблицы.</span><span class="sxs-lookup"><span data-stu-id="c2d23-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="c2d23-129">Сценарию необходимо итерировать каждую ячейку в используемом диапазоне.</span><span class="sxs-lookup"><span data-stu-id="c2d23-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="c2d23-130">Для этого ему требуется диапазон, количество строк и число столбцов.</span><span class="sxs-lookup"><span data-stu-id="c2d23-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="c2d23-131">Перед запуском цикла их следует сохранить в качестве локальных переменных.</span><span class="sxs-lookup"><span data-stu-id="c2d23-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="c2d23-132">В противном случае каждая итерация цикла принудительно возвращает книгу.</span><span class="sxs-lookup"><span data-stu-id="c2d23-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="c2d23-133">В качестве эксперимента попробуйте заменить `usedRangeValues` в цикле на `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="c2d23-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="c2d23-134">При работе с большими диапазонами может потребоваться значительно больше времени.</span><span class="sxs-lookup"><span data-stu-id="c2d23-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="c2d23-135">Удаление ненужных `console.log` заявлений</span><span class="sxs-lookup"><span data-stu-id="c2d23-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="c2d23-136">Ведение журнала консоли — это важный инструмент для отладки [сценариев.](../testing/troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="c2d23-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="c2d23-137">Однако он принудительно синхронизирует сценарий с книгой, чтобы убедиться, что зарегистрированные сведения имеют последние данные.</span><span class="sxs-lookup"><span data-stu-id="c2d23-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="c2d23-138">Перед совместным использованием скрипта можно удалить ненужные утверждения ведения журнала (например, используемые для тестирования).</span><span class="sxs-lookup"><span data-stu-id="c2d23-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="c2d23-139">Как правило, это не вызывает заметной проблемы с производительностью, если только данный отчет не `console.log()` находится в цикле.</span><span class="sxs-lookup"><span data-stu-id="c2d23-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="c2d23-140">Избегайте использования блоков try/catch</span><span class="sxs-lookup"><span data-stu-id="c2d23-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="c2d23-141">Мы не рекомендуем использовать [ `try` / `catch` блоки](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) в рамках ожидаемого потока управления сценария.</span><span class="sxs-lookup"><span data-stu-id="c2d23-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="c2d23-142">Большинство ошибок можно избежать, проверяя объекты, возвращенные из книги.</span><span class="sxs-lookup"><span data-stu-id="c2d23-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="c2d23-143">Например, следующий сценарий проверяет, существует ли таблица, возвращенная книгой, перед попыткой добавления строки.</span><span class="sxs-lookup"><span data-stu-id="c2d23-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

## <a name="case-by-case-help"></a><span data-ttu-id="c2d23-144">Справка по делу</span><span class="sxs-lookup"><span data-stu-id="c2d23-144">Case-by-case help</span></span>

<span data-ttu-id="c2d23-145">По мере расширения платформы сценариев Office для работы с [Power Automate,](https://flow.microsoft.com/) [адаптивными](/adaptive-cards)карточками и другими функциями для разных продуктов подробности взаимодействия между скриптами и книгой становятся более сложными.</span><span class="sxs-lookup"><span data-stu-id="c2d23-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="c2d23-146">Если вам нужна помощь в ускорении запуска скрипта, свяжитесь с [помощью Stack Overflow.](https://stackoverflow.com/questions/tagged/office-scripts)</span><span class="sxs-lookup"><span data-stu-id="c2d23-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="c2d23-147">Не забудьте пометить свой вопрос с помощью "office-scripts", чтобы эксперты могли найти его и помочь.</span><span class="sxs-lookup"><span data-stu-id="c2d23-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="c2d23-148">См. также</span><span class="sxs-lookup"><span data-stu-id="c2d23-148">See also</span></span>

- [<span data-ttu-id="c2d23-149">Основные сведения о сценариях Office в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="c2d23-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="c2d23-150">Веб-документы MDN: циклы и итерация</span><span class="sxs-lookup"><span data-stu-id="c2d23-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)