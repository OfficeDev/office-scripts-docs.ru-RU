---
title: Устранение Office скриптов
description: Отладка советов и методов для Office скриптов, а также ресурсов справки.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 251ad72588422a86c52c81666164c2c4bd79bdb5
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074650"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="efbde-103">Устранение Office скриптов</span><span class="sxs-lookup"><span data-stu-id="efbde-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="efbde-104">При разработке Office скриптов вы можете ошибаться.</span><span class="sxs-lookup"><span data-stu-id="efbde-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="efbde-105">Всё в порядке.</span><span class="sxs-lookup"><span data-stu-id="efbde-105">It's okay.</span></span> <span data-ttu-id="efbde-106">У вас есть средства, которые помогут найти проблемы и получить ваши сценарии работают идеально.</span><span class="sxs-lookup"><span data-stu-id="efbde-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="efbde-107">Типы ошибок</span><span class="sxs-lookup"><span data-stu-id="efbde-107">Types of errors</span></span>

<span data-ttu-id="efbde-108">Office Ошибки скриптов подпадают под одну из двух категорий:</span><span class="sxs-lookup"><span data-stu-id="efbde-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="efbde-109">Ошибки и предупреждения по времени компиляции</span><span class="sxs-lookup"><span data-stu-id="efbde-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="efbde-110">Ошибки во время работы</span><span class="sxs-lookup"><span data-stu-id="efbde-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="efbde-111">Ошибки со временем компиляции</span><span class="sxs-lookup"><span data-stu-id="efbde-111">Compile-time errors</span></span>

<span data-ttu-id="efbde-112">Ошибки и предупреждения по времени компиляции сначала показаны в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="efbde-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="efbde-113">Они показаны волнистыми красными линиями в редакторе.</span><span class="sxs-lookup"><span data-stu-id="efbde-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="efbde-114">Они также отображаются в вкладке **Проблемы** в нижней части области задач редактора кода.</span><span class="sxs-lookup"><span data-stu-id="efbde-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="efbde-115">Выбор ошибки даст дополнительные сведения о проблеме и предложит решения.</span><span class="sxs-lookup"><span data-stu-id="efbde-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="efbde-116">Ошибки со временем компиляции следует устранить перед запуском сценария.</span><span class="sxs-lookup"><span data-stu-id="efbde-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Ошибка компиляторов, показанная в тексте наведении редактора кода.":::

<span data-ttu-id="efbde-118">Кроме того, можно увидеть оранжевые предупреждения и серые информационные сообщения.</span><span class="sxs-lookup"><span data-stu-id="efbde-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="efbde-119">Они указывают на предложения по производительности или другие возможности, при которых сценарий может иметь непреднамеренные последствия.</span><span class="sxs-lookup"><span data-stu-id="efbde-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="efbde-120">Такие предупреждения следует внимательно исследовать перед их отклонением.</span><span class="sxs-lookup"><span data-stu-id="efbde-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="efbde-121">Ошибки во время работы</span><span class="sxs-lookup"><span data-stu-id="efbde-121">Runtime errors</span></span>

<span data-ttu-id="efbde-122">Ошибки во время работы происходят из-за проблем с логикой в скрипте.</span><span class="sxs-lookup"><span data-stu-id="efbde-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="efbde-123">Это может быть из-за того, что объект, используемый в сценарии, не находится в книге, таблица отформатирована иначе, чем ожидалось, или некоторыми другими незначительными несоответствиями между требованиями скрипта и текущей книгой.</span><span class="sxs-lookup"><span data-stu-id="efbde-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="efbde-124">В следующем сценарии создается ошибка, когда нет таблицы с именем "TestSheet".</span><span class="sxs-lookup"><span data-stu-id="efbde-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="efbde-125">Сообщения консоли</span><span class="sxs-lookup"><span data-stu-id="efbde-125">Console messages</span></span>

<span data-ttu-id="efbde-126">Ошибки времени компилирования и времени работы отображают сообщения об ошибках в консоли при запуске скрипта.</span><span class="sxs-lookup"><span data-stu-id="efbde-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="efbde-127">Они дают номер строки, где возникла проблема.</span><span class="sxs-lookup"><span data-stu-id="efbde-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="efbde-128">Имейте в виду, что основной причиной любой проблемы может быть другая строка кода, чем указано на консоли.</span><span class="sxs-lookup"><span data-stu-id="efbde-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="efbde-129">На следующем изображении показан выход консоли для [явной ошибки `any` ](../develop/typescript-restrictions.md) компиляторов.</span><span class="sxs-lookup"><span data-stu-id="efbde-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="efbde-130">Обратите внимание на `[5, 16]` текст в начале строки ошибки.</span><span class="sxs-lookup"><span data-stu-id="efbde-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="efbde-131">Это означает, что ошибка находится на строке 5, начиная с символа 16.</span><span class="sxs-lookup"><span data-stu-id="efbde-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Консоль редактора кода, отображаемая явное сообщение об ошибке &quot;любое&quot;.":::

<span data-ttu-id="efbde-133">На последующем изображении показан выход консоли для ошибки во время запуска.</span><span class="sxs-lookup"><span data-stu-id="efbde-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="efbde-134">Здесь сценарий пытается добавить таблицу с именем существующего таблицы.</span><span class="sxs-lookup"><span data-stu-id="efbde-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="efbde-135">Снова обратите внимание на "Строку 2", предшествуя ошибке, чтобы показать, какую строку следует исследовать.</span><span class="sxs-lookup"><span data-stu-id="efbde-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="Консоль редактора кода с ошибкой из вызова &quot;addWorksheet&quot;.":::

## <a name="console-logs"></a><span data-ttu-id="efbde-137">Журналы консоли</span><span class="sxs-lookup"><span data-stu-id="efbde-137">Console logs</span></span>

<span data-ttu-id="efbde-138">Печать сообщений на экран с помощью `console.log` заявления.</span><span class="sxs-lookup"><span data-stu-id="efbde-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="efbde-139">Эти журналы могут показать текущее значение переменных или запускать пути кода.</span><span class="sxs-lookup"><span data-stu-id="efbde-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="efbde-140">Для этого необходимо вызвать `console.log` любой объект в качестве параметра.</span><span class="sxs-lookup"><span data-stu-id="efbde-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="efbde-141">Обычно самый `string` простой тип для чтения на консоли.</span><span class="sxs-lookup"><span data-stu-id="efbde-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="efbde-142">Строки, переданные для отображения в консоли журнала редактора кода, в нижней `console.log` части области задач.</span><span class="sxs-lookup"><span data-stu-id="efbde-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="efbde-143">Журналы находятся на **вкладке Выход,** хотя вкладка автоматически получает фокус при записи журнала.</span><span class="sxs-lookup"><span data-stu-id="efbde-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="efbde-144">Журналы не влияют на книгу.</span><span class="sxs-lookup"><span data-stu-id="efbde-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="efbde-145">Автоматизация вкладки, не появляющихся или Office недоступных скриптов</span><span class="sxs-lookup"><span data-stu-id="efbde-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="efbde-146">Следующие действия должны помочь устранить проблемы, связанные с вкладками **Automate,** которые не отображаются в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="efbde-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="efbde-147">[Убедитесь, что Microsoft 365 лицензия включает Office скрипты.](../overview/excel.md#requirements)</span><span class="sxs-lookup"><span data-stu-id="efbde-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="efbde-148">[Убедитесь, что браузер поддерживается.](platform-limits.md#browser-support)</span><span class="sxs-lookup"><span data-stu-id="efbde-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="efbde-149">[Убедитесь, что сторонние файлы cookie включены.](platform-limits.md#third-party-cookies)</span><span class="sxs-lookup"><span data-stu-id="efbde-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="efbde-150">[Убедитесь, что администратор не отключил Office скрипты в Центр администрирования Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="efbde-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="efbde-151">Сценарии устранения неполадок в Power Automate</span><span class="sxs-lookup"><span data-stu-id="efbde-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="efbde-152">Сведения, специфические для запуска сценариев Power Automate, см. в Office сценариев, [запущенных в Power Automate.](power-automate-troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="efbde-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="efbde-153">Ресурсы справки</span><span class="sxs-lookup"><span data-stu-id="efbde-153">Help resources</span></span>

<span data-ttu-id="efbde-154">[Переполнение стека](https://stackoverflow.com/questions/tagged/office-scripts) — это сообщество разработчиков, готовых помочь с проблемами кодирования.</span><span class="sxs-lookup"><span data-stu-id="efbde-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="efbde-155">Часто вы сможете найти решение проблемы с помощью быстрого поиска переполнения стека.</span><span class="sxs-lookup"><span data-stu-id="efbde-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="efbde-156">Если нет, задайте свой вопрос и пометите его тегом "Office-scripts".</span><span class="sxs-lookup"><span data-stu-id="efbde-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="efbde-157">Не забудьте упомянуть, что вы создаете сценарий *Office,* а не Office *надстройки.*</span><span class="sxs-lookup"><span data-stu-id="efbde-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="efbde-158">Чтобы отправить запрос на Office скрипты, отправьте свою идею на страницу Голос пользователя [или](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439)если запрос на функции уже существует, добавьте свой голос за него.</span><span class="sxs-lookup"><span data-stu-id="efbde-158">To submit a feature request for Office Scripts, post your idea to our [User Voice page](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439), or if the feature request already exists there, add your vote for it.</span></span> <span data-ttu-id="efbde-159">Не забудьте подать запрос в Excel для Интернета в категории "Макрос, скрипты и надстройки".</span><span class="sxs-lookup"><span data-stu-id="efbde-159">Be sure to file the request under Excel for the web in the "Macros, Scripts and Add-ins" category.</span></span>

<span data-ttu-id="efbde-160">Если возникла проблема с регистратором действий или редактором, сообщите нам об этом.</span><span class="sxs-lookup"><span data-stu-id="efbde-160">If there is a problem with the Action Recorder or Editor, please let us know.</span></span> <span data-ttu-id="efbde-161">В меню области задач редактора кода **выберите** кнопку **Отправка** отзывов, чтобы поделиться любыми вопросами.</span><span class="sxs-lookup"><span data-stu-id="efbde-161">In the Code Editor task pane's **...** menu, select the **Send feedback** button to share any issues.</span></span>

:::image type="content" source="../images/code-editor-feedback.png" alt-text="Меню редактора кода переполнено кнопкой &quot;Отправить обратную связь&quot;.":::

## <a name="see-also"></a><span data-ttu-id="efbde-163">См. также</span><span class="sxs-lookup"><span data-stu-id="efbde-163">See also</span></span>

- [<span data-ttu-id="efbde-164">Рекомендации по сценариям Office</span><span class="sxs-lookup"><span data-stu-id="efbde-164">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="efbde-165">Ограничения платформы с Office скриптами</span><span class="sxs-lookup"><span data-stu-id="efbde-165">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="efbde-166">Повышение производительности Office скриптов</span><span class="sxs-lookup"><span data-stu-id="efbde-166">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="efbde-167">Устранение Office сценариев, запущенных в PowerAutomate</span><span class="sxs-lookup"><span data-stu-id="efbde-167">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="efbde-168">Отмена эффектов сценариев Office</span><span class="sxs-lookup"><span data-stu-id="efbde-168">Undo the effects of Office Scripts</span></span>](undo.md)
