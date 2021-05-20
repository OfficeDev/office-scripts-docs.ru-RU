---
title: Устранение неполадок Office скриптов
description: Отладка советов и методов для Office скриптов, а также помощь ресурсам.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545560"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="2c592-103">Устранение неполадок Office скриптов</span><span class="sxs-lookup"><span data-stu-id="2c592-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="2c592-104">При разработке Office скриптов вы можете ошибаться.</span><span class="sxs-lookup"><span data-stu-id="2c592-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="2c592-105">Всё в порядке.</span><span class="sxs-lookup"><span data-stu-id="2c592-105">It's okay.</span></span> <span data-ttu-id="2c592-106">У вас есть инструменты, чтобы помочь найти проблемы и получить ваши сценарии работают отлично.</span><span class="sxs-lookup"><span data-stu-id="2c592-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="2c592-107">Типы ошибок</span><span class="sxs-lookup"><span data-stu-id="2c592-107">Types of errors</span></span>

<span data-ttu-id="2c592-108">Office Ошибки скриптов подпадают под одну из двух категорий:</span><span class="sxs-lookup"><span data-stu-id="2c592-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="2c592-109">Ошибки или предупреждения в компиляции времени</span><span class="sxs-lookup"><span data-stu-id="2c592-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="2c592-110">Ошибки времени выполнения</span><span class="sxs-lookup"><span data-stu-id="2c592-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="2c592-111">Ошибки времени компиляции</span><span class="sxs-lookup"><span data-stu-id="2c592-111">Compile-time errors</span></span>

<span data-ttu-id="2c592-112">Ошибки и предупреждения в компиляции изначально отображаются в редакторе Кода.</span><span class="sxs-lookup"><span data-stu-id="2c592-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="2c592-113">Они показаны волнистыми красными подчеркивает в редакторе.</span><span class="sxs-lookup"><span data-stu-id="2c592-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="2c592-114">Они также отображаются под **вкладкой Проблемы** в нижней части панели задач редактора кода.</span><span class="sxs-lookup"><span data-stu-id="2c592-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="2c592-115">Выбор ошибки даст более подробную информацию о проблеме и предложит решения.</span><span class="sxs-lookup"><span data-stu-id="2c592-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="2c592-116">Ошибки времени компиляции должны быть устранены перед запуском скрипта.</span><span class="sxs-lookup"><span data-stu-id="2c592-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Ошибка компилятора, показанная в тексте наведении редактора Кода":::

<span data-ttu-id="2c592-118">Вы также можете увидеть оранжевые предупреждающие подчеркиваемы и серые информационные сообщения.</span><span class="sxs-lookup"><span data-stu-id="2c592-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="2c592-119">Они указывают на производительность предложения или другие возможности, где сценарий может иметь непреднамеренные последствия.</span><span class="sxs-lookup"><span data-stu-id="2c592-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="2c592-120">Такие предупреждения должны быть тщательно изучены, прежде чем отклонить их.</span><span class="sxs-lookup"><span data-stu-id="2c592-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="2c592-121">Ошибки времени выполнения</span><span class="sxs-lookup"><span data-stu-id="2c592-121">Runtime errors</span></span>

<span data-ttu-id="2c592-122">Ошибки выполнения происходят из-за проблем с логикой в скрипте.</span><span class="sxs-lookup"><span data-stu-id="2c592-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="2c592-123">Это может быть связано с тем, что объекта, используемого в скрипте, нет в рабочей книге, таблица отформатирована иначе, чем ожидалось, или какое-либо другое незначительное несоответствие между требованиями скрипта и текущей рабочей книгой.</span><span class="sxs-lookup"><span data-stu-id="2c592-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="2c592-124">Следующий скрипт создает ошибку, когда лист под названием "TestSheet" не присутствует.</span><span class="sxs-lookup"><span data-stu-id="2c592-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="2c592-125">Консольные сообщения</span><span class="sxs-lookup"><span data-stu-id="2c592-125">Console messages</span></span>

<span data-ttu-id="2c592-126">Ошибки компиляции и времени выполнения отображают сообщения об ошибках в консоли при запуске скрипта.</span><span class="sxs-lookup"><span data-stu-id="2c592-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="2c592-127">Они дают номер строки, где проблема была решена.</span><span class="sxs-lookup"><span data-stu-id="2c592-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="2c592-128">Имейте в виду, что основной причиной любой проблемы может быть другая строка кода, чем указано в консоли.</span><span class="sxs-lookup"><span data-stu-id="2c592-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="2c592-129">Следующее изображение показывает выход консоли для явной [ошибки `any` компилятора.](../develop/typescript-restrictions.md)</span><span class="sxs-lookup"><span data-stu-id="2c592-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="2c592-130">Обратите внимание на `[5, 16]` текст в начале строки ошибок.</span><span class="sxs-lookup"><span data-stu-id="2c592-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="2c592-131">Это указывает на ошибку на строке 5, начиная с символа 16.</span><span class="sxs-lookup"><span data-stu-id="2c592-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Консоль редактора кода, отображающая явное сообщение об ошибке «любого»":::

<span data-ttu-id="2c592-133">На последующем изображении показан выход консоли при ошибке времени выполнения.</span><span class="sxs-lookup"><span data-stu-id="2c592-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="2c592-134">Здесь скрипт пытается добавить лист с именем существующего листа.</span><span class="sxs-lookup"><span data-stu-id="2c592-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="2c592-135">Опять же, обратите внимание на "Линия 2", предшествующая ошибке, чтобы показать, какую строку исследовать.</span><span class="sxs-lookup"><span data-stu-id="2c592-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="Консоль редактора кода, отображающая ошибку из вызова 'addWorksheet'":::

## <a name="console-logs"></a><span data-ttu-id="2c592-137">Консольные журналы</span><span class="sxs-lookup"><span data-stu-id="2c592-137">Console logs</span></span>

<span data-ttu-id="2c592-138">Печать сообщений на экран с `console.log` заявлением.</span><span class="sxs-lookup"><span data-stu-id="2c592-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="2c592-139">Эти журналы могут показать текущее значение переменных или какие пути кода срабатывают.</span><span class="sxs-lookup"><span data-stu-id="2c592-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="2c592-140">Для этого звоните с `console.log` любым объектом в качестве параметра.</span><span class="sxs-lookup"><span data-stu-id="2c592-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="2c592-141">Как правило, `string` это самый простой тип для чтения в консоли.</span><span class="sxs-lookup"><span data-stu-id="2c592-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="2c592-142">Строки, `console.log` переданные для отображения в консоли журнала редактора кода, в нижней части панели задач.</span><span class="sxs-lookup"><span data-stu-id="2c592-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="2c592-143">Логи находятся на **вкладке Выход,** хотя вкладка автоматически получает фокус, когда журнал написан.</span><span class="sxs-lookup"><span data-stu-id="2c592-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="2c592-144">Логи не влияют на рабочую книгу.</span><span class="sxs-lookup"><span data-stu-id="2c592-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="2c592-145">Автоматизация вкладки не появляется или Office скрипты недоступны</span><span class="sxs-lookup"><span data-stu-id="2c592-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="2c592-146">Следующие шаги должны помочь устранению неполадок в любых проблемах, связанных **с вкладкой Automate,** не появляющимися Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="2c592-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="2c592-147">[Убедитесь, что ваша лицензия Microsoft 365 включает в себя Office скрипты.](../overview/excel.md#requirements)</span><span class="sxs-lookup"><span data-stu-id="2c592-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="2c592-148">[Убедитесь, что ваш браузер поддерживается.](platform-limits.md#browser-support)</span><span class="sxs-lookup"><span data-stu-id="2c592-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="2c592-149">[Убедитесь, что сторонние файлы cookie включены.](platform-limits.md#third-party-cookies)</span><span class="sxs-lookup"><span data-stu-id="2c592-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="2c592-150">[Убедитесь, что администратор не отключил Office скрипты в Microsoft 365 администратора.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="2c592-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="2c592-151">Скрипты устранения неполадок в Power Automate</span><span class="sxs-lookup"><span data-stu-id="2c592-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="2c592-152">Для получения информации, специфичной для запуска скриптов Power Automate, [см Power Automate Office.](power-automate-troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="2c592-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="2c592-153">Справка по ресурсам</span><span class="sxs-lookup"><span data-stu-id="2c592-153">Help resources</span></span>

<span data-ttu-id="2c592-154">[Stack Overflow –](https://stackoverflow.com/questions/tagged/office-scripts) это сообщество разработчиков, готовых помочь с проблемами кодирования.</span><span class="sxs-lookup"><span data-stu-id="2c592-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="2c592-155">Часто вы сможете найти решение вашей проблемы с помощью быстрого поиска переполнения стеков.</span><span class="sxs-lookup"><span data-stu-id="2c592-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="2c592-156">Если нет, задайте свой вопрос и пометите его тегом "office-scripts".</span><span class="sxs-lookup"><span data-stu-id="2c592-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="2c592-157">Не забудьте упомянуть, что вы создаете *Office,* а не Office *add-in.*</span><span class="sxs-lookup"><span data-stu-id="2c592-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="2c592-158">Если у вас возникла проблема с Office JavaScript API, создайте [проблему в репозитории OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub США.</span><span class="sxs-lookup"><span data-stu-id="2c592-158">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="2c592-159">Члены группы продуктов будут реагировать на вопросы и оказывать дополнительную помощь.</span><span class="sxs-lookup"><span data-stu-id="2c592-159">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="2c592-160">Создание проблемы в **репозитории OfficeDev/office-js** указывает на то, что вы обнаружили дефект в библиотеке API Office JavaScript, который должна устранить группа разработчиков.</span><span class="sxs-lookup"><span data-stu-id="2c592-160">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="2c592-161">Если возникла проблема с регистратором действий или редактором, отправьте отзывы **через кнопку Справка > обратной** связи в Excel.</span><span class="sxs-lookup"><span data-stu-id="2c592-161">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="2c592-162">См. также</span><span class="sxs-lookup"><span data-stu-id="2c592-162">See also</span></span>

- [<span data-ttu-id="2c592-163">Рекомендации по сценариям Office</span><span class="sxs-lookup"><span data-stu-id="2c592-163">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="2c592-164">Ограничения платформы с Office скриптами</span><span class="sxs-lookup"><span data-stu-id="2c592-164">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="2c592-165">Улучшение производительности ваших Office скриптов</span><span class="sxs-lookup"><span data-stu-id="2c592-165">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="2c592-166">Устранение неполадок Office, работающих в PowerAutomate</span><span class="sxs-lookup"><span data-stu-id="2c592-166">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="2c592-167">Отмена эффектов сценариев Office</span><span class="sxs-lookup"><span data-stu-id="2c592-167">Undo the effects of Office Scripts</span></span>](undo.md)
