---
title: Вы запустите Office скрипты с Power Automate
description: Как получить Office для Excel в Интернете с рабочим Power Automate процесса.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7562a2b2359cde67a9a47e0640515018fe23ac35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545042"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="f0d04-103">Вы запустите Office скрипты с Power Automate</span><span class="sxs-lookup"><span data-stu-id="f0d04-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="f0d04-104">[Power Automate](https://flow.microsoft.com) позволяет добавлять Office скрипты в более крупный автоматизированный рабочий процесс.</span><span class="sxs-lookup"><span data-stu-id="f0d04-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="f0d04-105">Вы можете использовать Power Automate делать такие вещи, как добавление содержимого электронной почты в таблицу листа или создавать действия в инструментах управления проектами на основе комментариев к рабочей книге.</span><span class="sxs-lookup"><span data-stu-id="f0d04-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="get-started"></a><span data-ttu-id="f0d04-106">Начало работы</span><span class="sxs-lookup"><span data-stu-id="f0d04-106">Get started</span></span>

<span data-ttu-id="f0d04-107">Если вы только что Power Automate, мы рекомендуем [посетить Начало работы с Power Automate](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="f0d04-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="f0d04-108">Там вы можете узнать больше обо всех доступных вам возможностях автоматизации.</span><span class="sxs-lookup"><span data-stu-id="f0d04-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="f0d04-109">Документы здесь сосредоточены на том, Office скрипты работают Power Automate и как это может помочь улучшить Excel опыт.</span><span class="sxs-lookup"><span data-stu-id="f0d04-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="f0d04-110">Чтобы начать Power Automate и Office, следуйте [учебнику Начните использовать скрипты с Power Automate.](../tutorials/excel-power-automate-manual.md)</span><span class="sxs-lookup"><span data-stu-id="f0d04-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="f0d04-111">Это научит вас, как создать поток, который вызывает простой сценарий.</span><span class="sxs-lookup"><span data-stu-id="f0d04-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="f0d04-112">После того как вы завершили этот учебник и [данные Pass для скриптов в автоматическом учебнике потока Power Automate,](../tutorials/excel-power-automate-trigger.md) вернитесь сюда для получения подробной информации о подключении Office scripts к Power Automate потокам.</span><span class="sxs-lookup"><span data-stu-id="f0d04-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="f0d04-113">Excel Онлайн (Бизнес) разъем</span><span class="sxs-lookup"><span data-stu-id="f0d04-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="f0d04-114">[Коннекторы](/connectors/connectors) являются мостами между Power Automate и приложениями.</span><span class="sxs-lookup"><span data-stu-id="f0d04-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="f0d04-115">Разъем [Excel Online (Business) предоставляет](/connectors/excelonlinebusiness) вашим потокам доступ к Excel книгам.</span><span class="sxs-lookup"><span data-stu-id="f0d04-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="f0d04-116">Действие "Run script" позволяет вызвать любую Office, доступную через выбранную трудовую книжку.</span><span class="sxs-lookup"><span data-stu-id="f0d04-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="f0d04-117">Вы также можете предоставить параметры ввода скриптов, чтобы данные могли быть предоставлены потоком, или иметь информацию о возврате скрипта для более поздних шагов в потоке.</span><span class="sxs-lookup"><span data-stu-id="f0d04-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f0d04-118">Действие "Run script" дает людям, которые используют Excel разъем значительный доступ к вашей рабочей книге и ее данным.</span><span class="sxs-lookup"><span data-stu-id="f0d04-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="f0d04-119">Кроме того, существуют риски безопасности со скриптами, которые делают внешние вызовы API, как это [объясняется во внешних звонках Power Automate.](external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="f0d04-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="f0d04-120">Если ваш администратор обеспокоен воздействием высокочувствительных данных, он может либо отключить разъем Excel Online, либо ограничить доступ к скриптам Office [через Office Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="f0d04-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="f0d04-121">Передача данных в потоках для скриптов</span><span class="sxs-lookup"><span data-stu-id="f0d04-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="f0d04-122">Power Automate позволяет передавать фрагменты данных между шагами вашего потока.</span><span class="sxs-lookup"><span data-stu-id="f0d04-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="f0d04-123">Скрипты могут быть настроены, чтобы принимать любые типы информации, вам нужно, и возвращать что-либо из вашей рабочей книги, что вы хотите в потоке.</span><span class="sxs-lookup"><span data-stu-id="f0d04-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="f0d04-124">Ввод скрипта определяется путем добавления параметров к `main` функции (в дополнение `workbook: ExcelScript.Workbook` к).</span><span class="sxs-lookup"><span data-stu-id="f0d04-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="f0d04-125">Выход из скрипта объявляется путем добавления типа возврата `main` к .</span><span class="sxs-lookup"><span data-stu-id="f0d04-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="f0d04-126">При создании блока "Run Script" в потоке заполняется принятые параметры и возвращенные типы.</span><span class="sxs-lookup"><span data-stu-id="f0d04-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="f0d04-127">Если вы измените параметры или вернете типы скрипта, вам нужно будет переписать блок потока "Run script".</span><span class="sxs-lookup"><span data-stu-id="f0d04-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="f0d04-128">Это гарантирует, что данные правильно разобраются.</span><span class="sxs-lookup"><span data-stu-id="f0d04-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="f0d04-129">Следующие разделы охватывают детали ввода и вывода скриптов, используемых в Power Automate.</span><span class="sxs-lookup"><span data-stu-id="f0d04-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="f0d04-130">Если вы хотите практический подход к изучению этой темы, попробуйте [данные Pass для скриптов в автоматическом учебнике по потоку Power Automate](../tutorials/excel-power-automate-trigger.md) или [исследуйте пример автоматического напоминания](../resources/scenarios/task-reminders.md) о задачах.</span><span class="sxs-lookup"><span data-stu-id="f0d04-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-pass-data-to-a-script"></a><span data-ttu-id="f0d04-131">`main` Параметры: Переведать данные в скрипт</span><span class="sxs-lookup"><span data-stu-id="f0d04-131">`main` Parameters: Pass data to a script</span></span>

<span data-ttu-id="f0d04-132">Все ввода скрипта указаны в качестве дополнительных параметров для `main` функции.</span><span class="sxs-lookup"><span data-stu-id="f0d04-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="f0d04-133">Например, если вы хотите, чтобы скрипт принял `string` имя, представляющее его в качестве ввода, вы бы изменили `main` подпись на `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="f0d04-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="f0d04-134">При настройке потока в Power Automate можно указать ввод скрипта как статические значения, [выражения или](/power-automate/use-expressions-in-conditions)динамическое содержимое.</span><span class="sxs-lookup"><span data-stu-id="f0d04-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="f0d04-135">Подробную информацию о разъеме отдельной службы можно найти [в Power Automate Connector.](/connectors/)</span><span class="sxs-lookup"><span data-stu-id="f0d04-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="f0d04-136">При добавлении параметров ввода в `main` функцию скрипта учитывайте следующие надбавки и ограничения.</span><span class="sxs-lookup"><span data-stu-id="f0d04-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="f0d04-137">Первый параметр должен быть `ExcelScript.Workbook` типа.</span><span class="sxs-lookup"><span data-stu-id="f0d04-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="f0d04-138">Его название параметра не имеет значения.</span><span class="sxs-lookup"><span data-stu-id="f0d04-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="f0d04-139">Каждый параметр должен иметь тип (например, `string` или `number` ).</span><span class="sxs-lookup"><span data-stu-id="f0d04-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="f0d04-140">Основные типы `string` , , , и `number` `boolean` `unknown` `object` `undefined` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="f0d04-140">The basic types `string`, `number`, `boolean`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="f0d04-141">Поддерживаются массивы ранее перечисленных базовых типов.</span><span class="sxs-lookup"><span data-stu-id="f0d04-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="f0d04-142">Вложенные массивы поддерживаются в качестве параметров (но не в качестве типов возврата).</span><span class="sxs-lookup"><span data-stu-id="f0d04-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="f0d04-143">Типы союзов допускаются, если они являются союзом букватов, принадлежащих к одному типу `"Left" | "Right"` (например).</span><span class="sxs-lookup"><span data-stu-id="f0d04-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="f0d04-144">Поддерживаются также союзы поддерживаемого типа с неопределенными `string | undefined` (например).</span><span class="sxs-lookup"><span data-stu-id="f0d04-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="f0d04-145">Типы объектов допускаются, если они содержат свойства `string` `number` типа, `boolean` поддерживаемые массивы или другие поддерживаемые объекты.</span><span class="sxs-lookup"><span data-stu-id="f0d04-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="f0d04-146">В следующем примере показаны вложенные объекты, которые поддерживаются в качестве типов параметров:</span><span class="sxs-lookup"><span data-stu-id="f0d04-146">The following example shows nested objects that are supported as parameter types:</span></span>

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. <span data-ttu-id="f0d04-147">Объекты должны иметь свой интерфейс или определение класса, определенные в скрипте.</span><span class="sxs-lookup"><span data-stu-id="f0d04-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="f0d04-148">Объект также может быть определен анонимно, как в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="f0d04-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="f0d04-149">Дополнительные параметры разрешены и могут быть обозначены как таковые с помощью дополнительного `?` модификатора (например, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="f0d04-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="f0d04-150">Разрешены значения параметров по умолчанию `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` (например.</span><span class="sxs-lookup"><span data-stu-id="f0d04-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="return-data-from-a-script"></a><span data-ttu-id="f0d04-151">Возврат данных из скрипта</span><span class="sxs-lookup"><span data-stu-id="f0d04-151">Return data from a script</span></span>

<span data-ttu-id="f0d04-152">Скрипты могут возвращать данные из рабочей книги, которые будут использоваться в качестве динамического содержимого в Power Automate потоке.</span><span class="sxs-lookup"><span data-stu-id="f0d04-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="f0d04-153">Как и в случае с параметрами Power Automate, он устанавливает некоторые ограничения на тип возврата.</span><span class="sxs-lookup"><span data-stu-id="f0d04-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="f0d04-154">Основные типы `string` , , , и `number` `boolean` `void` `undefined` поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="f0d04-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="f0d04-155">Типы союзов, используемые в качестве типов возврата, следуют тем же ограничениям, что и при использовании в качестве параметров скрипта.</span><span class="sxs-lookup"><span data-stu-id="f0d04-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="f0d04-156">Типы массивов разрешены, если они `string` `number` типа, или `boolean` .</span><span class="sxs-lookup"><span data-stu-id="f0d04-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="f0d04-157">Они также допускаются, если тип поддерживается союзом или поддерживается буквальным типом.</span><span class="sxs-lookup"><span data-stu-id="f0d04-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="f0d04-158">Типы объектов, используемые в качестве типов возврата, следуют тем же ограничениям, что и при использовании в качестве параметров скрипта.</span><span class="sxs-lookup"><span data-stu-id="f0d04-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="f0d04-159">Неявная ввод поддерживается, хотя она должна следовать тем же правилам, что и определенный тип.</span><span class="sxs-lookup"><span data-stu-id="f0d04-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="f0d04-160">Пример</span><span class="sxs-lookup"><span data-stu-id="f0d04-160">Example</span></span>

<span data-ttu-id="f0d04-161">На следующем скриншоте показан Power Automate поток, который срабатывает [всякий раз, GitHub](https://github.com/) вам назначается проблема с номером.</span><span class="sxs-lookup"><span data-stu-id="f0d04-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="f0d04-162">Поток выполняет скрипт, который добавляет проблему в таблицу в Excel книги.</span><span class="sxs-lookup"><span data-stu-id="f0d04-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="f0d04-163">Если в этой таблице есть пять или более проблем, поток отправляет напоминание по электронной почте.</span><span class="sxs-lookup"><span data-stu-id="f0d04-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Редактор Power Automate потока, показывающий поток примера":::

<span data-ttu-id="f0d04-165">Функция `main` скрипта определяет идентификатор проблемы и название вопроса в качестве входных параметров, а скрипт возвращает количество строк в таблице проблем.</span><span class="sxs-lookup"><span data-stu-id="f0d04-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a><span data-ttu-id="f0d04-166">См. также</span><span class="sxs-lookup"><span data-stu-id="f0d04-166">See also</span></span>

- [<span data-ttu-id="f0d04-167">Вы запустите Office сценарии в Excel в Интернете с Power Automate</span><span class="sxs-lookup"><span data-stu-id="f0d04-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="f0d04-168">Передача данных сценариям в автоматически запускаемых рабочих процессах Power Automate</span><span class="sxs-lookup"><span data-stu-id="f0d04-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="f0d04-169">Возвращение данных из сценария в автоматически запускаемый поток Power Automate</span><span class="sxs-lookup"><span data-stu-id="f0d04-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="f0d04-170">Информация о устранении неполадок для Power Automate с помощью Office скриптов</span><span class="sxs-lookup"><span data-stu-id="f0d04-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="f0d04-171">Начало работы с Power Automate</span><span class="sxs-lookup"><span data-stu-id="f0d04-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="f0d04-172">Excel Онлайн (Бизнес) разъем справочная документация</span><span class="sxs-lookup"><span data-stu-id="f0d04-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
