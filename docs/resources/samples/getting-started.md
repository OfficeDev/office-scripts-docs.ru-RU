---
title: Начало работы с Office скриптами
description: Основные принципы Office скриптов, включая шаблоны доступа, среды и скриптов.
ms.date: 04/01/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: d30c4fb4523c49b559e057eede4d5de162b74f9c
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232762"
---
# <a name="getting-started"></a>Начало работы

В этом разделе приводится подробная информация об основах Office скриптов, включая доступ, среду, основы скрипта и несколько базовых шаблонов сценариев.

## <a name="environment-setup"></a>Настройка среды

Узнайте об основах доступа, среды и редактора сценариев.

[![Основы приложения Office скриптов](../../images/getting-started-env.png)](https://youtu.be/vvCtxsjPxo8 "Основы приложения Office скриптов")

### <a name="access"></a>Access

Office Скрипты требуют параметров администратора, доступных для Microsoft 365 администратора в **Параметры**  >  **параметров Org** Office  >  **Скрипты**. По умолчанию он включен для всех пользователей. Существует два подпараметров, которые администратор может включить и отключить.

* Возможность обмена скриптами в организации
* Возможность использования скриптов в Power Automate

Вы можете сказать, имеете ли вы доступ к Office скриптам, открывая файл в Excel в Интернете (браузере) и видя, появится ли вкладка **Automate** в ленте Excel или нет.
Если вы все еще не видите вкладку **Automate,** проверьте этот раздел [устранения неполадок.](../../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable)

### <a name="availability"></a>Доступность

Office Скрипты доступны только в Excel в Интернете лицензий Enterprise E3+ (учетные записи потребителей и E1 не поддерживаются). Office Скрипты еще не поддерживаются в Excel на Windows и Mac.

### <a name="scripts-and-editor"></a>Сценарии и редактор

Редактор кода встроен прямо в Excel в Интернете (онлайн-версия). Если вы использовали редакторы Visual Studio Code или Sublime, этот опыт редактирования будет очень похож.
Большинство клавиш ярлыков, Visual Studio Code редактор использует работу и в Office сценариев. Ознакомьтесь со следующими раздатями клавиш ярлыков.

* [macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)
* [Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)

#### <a name="key-things-to-note"></a>Ключевые моменты, которые следует отметить

* Office Скрипты доступны только для файлов, хранимых на OneDrive для бизнеса, SharePoint сайтах и сайтах Team.
* Редактор не показывает расширение скрипта. На самом деле это файлы TypeScript, но они хранятся с пользовательским расширением под названием `.osts` .
* Скрипты хранятся в вашей OneDrive для бизнеса `My Files/Documents/OfficeScripts` папке. Управление этой папкой не требуется. Со своей стороны, этот аспект можно игнорировать, так как редактор управляет просмотром и редактированием.
* Скрипты не хранятся как часть Excel файлов. Они хранятся отдельно.
* Вы можете поделиться сценарием с файлом Excel, что фактически означает, что вы связываете сценарий с файлом, а не привязывая его. Тот, кто имеет доступ к Excel файлу, также сможет **просматривать,** **запускать** или делать **копию** сценария. Это ключевое отличие по сравнению с макросами VBA.
* Если вы не поделитесь своими скриптами, никто не сможет получить к нему доступ, так как он находится в вашей собственной библиотеке.
* Сценарии нельзя связывать с локальным диском или настраиваемой облачной локацией. Office Скрипты распознают и запускает только сценарий, который находится в предварительном расположении (OneDrive папке, упомянутой выше), или общие скрипты.
* Во время редактирования файлы временно сохраняются в браузере, но перед закрытием окна Excel необходимо сохранить его в OneDrive месте. Не забудьте сохранить файл после изменений.

## <a name="gentle-introduction-to-scripting"></a>Мягкое введение в сценарии

Office Скрипты — это автономные скрипты, написанные на языке TypeScript, содержащие инструкции по выполнению некоторой автоматизации в отношении выбранной Excel книги. Все инструкции по автоматизации содержатся в сценарии, и скрипты не могут вызывать или вызывать другие сценарии. Все скрипты хранятся в автономных файлах и хранятся в папке OneDrive пользователя. Вы можете записать новый сценарий, изменить записанный сценарий или написать совершенно новый сценарий с нуля, все это в встроенной интерфейс редактора. Самая лучшая часть Office в том, что они не нуждаются в дальнейшей настройке от пользователей. Нет внешних библиотек, веб-страниц или элементов пользовательского интерфейса, установки и т.д. Вся настройка среды обрабатывается Office скриптами и позволяет легко и быстро получить доступ к автоматизации с помощью простого интерфейса API.

Некоторые из основных понятий, полезных для понимания редактирования и навигации по сценариям, включают:

* Основной синтаксис языка TypeScript
* Понимание `main` функций и аргументов
* Объекты и иерархия, методы, свойства
* Коллекция (массив): навигация и операции
* Определения типа
* Среда: запись/редактирование, запуск, изучение результатов, совместное

В этом видео и разделе подробно объясняются некоторые из этих понятий.

[![Основы Office скриптов](../../images/getting-started-v_script.png)](https://youtu.be/8Zsrc1uaiiU "Основы скриптов")

### <a name="language-typescript"></a>Язык: TypeScript

[Office](../../index.md) скрипты написаны с помощью языка [TypeScript](https://www.typescriptlang.org/), который является языком с открытым исходным кодом, который строится на JavaScript (один из наиболее используемых в мире) путем добавления статических определений типов. Как говорится на веб-сайте, предопишите форму объекта, предоставляя лучшую документацию и позволяя TypeScript проверять правильность работы `Types` кода.

Сам синтаксис языка написан с помощью [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) с дополнительными вводами, определенными в скрипте с помощью соглашений TypeScript. По большей части можно придумать сценарии Office, написанные в JavaScript. Важно, чтобы вы поняли основы языка JavaScript, чтобы начать Office скрипты; хотя вам не нужно быть опытным в этом, чтобы начать свой путь автоматизации. С помощью Office скриптов можно понять утверждения сценариев, так как в него включены комментарии к коду, и вы можете следовать за ними и внести небольшие изменения.

Office API скриптов, которые позволяют скрипту взаимодействовать с Excel, предназначены для конечных пользователей, у которых может быть не так много фона кодирования. API можно вызывать синхронно, и вам не нужно знать расширенные темы, такие как обещания или вызовы. Office Разработка API скриптов обеспечивает:

* Простая объектная модель с методами, getters/setters.
* Простые в доступе коллекции объектов в качестве обычных массивов.
* Простые параметры обработки ошибок.
* Оптимизированная производительность для выбора сценариев, помогающих пользователям сосредоточиться на сценарии.

### <a name="main-function-the-scripts-starting-point"></a>`main` функция: отправная точка сценария

Office Выполнение скриптов начинается с `main` функции. Скрипт — это один файл, содержащий одну или несколько функций, а также объявления типов, интерфейсов, переменных и т.д. Чтобы следовать вместе со сценарием, начните с функции, Excel всегда сначала вызывает функцию `main` `main` при выполнении любого сценария. Функция всегда будет иметь по крайней мере один аргумент (или параметр) с именем , которое является просто переменной имя, определяя текущую книгу, против которой `main` `workbook` работает сценарий. Дополнительные аргументы для использования можно определить при Power Automate (автономном режиме).

* `function main(workbook: ExcelScript.Workbook)`

Скрипт можно организовать на более мелкие функции, которые помогают в повторном использования кода, ясности и т.д. Другие функции могут быть внутри или за пределами основной функции, но всегда в одном файле. Сценарий является автономным и может использовать только функции, определенные в одном файле. Скрипты не могут вызывать или вызывать Office скрипта.

Итак, в сводке:

* Функция `main` является точкой входа для любого сценария. Когда функция выполняется, Excel вызывает эту основную функцию, предоставляя книгу в качестве первого параметра.
* Важно сохранить первый аргумент и его объявление `workbook` типа, как он появляется. Вы можете добавить новые аргументы в функцию (см. следующий раздел), но сохранить первый `main` аргумент как есть.

:::image type="content" source="../../images/getting-started-main-introduction.png" alt-text="Основная функция — точка входа скрипта":::

#### <a name="send-or-receive-data-from-other-apps"></a>Отправка или получение данных из других приложений

Вы можете подключить Excel к другим частям организации, запуская [скрипты в Power Automate.](https://flow.microsoft.com) Дополнительные новости о [запуске Office скриптов в Power Automate потоках.](../../develop/power-automate-integration.md)

Способ получения или отправки данных из и Excel через `main` функцию. Думайте об этом как о шлюзе информации, который позволяет описывать и использовать в скрипте входящие и исходяющие данные. Вы можете получать данные из-за пределов скрипта с помощью типа данных и возвращать любые данные, признанные TypeScript, такие как , или любые объекты в виде интерфейсов, определенных в `string` `string` `number` `boolean` скрипте.

:::image type="content" source="../../images/getting-started-data-in-out.png" alt-text="Входные данные и выходы сценария":::

#### <a name="use-functions-to-organize-and-reuse-code"></a>Использование функций для организации и повторного использования кода

Вы можете использовать функции для организации и повторного использования кода в скрипте.

:::image type="content" source="../../images/getting-started-use-functions.png" alt-text="Использование функций в скрипте":::

### <a name="objects-hierarchy-methods-properties-collections"></a>Объекты, иерархия, методы, свойства, коллекции

Все объектные Excel определяются в иерархической структуре объектов, начиная с объекта книги типа `ExcelScript.Workbook` . Объект может содержать методы, свойства и другие объекты в нем. Объекты связаны друг с другом с помощью методов. Метод объекта может возвращать другой объект или коллекцию объектов. Использование функции IntelliSense (завершение кода) — отличный способ изучить иерархию объектов. Вы также можете использовать официальный [сайт справочной документации,](/javascript/api/office-scripts/overview) чтобы следовать за отношениями между объектами.

Объект [—](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Object) это коллекция свойств, а свойство — это связь между именем (или ключом) и значением. Значение свойства может быть функцией, в этом случае свойство называется методом. В случае объектной модели Office Scripts объект представляет вещь в файле Excel, с чем взаимодействуют пользователи, такие как диаграмма, гиперссылка, поворотная таблица и т. д. Он также может представлять поведение объекта, например атрибуты защиты таблицы.

Тема объектов и свойств TypeScript по сравнению с методами довольно глубока. Чтобы начать работу со сценарием и быть продуктивным, можно вспомнить несколько основных вещей:

* Как объекты, так и свойства доступны с помощью (точка) нотации, а объект слева и свойства или метода `.` `.` на правой стороне. Примеры: `hyperlink.address` , `range.getAddress()` .
* Свойства являются scalar в природе (строки, booleans, номера). Например, имя книги, положение таблицы, значение того, имеет ли таблица подножку или нет.
* Методы "вызываются" или "выполняются" с помощью скобок с открытыми закрытиями. Пример: `table.delete()`. Иногда аргумент передается функции, включив их между открытыми скобами: `range.setValue('Hello')` . Вы можете передать множество аргументов функции (как определено ее контрактом/подписью) и отделить их с помощью `,` .  Пример: `worksheet.addTable('A1:D6', true)`. Вы можете передавать аргументы любого типа, как того требует метод, например строки, число, boolean или даже другие объекты, например, где находится объект, созданный в другом месте `worksheet.addTable(targetRange, true)` `targetRange` сценария.
* Методы могут возвращать свойство scalar (имя, адрес и т. д.) или другой объект (диапазон, диаграмма) или вообще ничего не возвращать (например, в случае с `delete` методами). Вы получаете то, что возвращает метод, объявив переменную или назначив существующую переменную. Вы можете видеть, что на левой стороне заявления, таких как `const table = worksheet.addTable('A1:D6', true)` .
* По большей части объектная модель Office scripts состоит из объектов с методами, связывающие различные части Excel объектной модели. Очень редко встречаются свойства, которые имеют масштабарные или объектные значения.
* В Office Скрипты метод Excel объектной модели должен содержать скобки с открытыми закрытиями. Использование без них методов запрещено (например, назначение метода переменной).

Рассмотрим несколько методов на `workbook` объекте.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Return a boolean (true or false) setting of whether the workbook is set to auto-save or not. 
    const autoSave = workbook.getAutoSave(); 
    // Get workbook name.
    const name = workbook.getName();
    // Get active cell range object.
    const cell = workbook.getActiveCell();
    // Get table named SALES.
    const cell = workbook.getTable('SALES');
    // Get all slicer objects.
    const slicers = workbook.getSlicers();
}
```

В этом примере:

* Методы объекта, такие как и возвращение свойства `workbook` `getAutoSave()` `getName()` scalar (строка, номер, boolean).
* Такие методы, как `getActiveCell()` возвращение другого объекта.
* Метод принимает аргумент (имя таблицы в этом случае) и возвращает определенную таблицу `getTable()` в книге.
* Метод возвращает массив (который во многих местах именуется коллекцией) всех объектов `getSlicers()` slicer в книге.

Вы заметите, что все эти методы имеют префикс, который является просто конвенцией, используемой в объектной модели Office Scripts, чтобы передать, что метод возвращает `get` что-то. Они также часто называются "getters".

В следующем примере мы увидим два других типа методов:

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get a worksheet named 'Sheet1.
    const sheet = workbook.getWorksheet('Sheet1'); 
    // Set name to SALES.
    sheet.setName('SALES');
    // Position the worksheet at the beginning.
    sheet.setPosition(0);
}
```

В этом примере:

* Метод `setName()` задает новое имя для таблицы. `setPosition()` задает позицию для первой ячейки.
* Такие методы изменяют Excel файл, устанавливая свойство или поведение книги. Эти методы называются "setters".
* Как правило, "сеттеры" имеют компаньона "getter", например, и , оба `worksheet.getPosition` `worksheet.setPosition` из которых являются методами.

#### <a name="undefined-and-null-primitive-types"></a>`undefined` и `null` примитивные типы

Ниже приводится два примитивных типа данных, которые необходимо знать:

1. Это значение [`null`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/null) представляет преднамеренное отсутствие какого-либо значения объекта. Это одно из примитивных значений JavaScript, которое используется для того, чтобы указать, что переменная не имеет значения.
1. Переменная, не назначенная значению, имеет тип [`undefined`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/undefined) . Метод или утверждение также могут возвращаться, если оцениваемая переменная не имеет `undefined` назначенного значения.

Эти два типа возникают как часть обработки ошибок и могут вызвать довольно много головной боли, если не обрабатываются должным образом. К счастью, TypeScript/JavaScript предлагает способ проверить, имеет ли переменная тип `undefined` или `null` . Мы поговорим о некоторых из этих проверок в более поздних разделах, включая обработку ошибок.

#### <a name="method-chaining"></a>Цепочка метода

Чтобы сократить код, можно использовать пунктовое нотирование, чтобы подключить возвращаемые из метода объекты. Иногда этот метод упрощает чтение и управление кодом. Тем не менее, есть несколько вещей, которые следует знать. Рассмотрим следующие примеры.

Следующий код получает активную ячейку и следующую ячейку, а затем задает значение. Это хороший кандидат для использования цепочки, поскольку этот код будет успешной все время.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getActiveCell().getOffsetRange(0,1).setValue('Next cell');
}
```

Однако следующий код (который получает таблицу с именем **SALES** и включает свой полосатой стиль столбца) имеет проблему.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  workbook.getTable('SALES').setShowBandedColumns(true);
}
```

Что делать, если таблица **SALES** не существует? Сценарий не работает с ошибкой (показано далее), так как возвращается (это тип JavaScript, указывающий на отсутствие `getTable('SALES')` `undefined` таблицы, например **SALES).** Вызов метода `setShowBandedColumns` не `undefined` имеет смысла, то есть, и, следовательно, `undefined.setShowBandedColumns(true)` сценарий заканчивается ошибкой.

```text
Line 2: Cannot read property 'setShowBandedColumns' of undefined
```

Для обработки [](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Optional_chaining) этого условия можно использовать необязательный оператор цепочки, который предоставляет способ упростить доступ к значениям через подключенные объекты, если возможно, что ссылкой или методом может быть или (который является способом `undefined` JavaScript, указывающим неназваванный или несущестущий объект или `null` результат).

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // This line will not fail as the setShowBandedColumns method is executed only if the SALES table is present.
    workbook.getTable('SALES')?.setShowBandedColumns(true); 
}
```

Если вы хотите обрабатывать несущестутные условия объекта или тип, возвращаемый методом, то лучше назначить возвращаемого значения из метода и обрабатывать его `undefined` отдельно.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const salesTable = workbook.getTable('SALES');
    if (salesTable) {
        salesTable.setShowBandedColumns(true);
    } else { 
        // Handle this condition.
    }
}
```

#### <a name="get-object-reference"></a>Получить ссылку на объект

Объект `workbook` предоставляется в `main` функции. Вы можете начать использовать объект `workbook` и получить доступ к его методам напрямую.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get workbook name.
    const name = workbook.getName();
    // Display name to console.
    console.log(name);
}
```

Для использования всех других объектов в книге начните с объекта и спуститься по иерархии, пока не доберется до `workbook` объекта, который вы ищете. Вы можете получить ссылку на объект, извлекая объект с помощью его метода или извлекая коллекцию объектов, как показано `get` ниже:

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    const sheet = workbook.getActiveWorksheet();
    // Fetch using an ID or key.
    const sheet = workbook.getWorksheet('SomeSheetName');
    // Invoke methods on the object.
    sheet.setPosition(0); 
    
    // Get collection of methods.
    const tables = sheet.getTables();
    console.log('Total tables in this sheet: ' + tables.length);
}
```

#### <a name="check-if-an-object-exists-then-delete-and-add"></a>Проверьте, существует ли объект, затем удалите и добавьте

Для создания объекта, скажем с заранее, всегда лучше удалить аналогичный объект, который может существовать, а затем добавить его. Это можно сделать с помощью следующего шаблона.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added. 
  let name = "Index";
  // Check if the worksheet already exists. If not, add the worksheet.
  let sheet = workbook.getWorksheet('Index');
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Call the delete method on the object to remove it. 
    sheet.delete();
  } 
    // Add a blank worksheet. 
  console.log(`Adding the worksheet named  ${name}.`)
  const indexSheet = workbook.addWorksheet("Index");
}

```

Кроме того, для удаления объекта, который может существовать или не существует, используйте следующий шаблон.

```TypeScript
    // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
    workbook.getWorksheet('Index')?.delete(); 
```

#### <a name="note-about-adding-an-object"></a>Примечание о добавлении объекта

Чтобы создать, вставить или добавить объект, например срез, таблицу поворота, таблицу, таблицу и т. д., используйте соответствующий метод **add_Object_.** Такой метод доступен на родительском объекте. Например, метод `addChart()` доступен на `worksheet` объекте. Метод **add_Object_** возвращает объект, который он создает. Получите возвращенные значения и используйте его позже в скрипте.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Add object and get a reference to it. 
  const indexSheet = workbook.addWorksheet("Index");
  // Use it elsewhere in the script 
  console.log(indexSheet.getPosition());
}

```

Кроме того, для удаления объекта, который может существовать или не существует, используйте этот шаблон:

```TypeScript
    workbook.getWorksheet('Index')?.delete(); // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
```

#### <a name="collections"></a>Коллекции

Коллекции — это объекты, такие как таблицы, диаграммы, столбцы и т. д., которые можно получить в качестве массива и итерировать для обработки. Вы можете получить коллекцию с помощью соответствующего метода и обрабатывать данные в цикле с помощью одного из многих методов обхода массива `get` TypeScript, таких как:

* [`for` или `while`](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
* [`for..of`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/for...of)
* [`forEach`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/forEach)

* [Языковые основы массивов](https://developer.mozilla.org//docs/Learn/JavaScript/First_steps/Arrays)

В этом скрипте показано, как использовать коллекции, поддерживаемые Office API скриптов. Она красят каждую вкладку таблицы в файле случайным цветом.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get all sheets as a collection.
  const sheets = workbook.getWorksheets();
  const names = sheets.map ((sheet) => sheet.getName());
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  // Get information from specific sheets within the collection.
  console.log(`First sheet name is: ${names[0]}`);
  if (sheets.length > 1) {
    console.log(`Last sheet's Id is: ${sheets[sheets.length -1].getId()}`);
  }
  // Color each worksheet with random color.
  for (const sheet of sheets) {
    sheet.setTabColor(`#${Math.random().toString(16).substr(-6)}`);
  }
}
```

## <a name="type-declarations"></a>Объявления типа

Объявления типа помогают пользователям понять тип переменной, с какой переменной они имеют дело. Это помогает при автоматическом завершении методов и помогает в проверках качества времени разработки.

Объявления типов в скрипте можно найти в различных местах, включая объявление функций, переменную декларацию, IntelliSense определения и т.д.

Примеры:

* `function main(workbook: ExcelScript.Workbook)`
* `let myRange: ExcelScript.Range;`
* `function getMaxAmount(range: ExcelScript.Range): number`

Вы можете легко определить типы в редакторе кода, так как он обычно отчетливо отображается в другом цвете. Обычно `:` двоеточие предшествует объявлению типа.  

Типы записи могут быть необязательными в TypeScript, так как вывод типа позволяет получать много энергии без написания дополнительного кода. По большей части язык TypeScript хорош для того, чтобы сделать вывод о типах переменных. Однако в некоторых Office скрипты требуют явного определения типовых деклараций, если язык не может четко определить тип. Кроме того, явный или неявный не допускается в `any` Office Скрипт. Подробнее об этом см. ниже.

### <a name="excelscript-types"></a>`ExcelScript` типы

В Office скрипты будут использовать следующие типы.

* Типы родного `number` языка, такие как , , , и `string` `object` `boolean` `null` т.д.
* Excel Типы API. Они начинаются `ExcelScript` с . Например, `ExcelScript.Range` и `ExcelScript.Table` т.д.
* Любые настраиваемые интерфейсы, которые вы могли определить в сценарии с помощью `interface` заявлений.

Далее см. примеры каждой из этих групп.

**_Типы родного языка_**

В следующем примере обратите внимание на места, где `string` `number` и были `boolean` использованы. Это родные типы языков **TypeScript.**

```TypeScript
function main(workbook: ExcelScript.Workbook)
{
  const table = workbook.getActiveWorksheet().getTables()[0];
  const sales = table.getColumnByName('Sales').getRange().getValues();
  console.log(sales);
  // Add 100 to each value.
  const revisedSales = salesAs1DArray.map(data => data as number + 100);
  // Add a column.
  table.addColumn(-1, revisedSales);  
}
/**
 * Extract a column from 2D array and return result.
 */
function extractColumn(data: (string | number | boolean)[][], index: number): (string | number | boolean)[] {

  const column = data.map((row) => {
    return row[index];
  })
  return column;
}
/**
 * Convert a flat array into a 2D array that can be used as range column.
 */
function convertColumnTo2D(data: (string | number | boolean)[]): (string | number | boolean)[][] {

  const columnAs2D = data.map((row) => {
    return [row];
  })
  return columnAs2D;
}
```

**_Типы ExcelScript_**

В следующем примере функция помощника принимает два аргумента. Первый — переменная `sheet` типа `ExcelScript.Worksheet` типа.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for update.
    if (usedRange) { 
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);      
    targetRange.setValues([data]);
    return;
}
```

**_Настраиваемые типы_**

Пользовательский интерфейс `ReportImages` используется для возврата изображений в другое действие потока. В `main` декларации функций содержится инструкция по указанию TypeScript о том, что объект этого `: ReportImages` типа возвращается.

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  let chart = workbook.getWorksheet("Sheet1").getCharts()[0];
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  
  const chartImage = chart.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

### <a name="type-assertion-overriding-the-type"></a>Утверждение типа (переопределение типа)

Как говорится в документации [TypeScript,](https://www.typescriptlang.org/docs/handbook/basic-types.html#type-assertions) "Иногда вы будете в конечном итоге в ситуации, когда вы будете знать больше о значении, чем TypeScript делает. Обычно это происходит, когда вы знаете, что тип какой-либо сущности может быть более конкретным, чем его текущий тип. Тип утверждений — это способ сказать компилятору "доверяйте мне, я знаю, что делаю". Утверждение типа похоже на тип, отлитый на других языках, но не выполняет никакой специальной проверки или реструктуризации данных. Он не влияет на время работы и используется исключительно компилятором".

Вы можете утверждать тип с помощью ключевого слова `as` или с помощью угловой скобки, как показано в следующем коде.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let data = workbook.getActiveCell().getValue();
  // Since the add10 function only accepts number, assert data's type as number, otherwise the script cannot be run.
  const answer1 = add10(data as number);
  const answer2 = add10(<number> data);
}

function add10(data: number) { 
  return data + 10;
}
```

#### <a name="any-type-in-the-script"></a>'any' type in the script

На [веб-сайте TypeScript говорится:](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)

  В некоторых ситуациях не все сведения о типе доступны, а для ее объявления необходимо приложить неоцелесообразные усилия. Они могут возникать для значений из кода, который был написан без TypeScript или 3-й библиотеки сторон. В этих случаях может потребоваться отказаться от проверки типа. Для этого мы пометим эти значения с помощью `any` типа:

  ```TypeScript
  declare function getValue(key: string): any;
  // OK, return value of 'getValue' is not checked
  const str: string = getValue("myString");
  ```

**Явный `any` не допускается**

```TypeScript
// This is not allowed
let someVariable: any; 
```

Тип `any` представляет проблемы с тем, как Office скрипты обрабатывает Excel API. Это вызывает проблемы, когда переменные отправляются Excel API для обработки. Знание типа переменных, используемых в сценарии, имеет важное значение для обработки скрипта и, следовательно, запрещается явное определение любой `any` переменной с типом. Вы получите ошибку времени компилирования (ошибка перед запуском скрипта), если в скрипте есть переменная с `any` объявленным типом. Вы увидите ошибку и в редакторе.

:::image type="content" source="../../images/getting-started-eanyi.png" alt-text="Явные ошибки &quot;любой&quot;":::

:::image type="content" source="../../images/getting-started-expany.png" alt-text="Явные &quot;любые&quot; ошибки, показанные в выходе":::

В коде, отображаемом на предыдущем изображении, указывается, что `[5, 16] Explicit Any is not allowed` строка 5 столбца 16 объявляет `any` тип. Это поможет найти строку кода, содержаную ошибку.

Чтобы обойти эту проблему, всегда объявите тип переменной.

Если вы не уверены в типе переменной, один классный прием в TypeScript позволяет определить [типы союзов.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Это может быть использовано для переменных для удержания значений диапазона, которые могут быть многих типов.

```TypeScript
// Define value as a union type rather than 'any' type.
let value: (string | number | boolean);
value = someValue_from_another_source;
//...
someRange.setValue(value);
```

### <a name="type-inference"></a>Вывод типа

В TypeScript существует несколько [](https://www.typescriptlang.org/docs/handbook/type-inference.html) мест, где для предоставления сведений о типах используется несколько мест, где нет явной аннотации типа. Например, тип переменной x высмеяется как номер в следующем коде.

```TypeScript
let x = 3;
//  ^ = let x: number
```

Этот вид выводов происходит при инициализации переменных и членов, задании значений параметров по умолчанию и определении типов возврата функций.

### <a name="no-implicit-any-rule"></a>правило no-implicit-any

Сценарий требует, чтобы типы переменных, которые были объявлены явно или неявно. Если компилятор TypeScript не может определить тип переменной (либо из-за явного декларирования типа, либо невозможно сделать вывод о типе), вы получите ошибку времени компиляции (ошибка перед запуском скрипта). Вы увидите ошибку и в редакторе.

:::image type="content" source="../../images/getting-started-iany.png" alt-text="Неявная ошибка &quot;любая&quot;, показанная в редакторе":::

В следующих скриптах имеются ошибки времени компиляции, так как переменные объявляются без типов, а TypeScript не может определить тип на момент объявления.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'value' gets 'any' type
    // because no type is declared.
    let value; 
    // Even when a number type is assigned,
    // the type of 'value' remains any.
    value = 10; 
    // The following statement fails because
    // Office Scripts can't send an argument
    // of type 'any' to Excel for processing.
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'cell' gets 'any' type
    // because no type is defined.
    let cell; 
    cell = workbook.getActiveCell().getValue();
    // Office Scripts can't assign Range type object
    // to a variable of 'any' type.
    console.log(cell.getValue());
    return;
}
```

Чтобы избежать этой ошибки, используйте следующие шаблоны. В каждом случае переменная и ее тип объявляются одновременно.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const value: number = 10; 
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const cell: ExcelScript.Range = workbook.getActiveCell().getValue();
    console.log(cell.getValue()); 
    return;
}
```

## <a name="error-handling"></a>Обработка ошибок

Office Ошибку скриптов можно классифицировать в одну из следующих категорий.

1. Предупреждение о времени компилляции, показанное в редакторе
1. Ошибка времени компиляции, которая появляется при запуске, но возникает до начала выполнения
1. Ошибка при запуске

Предупреждения редактора можно определить с помощью волнистых красных линий в редакторе:

:::image type="content" source="../../images/getting-started-eanyi.png" alt-text="Предупреждение о времени компилляции, показанное в редакторе":::

Иногда вы также можете видеть оранжевые линии предупреждения и серые информационные сообщения. Их следует внимательно исследовать, хотя они не вызывают ошибок.

Невозможно различить между ошибками времени компилирования и временем работы, так как оба сообщения об ошибках выглядят одинаково. Оба они возникают при выполнении сценария. На следующих изображениях покажут примеры ошибки времени компилирования и ошибки времени запуска.

:::image type="content" source="../../images/getting-started-expany.png" alt-text="Пример ошибки времени компилирования":::

:::image type="content" source="../../images/getting-started-error-basic.png" alt-text="Пример ошибки во время запуска":::

В обоих случаях вы увидите номер строки, где произошла ошибка. Затем можно изучить код, устранить проблему и снова запустить.

Ниже приводится несколько наиболее оптимальных методов, чтобы избежать ошибок во время работы.

### <a name="check-for-object-existence-before-deletion"></a>Проверка на наличие объекта перед удалением

Кроме того, для удаления объекта, который может существовать или не существует, используйте этот шаблон:

```TypeScript
// The ? ensures that the delete() API is only invoked if the object exists.
workbook.getWorksheet('Index')?.delete();

// Alternative:
const indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
    indexSheet.delete();
}
```

### <a name="do-pre-checks-at-the-beginning-of-the-script"></a>Предварительные проверки в начале сценария

В качестве наилучшей практики всегда убедитесь, что все входные данные присутствуют в файле Excel перед запуском скрипта. Возможно, вы сделали определенные предположения о том, что объекты присутствуют в книге. Если эти объекты не существуют, скрипт может столкнуться с ошибкой при прочтете объекта или его данных. Вместо того чтобы начинать обработку и ошибки в середине после завершения части обновлений или обработки, лучше сделать все предварительные проверки в начале сценария.

Например, в следующем сценарии необходимо иметь две таблицы с именами Table1 и Table2. Поэтому сценарий проверяет их присутствие и заканчивается заявлением и соответствующим `return` сообщением, если они не присутствуют.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

Если проверка для обеспечения присутствия входных данных происходит в отдельной функции, важно закончить сценарий, выпустив заявление `return` из `main` функции.

В следующем примере функция `main` вызывает `inputPresent` функцию для предварительной проверки. `inputPresent` возвращает boolean (или) с указанием, присутствуют ли все необходимые входные `true` `false` данные или нет. После этого функции должны немедленно издать заявление (то есть изнутри функции) для `main` `return` `main` немедленного окончания скрипта.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }
  return true;
}
```

### <a name="when-to-abort-throw-the-script"></a>Когда прервать `throw` () сценарий  

По большей части вам не нужно прервать `throw` () из сценария. Это связано с тем, что сценарий обычно информирует пользователя о том, что сценарий не удалось выполнить из-за проблемы. В большинстве случае достаточно закончить сценарий сообщением об ошибке и `return` заявлением из `main` функции.

Однако, если сценарий работает в Power Automate, может потребоваться прервать поток, если определенные условия не выполнены. Поэтому важно не ошибться, а издать заявление, чтобы прервать сценарий, чтобы не запускать все последующие утверждения `return` `throw` кода.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    // Abort script.
    throw `Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

Как уже упоминалось в следующем разделе, другой сценарий , когда у вас есть несколько функций, участвующих (вызовы, которые вызовы `main` `functionX` и т.д.), что затрудняет распространение `functionY` ошибки. Прерывание или выкидыш из вложенной функции с сообщением может быть проще, чем возвращать ошибку до и возвращать из сообщения `main` `main` об ошибке.

### <a name="when-to-use-trycatch-throw-exception"></a>Когда использовать try.. catch (исключение броска)

Метод — это способ обнаружения сбой вызова API и обработки [`try..catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) этой ошибки в скрипте. Возможно, важно проверить возвращаемую стоимость API, чтобы убедиться, что он успешно выполнен.

Рассмотрим фрагмент следующего примера.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Somewhere in the script, perform a large data update.
  range.setValues(someLargeValues);

}
```

Вызов `setValues()` может привести к сбою сценария. Возможно, вы захочете обработать это условие в коде и, возможно, настроить сообщение об ошибке или разбить обновление на меньшие единицы и т.д. В этом случае важно знать, что API возвращает ошибку и интерпретирует или обрабатывает эту ошибку.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Please inspect and run again.`);
    console.log(error);
    return; // End script (assuming this is in main function).
}

// OR...

try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Trying a different approach`);
    handleUpdatesInSmallerChunks(someLargeValues);
}

// Continue...
}
```

Другой сценарий — когда основная функция вызывает другую функцию, которая, в свою очередь, вызывает другую функцию (и так далее.), а вызов API, о котором вы заботитесь, происходит в нижней функции. Распространение ошибки в любом случае может `main` оказаться нецелесообразным или удобным. В этом случае наиболее удобно бросать ошибку в нижнюю функцию.

```TypeScript

function main(workbook: ExcelScript.Workbook) {
    ...
    updateRangeInChunks(sheet.getRange("B1"), data);
    ...
}

function updateRangeInChunks(
    ...
    updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
    ...
}

function updateTargetRange(
      targetCell: ExcelScript.Range,
      values: (string | boolean | number)[][]
    ) {
    const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
    console.log(`Updating the range: ${targetRange.getAddress()}`);
    try {
      targetRange.setValues(values);
    } catch (e) {
      throw `Error while updating the whole range: ${JSON.stringify(e)}`;
    }
    return;
}
```

*Предупреждение.* `try..catch` Использование внутри цикла притормозит сценарий. Избегайте использования этого внутри или вокруг циклов.

## <a name="basic-performance-considerations"></a>Основные соображения производительности

### <a name="avoid-slow-operations-in-the-loop"></a>Избегайте медленных операций в цикле

Некоторые операции, когда они делаются внутри или вокруг таких циклов, как `for` , , , и `for..of` `map` `forEach` т.д. может привести к медленной производительности. Избегайте следующих категорий API.

* `get*` API

Ознакомьтесь со всеми данными, которые необходимы за пределами цикла, а не чтением их внутри цикла. Иногда трудно избежать чтения внутри циклов; в этом случае убедитесь, что количество циклов не слишком большое или управление ими в пакетах, чтобы избежать необходимости цикла через большую структуру данных.

**Примечание.** Если диапазон/данные, с которые вы имеете дело, достаточно велик (скажем, >ячейки 100K), вам может потребоваться использовать расширенные методы, такие как разрыв чтения и записи на несколько фрагментов. Следующее видео действительно для установки данных малого среднего размера. Для большого количества данных обратитесь к [сценарию записи расширенных данных.](write-large-dataset.md)

[![Видео, предоставляющая совет по оптимизации чтения и записи](../../images/getting-started-v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "Совет по оптимизации чтения и записи")

* `console.log` заявление (см. следующий пример)

```TypeScript
// Color each cell with random color.
for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
        range
            .getCell(row, col)
            .getFormat()
            .getFill()
            .setColor(`#${Math.random().toString(16).substr(-6)}`);
        /* Avoid such console.log inside loop */
        // console.log("Updating" + range.getCell(row, col).getAddress());
    }
}
```

* `try {} catch ()` заявление

Избегайте циклов `for` обработки исключений. Внутренние и внешние циклы.

## <a name="note-to-vba-developers"></a>Примечание для разработчиков VBA

Язык TypeScript отличается от VBA как синтаксически, так и в соглашениях именования.

Ознакомьтесь со следующими эквивалентными фрагментами.

```vba
Worksheets("Sheet1").Range("A1:G37").Clear
```

```TypeScript
workbook.getWorksheet('Sheet1').getRange('A1:G37').clear(ExcelScript.ClearApplyTo.all);
```

Несколько вещей, которые нужно вызвать о TypeScript:

* Вы можете заметить, что для выполнения всех методов необходимо иметь скобки с открытыми дверями. Аргументы передаются одинаково, но для выполнения могут потребоваться некоторые аргументы (то есть необходимые или необязательные).
* Конвенция именования следует camelCase вместо pascalCase convention.
* Методы обычно имеют `get` или `set` префиксы, указывающие, является ли это чтение или написание участников объекта.
* Блоки кода определяются и определяются фигурными скобами с открытыми дверями: `{` `}` . Блоки необходимы для `if` условий, `while` заявлений, `for` циклов, определений функций и т.д.
* Функции могут вызывать другие функции, и вы даже можете определить функции в пределах функции.

В целом TypeScript — это другой язык, и между ними мало общего. Однако В API Office сценариев используется аналогичная иерархия терминологии и модели данных (объектная модель) в качестве API VBA, что должно помочь вам ориентироваться.
