---
title: Когда использовать Power Query или сценарии Office
description: Сценарии, наиболее подходящие для платформ Power Query и Office скриптов.
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: e91077d635d66dde692c129bdd4b2f32657d5283
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585907"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Когда использовать Power Query или сценарии Office

[Power Query](https://powerquery.microsoft.com) и Office скрипты являются мощными решениями автоматизации для Excel. Оба решения Excel пользователям очищать и преобразовывать данные в книгах. Один Power Query или Office можно обновить и повторно использовать новые данные для получения согласованных результатов, что экономит время и позволяет быстрее работать с полученными сведениями.

В этой статье представлен общий обзор того, когда можно выступают за одну платформу над другой. Как правило, Power Query хорошо для вывода и преобразования данных из больших, внешних источников данных и Office Скрипты хорошо для быстрых, Excel-ориентированных решений и [Power Automate](../develop/power-automate-integration.md) интеграций.

## <a name="large-data-sources-and-data-retrieval-power-query"></a>Большие источники данных и сбор данных: Power Query

Рекомендуется Power Query при работе с источниками данных с поддерживаемых платформ.

Power Query [имеет встроенные подключения к](https://powerquery.microsoft.com/connectors/) данным сотен источников. Power Query специально разработан для задач по сбору, преобразованию и комбинации данных. Если вам нужны данные из одного из этих источников, Power Query позволяет без кода вводить эти данные в Excel в нужной форме.

Эти Power Query предназначены для больших наборов данных. Они не имеют [таких же](../testing/platform-limits.md) ограничений передачи, как Power Automate или Excel для Интернета.

Office Scripts предлагает легкое решение для небольших источников данных или источников данных, не Power Query соединители. Это включает использование [API `fetch` REST](../develop/external-calls.md) или получение сведений из различных источников данных, например Teams [адаптивной карты](../resources/scenarios/task-reminders.md).

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>Форматирование, визуализация и программный контроль: Office скрипты

Мы рекомендуем Office скрипты, если ваши потребности выходят за рамки импорта и преобразования данных.

Практически все, что можно сделать вручную с помощью Excel пользовательского интерфейса, можно сделать с помощью Office скриптов. Они отлично подходит для применения последовательного форматирования к книгам. Скрипты создают диаграммы, pivotTables, фигуры, изображения и другие визуализации таблиц. Скрипты также дают точный контроль над положениями, размерами, цветами и другими атрибутами этих визуализаций.

Включение кода TypeScript обеспечивает высокую степень настройки. Программная логика управления, как `if...else` и утверждения, делает сценарий надежным. Это позволяет делать такие вещи, как условно считывая данные, не полагаясь на сложные формулы Excel или сканировать книгу на непредвиденные изменения перед изменением книги.

Форматирование можно применять с помощью Power Query Excel [шаблонов](https://templates.office.com/power-query-tutorial-tm11414620). Однако шаблоны обновляются на уровне отдельных или организаций, в то время как Office скрипты предоставляют более подробное управление доступом.

## <a name="power-automate-integrations"></a>Power Automate интеграции

Office скрипты предлагают дополнительные возможности для Power Automate интеграции. Сценарии адаптированы к вашим решениям. Вы [определяете вход и выход сценария](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts), поэтому он работает с любым другим соединитетелем или данными в потоке. На следующем скриншоте показан пример Power Automate, который передает данные из Teams адаптивной карты в сценарий Office.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="Снимок экрана, на Excel сетевом (бизнес) соединители в конструкторе потока. Соединительные устройства используют действие run script для вводимой Teams адаптивной карты и предоставления ее сценарию.":::

Power Query используется в [соедините](https://powerquery.microsoft.com/flow/) SQL Server Power Automate. [Преобразование данных с Power Query](/connectors/sql/#transform-data-using-power-query) позволяет создавать запрос в Power Automate. Хотя это мощный инструмент для использования с SQL Server, он ограничивает доступ Power Query к этому источнику ввода, как показано на следующем скриншоте потока.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="Снимок экрана, на SQL Server соединителета в конструкторе потока. Соединитетелем используется преобразование данных с Power Query действий.":::

## <a name="platform-dependencies"></a>Зависимости платформы

Office скрипты доступны только для Excel в Интернете. Power Query в настоящее время доступен только для Excel на рабочем столе. Оба можно использовать через Power Automate, что позволяет потоку работать с Excel книгами, хранимыми в OneDrive.

## <a name="see-also"></a>См. также

- [Power Query Портал](https://powerquery.microsoft.com/)
- [Power Query с Excel](https://powerquery.microsoft.com/excel/)
- [Запустите Office скрипты с Power Automate](../develop/power-automate-integration.md)
