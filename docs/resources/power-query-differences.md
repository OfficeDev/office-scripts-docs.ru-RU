---
title: При использовании Power Query или Office скриптов
description: Сценарии, наиболее подходящие для платформ Power Query и Office Scripts.
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1812b508b2cde4d304ecf228adfdd8f68de9808a
ms.sourcegitcommit: 383880e0dc0d09b8f76884675531e462a292d747
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/01/2021
ms.locfileid: "61245618"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>При использовании Power Query или Office скриптов

[Power Query](https://powerquery.microsoft.com) и Office scripts являются мощными решениями автоматизации для Excel. Оба решения Excel пользователям очищать и преобразовывать данные в книгах. Один power query или Office скрипт можно обновить и повторно использовать новые данные для получения согласованных результатов, что экономит время и позволяет быстрее работать с полученными сведениями.

В этой статье представлен общий обзор того, когда можно выступают за одну платформу над другой. В общем, Power Query хорош для вывода и преобразования данных из больших, внешних источников данных и Office Скрипты хорошо для быстрых, Excel-ориентированных решений и [Power Automate](../develop/power-automate-integration.md)интеграций .

## <a name="large-data-sources-and-data-retrieval-power-query"></a>Большие источники данных и ирисовка данных: запрос power

Мы рекомендуем Power Query при работе с источниками данных с поддерживаемых платформ.

Power Query имеет [встроенные подключения к](https://powerquery.microsoft.com/connectors/) данным для сотен источников. Power Query специально разработан для задач по сбору, преобразованию и комбинации данных. Когда вам нужны данные из одного из этих источников, Power Query предоставляет вам способ не кодить эти данные в Excel в нужной форме.

Эти подключения Power Query предназначены для больших наборов данных. Они не имеют таких же ограничений [передачи,](../testing/platform-limits.md) как Power Automate или Excel для Интернета.

Office Scripts предлагает легкое решение для небольших источников данных или источников данных, не охваченных соединитетелями Power Query. Это включает в [себя `fetch` использование API rest](../develop/external-calls.md) или получение сведений из различных источников данных, например Teams [адаптивной карты.](../resources/scenarios/task-reminders.md)

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>Форматирование, визуализация и программный контроль: Office скрипты

Мы рекомендуем Office скрипты, если ваши потребности выходят за рамки импорта и преобразования данных.

Практически все, что можно сделать вручную с помощью Excel пользовательского интерфейса, можно сделать с помощью Office скриптов. Они отлично подходит для применения последовательного форматирования к книгам. Скрипты создают диаграммы, pivotTables, фигуры, изображения и другие визуализации таблиц. Скрипты также дают точный контроль над положениями, размерами, цветами и другими атрибутами этих визуализаций.

Включение кода TypeScript обеспечивает высокую степень настройки. Программная логика управления, `if...else` как и утверждения, делает сценарий надежным. Это позволяет делать такие вещи, как условно считывающиеся данные, не полагаясь на сложные формулы Excel или сканировать книгу на непредвиденные изменения перед изменением книги.

Форматирование можно применять с помощью Power Query с помощью Excel [шаблонов.](https://templates.office.com/power-query-tutorial-tm11414620) Однако шаблоны обновляются на уровне отдельных или организаций, в то время как Office скрипты предоставляют более подробное управление доступом.

## <a name="power-automate-integrations"></a>Power Automate интеграции

Office скрипты предлагают дополнительные возможности для Power Automate интеграции. Сценарии адаптированы к вашим решениям. Вы [определяете вход и выход сценария,](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts)поэтому он работает с любым другим соединитетелем или данными в потоке. На следующем скриншоте показан пример Power Automate, который передает данные с Teams адаптивной карты в сценарий Office.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="Снимок экрана, на который Excel соединителет Online (Business) в конструкторе потока. Соединительные устройства используют действие run script для вводимого ввода из Teams адаптивной карты и предоставления его сценарию.":::

Power Query используется в [соедините SQL Server](https://powerquery.microsoft.com/flow/) Power Automate. Преобразование [данных с помощью действия Power Query](/connectors/sql/#transform-data-using-power-query) позволяет создавать запрос в Power Automate. Хотя это мощный инструмент для использования с SQL Server, он ограничивает power Query для этого источника ввода, как показано на следующем скриншоте потока.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="Снимок экрана, на SQL Server соединителета в конструкторе потока. Соединитетелем используется преобразование данных с помощью действия Power Query.":::

## <a name="platform-dependencies"></a>Зависимости платформы

Office скрипты доступны только для Excel в Интернете. Power Query в настоящее время доступен только для Excel на рабочем столе. Оба можно использовать через Power Automate, что позволяет потоку работать с Excel книгами, хранимыми в OneDrive.

## <a name="see-also"></a>См. также

- [Портал запроса питания](https://powerquery.microsoft.com/)
- [Power Query with Excel](https://powerquery.microsoft.com/excel/)
- [Запустите Office скрипты с Power Automate](../develop/power-automate-integration.md)
