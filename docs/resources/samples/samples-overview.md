---
title: Office Примеры сценариев
description: Доступные Office сценарии и сценарии.
ms.date: 05/25/2021
localization_priority: Normal
ms.openlocfilehash: 1b7e9cdd9e23f57d59e5e878a37b50afb63965fd
ms.sourcegitcommit: a063b3faf6c1b7c294bd6a73e46845b352f2a22d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/29/2021
ms.locfileid: "53202853"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Примеры сценариев и сценарии

В этом разделе Office решения автоматизации [сценариев,](../../overview/excel.md) которые помогают конечным пользователям выполнять повседневные задачи. Он содержит реалистичные сценарии, с которые сталкиваются бизнес-пользователи, и предоставляет подробные решения, а также пошаговую инструкцию по видеосвязи.

Для каждого из проектов в [Basics](#basics) и Beyond the [basics](#beyond-the-basics)ознакомьтесь с исходным кодом, пошаговую передачу [**видео на YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)и другие.

В [сценарии](#scenarios)мы включили несколько более крупных примеров сценариев, которые демонстрируют реальные примеры использования.

Мы также приветствуем [вклады сообщества](#community-contributions-and-fun-samples).

## <a name="basics"></a>Основы

| Project | Сведения |
|---------|---------|
| [Основы создания сценариев](../excel-samples.md) | В этих примерах демонстрируются основные строительные блоки для Office скриптов. |
| [Добавление комментариев в Excel](add-excel-comments.md) | Этот пример добавляет комментарии к ячейке, @mentioning коллеге. |
| [Добавление изображений в книгу](add-image-to-workbook.md) | Этот пример добавляет изображение в книгу и копирует изображение на листах.|
| [Скопируйте несколько Excel таблиц в одну таблицу](copy-tables-combine.md) | Этот пример объединяет данные из нескольких Excel таблиц в одну таблицу, которая включает все строки. |

## <a name="beyond-the-basics"></a>Более сложные действия

Ознакомьтесь со следующим конечным проектом, который автоматизирует примеры сценариев наряду с полными сценариями, Excel используемыми файлами и видео [(на YouTube).](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Сведения |
|---------|---------|
| [Перекрестные справочники](excel-cross-reference.md) | В этом примере Office скрипты и Power Automate для перекрестной ссылки и проверки сведений в различных книгах. |
| [Подсчет пустых строк в определенном листе или во всех листах](count-blank-rows.md) | В этом примере обнаруживается, есть ли пустые строки в листах, в которых ожидается, что данные будут присутствовать, а затем сообщается о том, что количество пустых строк Power Automate потоке. |
| [Диаграмма электронной почты и изображения таблиц](email-images-chart-table.md) | В этом примере Office сценарии и Power Automate для создания диаграммы и отправки этой диаграммы в качестве изображения по электронной почте. |
| [Внешние вызовы на извлечение](external-fetch-calls.md) | В этом примере используется для получения `fetch` GitHub для сценария. |
| [Фильтр Excel таблицы и получить видимый диапазон](filter-table-get-visible-range.md) | Этот пример фильтрует таблицу Excel и возвращает видимый диапазон в качестве объекта JSON. Этот JSON может быть предоставлен потоку Power Automate как часть более крупного решения. |
| [Управление режимом вычисления в Excel](excel-calculation.md) | В этом примере показано, как использовать режим вычисления и вычислять методы в Excel в Интернете с помощью Office скриптов. |
| [Перемещение строк по таблицам](move-rows-across-tables.md) | В этом примере показано, как перемещать строки по таблицам, экономя фильтры, а затем обрабатывая и повторно примыкая к фильтрам. |
| [Выходные Excel как JSON](get-table-data.md) | В этом решении показано, как Excel данные таблицы как JSON для использования в Power Automate. |
| [Удаление гиперссылки из каждой ячейки в Excel таблицы](remove-hyperlinks-from-cells.md) | Этот пример очищает все гиперссылки из текущего таблицы. |
| [Запуск сценария для всех файлов Excel в папке](automate-tasks-on-all-excel-files-in-folder.md) | Этот проект выполняет набор задач автоматизации для всех файлов, расположенных в папке OneDrive для бизнеса (также может использоваться для SharePoint папки). Он выполняет вычисления Excel файлов, добавляет форматирование и вставляет комментарий, @mentions коллегу. |
| [Запись большого набора данных](write-large-dataset.md) | В этом примере показано, как отправить большой диапазон в качестве более мелких субрангов. |

## <a name="scenarios"></a>Сценарии

Office Скрипты могут автоматизировать части вашей повседневной работы. Эти задачи часто существуют в уникальных экосистемах с Excel книгами, которые настроены определенными способами. Эти более крупные примеры сценариев демонстрируют такие реальные примеры использования. Они включают как Office, так и книги, чтобы вы могли видеть сценарий от конца до конца.

| Сценарий | Сведения |
|---------|---------|
| [Анализ загруженного из Интернета](../scenarios/analyze-web-downloads.md) | В этом сценарии имеется скрипт, который разберет записи веб-трафика для определения страны происхождения пользователя. В нем представлены навыки разбора текста, использования подфункции в скриптах, применения условного форматирования и работы со таблицами. |
| [Извлечение и построение графика данных об уровне воды от NOAA](../scenarios/noaa-data-fetch.md) | В этом сценарии Office сценарий для получения данных из внешнего источника (базы данных [NOAA Tides и Currents)](https://tidesandcurrents.noaa.gov/)и написания полученных сведений. Он выделяет навыки использования для `fetch` получения данных и использования диаграмм. |
| [Калькулятор оценок](../scenarios/grade-calculator.md) | В этом сценарии имеется скрипт, который проверяет запись инструктора для оценок класса. В нем представлены навыки проверки ошибок, форматирования ячейки и регулярных выражений. |
| [Планирование собеседований в Teams](../scenarios/schedule-interviews-in-teams.md) | В этом сценарии показано, как использовать Excel таблицу для управления временем собраний интервью и внести поток в расписание собраний в Teams. |
| [Напоминания о задачах](../scenarios/task-reminders.md) | В этом сценарии Office скрипт в потоке Power Automate для отправки напоминаний коллегам для обновления состояния проекта. В нем освещаются навыки интеграции Power Automate и передачи данных в скрипты и из них. |

## <a name="community-contributions-and-fun-samples"></a>Community и интересные примеры

Мы приветствуем [вклады](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) нашего Office скрипты сообщества! Не стесняйся создавать запрос на отзыв.

| Project | Сведения |
|---------|---------|
| [Игра жизни](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | Блог Yutao Huang в Excel Tech Community содержит сценарий для моделирования игры жизни Джона [*Конвея.*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life) |
| [Анимация поздравление сезонов](community-seasons-greetings.md) | Этот сценарий был предоставлен [Лесли Блэк](https://www.linkedin.com/in/lesblackconsultant/) в духе курортного сезона! Это забавный сценарий, который показывает поющие елки в Excel в Интернете с Office скриптами. |

## <a name="try-it-out"></a>Проверка

Эти примеры являются открытым исходным кодом. Попробуйте их самостоятельно. Вам понадобится учетная запись Майкрософт для работы или школы с лицензией на подписку Microsoft 365 (E3 или выше). Просто перенапишитесь, чтобы войти в свою учетную запись https://office.com и начать работу.

## <a name="leave-a-comment"></a>Оставьте комментарий

Не стесняйся оставлять комментарии, делать предложение или логить проблему с помощью раздела **Отзывов** в нижней части страницы документации определенного образца.
