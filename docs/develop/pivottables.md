---
title: Работа с сводными таблицами в Office скриптах
description: Сведения об объектной модели для сводных таблиц в API JavaScript Office скриптов.
ms.date: 04/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: 579f94140214674912c9610e707123924e4aef18
ms.sourcegitcommit: 4e3d3aa25fe4e604b806fbe72310b7a84ee72624
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/27/2022
ms.locfileid: "65077092"
---
# <a name="work-with-pivottables-in-office-scripts"></a>Работа с сводными таблицами в Office скриптах

Сводные таблицы позволяют быстро анализировать большие коллекции данных. Благодаря их возможности усложняется. API Office скриптов позволяют настраивать сводную таблицу в соответствии со своими потребностями, но область набора API усложняет начало работы. В этой статье показано, как выполнять общие задачи сводной таблицы, а также описываются важные классы и методы.

> [!NOTE]
> Чтобы лучше понять контекст терминов, используемых интерфейсами API, сначала ознакомьтесь Excel сводной таблицы. Начните [с создания сводной таблицы для анализа данных листа](https://support.microsoft.com/office/a9a84538-bfe9-40a9-a8e9-f99134456576).

## <a name="object-model"></a>Объектная модель

:::image type="content" source="../images/pivottable-object-model.png" alt-text="Упрощенное изображение классов, методов и свойств, используемых при работе со сводными таблицами.":::

[Сводная таблица](/javascript/api/office-scripts/excelscript/excelscript.pivottable) — это центральный объект для сводных таблиц в API Office скриптов.

- Объект [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) содержит коллекцию всех сводных [таблиц](/javascript/api/office-scripts/excelscript/excelscript.pivottable). [Каждый лист также](/javascript/api/office-scripts/excelscript/excelscript.worksheet) содержит коллекцию сводных таблиц, которая является локальной для этого листа.
- [Сводная таблица](/javascript/api/office-scripts/excelscript/excelscript.pivottable) содержит [pivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy). Иерархию можно рассматривать как столбец в таблице.
- [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) можно добавить в виде строк или столбцов ([RowColumnPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.rowcolumnpivothierarchy)), данных ([DataPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.datapivothierarchy)) или фильтров ([FilterPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)).
- [Каждый элемент PivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) содержит ровно одно [pivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield). Структуры сводных таблиц за пределами Excel могут содержать несколько полей для каждой иерархии, поэтому такая структура существует для поддержки будущих параметров. Для Office, поля и иерархии сопоставляются с одной и той же информацией.
- [PivotField содержит](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) несколько [элементов pivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem). Каждый элемент PivotItem является уникальным значением в поле. Каждый элемент можно рассматривать как значение в столбце таблицы. Элементы также могут быть агрегируемыми значениями, например суммами, если поле используется для данных.
- [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) определяет, как [отображаются pivotFields](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) и [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem).
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters) фильтрует [данные из](/javascript/api/office-scripts/excelscript/excelscript.pivottable) сводной таблицы, используя разные критерии.

Посмотрите, как эти связи работают на практике. В следующих данных описываются продажи деревьев из разных ферм. Это основа для всех примеров в этой статье. Используйте <a href="pivottable-sample.xlsx">pivottable-sample.xlsx</a> для выполнения.

:::image type="content" source="../images/pivottable-raw-data.png" alt-text="Коллекция продаж деревьев различных типов из разных ферм.":::

## <a name="create-a-pivottable-with-fields"></a>Создание сводной таблицы с полями

Сводные таблицы создаются со ссылками на существующие данные. Как диапазоны, так и таблицы могут быть источником сводной таблицы. Им также требуется место в книге. Так как размер сводной таблицы является динамическим, указывается только верхний левый угол диапазона назначения.

Следующий фрагмент кода создает сводную таблицу на основе диапазона данных. Сводная таблица не имеет иерархий, поэтому данные еще не сгруппированы.

```typescript
  const dataSheet = workbook.getWorksheet("Data");
  const pivotSheet = workbook.getWorksheet("Pivot");

  const farmPivot = pivotSheet.addPivotTable(
    "Farm Pivot", /* The name of the PivotTable. */
    dataSheet.getUsedRange(), /* The source data range. */
    pivotSheet.getRange("A1") /* The location to put the new PivotTable. */);
```

:::image type="content" source="../images/pivottable-empty.png" alt-text="Сводная таблица с именем Farm Pivot без иерархий.":::

### <a name="hierarchies-and-fields"></a>Иерархии и поля

Сводные таблицы организованы по иерархиям. Эти иерархии используются для сведений данных при добавлении в виде определенного типа иерархии. Существует четыре типа иерархий.

- **Строка**: отображает элементы в горизонтальных строках.
- **Столбец**: отображает элементы в вертикальных столбцах.
- **Данные**: отображает агрегаты значений на основе строк и столбцов.
- **Фильтр**. Добавление или удаление элементов из сводной таблицы.

Для сводной таблицы может быть назначено как можно больше или меньше полей для этих определенных иерархий. Сводной таблице требуется по крайней мере одна иерархия данных для отображения сводные числовые данные и по крайней мере одна строка или столбец для сводки по этой сводке. В следующем фрагменте кода добавляются две иерархии строк и две иерархии данных.

```typescript
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Farm"));
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Type"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold at Farm"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold Wholesale"));
```

:::image type="content" source="../images/pivottable-data-hierarchy.png" alt-text="Сводная таблица, показывающая общий объем продаж различных деревьев в зависимости от фермы, из которой они поступили.":::

## <a name="layout-ranges"></a>Диапазоны макетов

Каждая часть сводной таблицы сопоставляется с диапазоном. Это позволяет скрипту получать данные из сводной таблицы для последующего использования в скрипте или для возврата в [Power Automate потоке](power-automate-integration.md). Доступ к этим диапазонам осуществляется через объект [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) , полученный из `PivotTable.getLayout()`. На следующей схеме показаны диапазоны, возвращаемые методами в `PivotLayout`.

:::image type="content" source="../images/pivottable-layout-breakdown.png" alt-text="Схема, показывающая, какие разделы сводной таблицы возвращаются функциями получения диапазона макета.":::

## <a name="filters-and-slicers"></a>Фильтры и срезы

Существует три способа фильтрации сводной таблицы.

- [FilterPivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters)
- [Slicers](/javascript/api/office-scripts/excelscript/excelscript.slicer)

### <a name="filterpivothierarchies"></a>FilterPivotHierarchies

`FilterPivotHierarchies` добавьте дополнительную иерархию для фильтрации каждой строки данных. Любая строка с отфильтрованным элементом исключается из сводной таблицы и ее сводных данных. Так как эти фильтры основаны на элементах, они работают только с дискретными значениями. Если "Классификация" является иерархией фильтров в нашем примере, пользователи могут выбрать значения "Органическая" и "Обычная" для фильтра. Аналогичным образом, если выбран параметр "Ящики, проданные по оптово", то вместо числовых диапазонов будут использоваться отдельные числа, например 120 и 150.

`FilterPivotHierarchies` создаются со всеми выбранными значениями. Это означает, что ничего `PivotManualFilter` `FilterPivotHierarchy`не фильтруется, пока пользователь не вручную не будет взаимодействовать с элементом управления фильтром или не будет установлен в поле, принадлежащее .

В следующем фрагменте кода в качестве иерархии фильтров добавляется "Классификация".

```typescript
  farmPivot.addFilterHierarchy(farmPivot.getHierarchy("Classification"));
```

:::image type="content" source="../images/pivottable-filter-hierarchy.png" alt-text="Элемент управления фильтром, использующий &quot;Классификация&quot; для сводной таблицы.":::

### <a name="pivotfilters"></a>PivotFilters

Объект `PivotFilters` представляет собой коллекцию фильтров, примененных к одному полю. Так как каждая иерархия имеет ровно одно поле, `PivotHierarchy.getFields()` при применении фильтров всегда следует использовать первое поле. Существует четыре типа фильтров.

- **Фильтр дат**: фильтрация по дате календаря.
- **Фильтр меток**: фильтрация сравнения текста.
- **Фильтр вручную**: настраиваемая фильтрация входных данных.
- **Фильтр значений**: фильтрация сравнения номеров. При этом элементы в связанной иерархии сравниваются со значениями в указанной иерархии данных.

Как правило, к полю создается и применяется только один из четырех типов фильтров. Если скрипт пытается использовать несовместимые фильтры, возникает ошибка с текстом "Аргумент является недопустимым, отсутствует или имеет неправильный формат".

В следующем фрагменте кода добавляются два фильтра. Первый — это ручной фильтр, который выбирает элементы в существующей иерархии фильтров "Классификация". Второй фильтр удаляет все фермы, у которых меньше 300 "Crates Sold Sold". Обратите внимание, что это позволяет отфильтровать "Сумму" этих ферм, а не отдельные строки из исходных данных.

```typescript
  const classificationField = farmPivot.getFilterHierarchy("Classification").getFields()[0];
  classificationField.applyFilter({
    manualFilter: { 
      selectedItems: ["Organic"] /* The included items. */
    }
  });

  const farmField = farmPivot.getHierarchy("Farm").getFields()[0];
  farmField.applyFilter({
    valueFilter: {
      condition: ExcelScript.ValueFilterCondition.greaterThan, /* The relationship of the value to the comparator. */
      comparator: 300, /* The value to which items are compared. */
      value: "Sum of Crates Sold Wholesale" /* The name of the data hierarchy. Note the "Sum of" prefix. */
      }
  });
```

:::image type="content" source="../images/pivottable-filters.png" alt-text="Сводная таблица после применения фильтра значений и ручного фильтра.":::

### <a name="slicers"></a>Срезы

[Срезы](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) фильтруют данные в сводной таблице (или стандартной таблице). Это перемещаемые объекты на листе, которые позволяют быстро фильтровать выбранные элементы. Срез работает так же, как ручной фильтр и `PivotFilterHierarchy`. Элементы из сводной `PivotField` таблицы переключаются для включения или исключения из сводной таблицы.

В следующем фрагменте кода добавляется срез для поля "Тип". Для выбранных элементов задается значение "Гоголь" и "Пометка", а затем перемещает срез на 400 пикселей влево.

```typescript
  const fruitSlicer = pivotSheet.addSlicer(
    farmPivot, /* The table or PivotTale to be sliced. */
    farmPivot.getHierarchy("Type").getFields()[0] /* What source to use as the slicer options. */
  );
  fruitSlicer.selectItems(["Lemon", "Lime"]);
  fruitSlicer.setLeft(400);
```

:::image type="content" source="../images/slicer.png" alt-text="Срез, фильтруя данные в сводной таблице.":::

## <a name="see-also"></a>См. также

- [Основные сведения о сценариях Office в Excel для Интернета](scripting-fundamentals.md)
- [Справочник API для сценариев Office](/javascript/api/office-scripts/overview)
