---
title: Добавление изображений в книгу
description: Узнайте, как использовать Office скрипты, чтобы добавить изображение в книгу и скопировать его на листах.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 2ae77ca1295cf6e55443789506242d9cf888f9a1
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313906"
---
# <a name="add-images-to-a-workbook"></a>Добавление изображений в книгу

В этом примере показано, как работать с изображениями с помощью Office скрипта в Excel.

## <a name="scenario"></a>Сценарий

Изображения помогают с брендингом, визуальной идентичностью и шаблонами. Они помогают сделать книгу больше, чем просто гигантская таблица.

Первый пример копирует изображение из одного таблицы в другой. Это можно использовать для того, чтобы поместить логотип вашей компании в одинаковое положение на каждом листе.

Второй пример копирует изображение из URL-адреса. Это можно использовать для копирования фотографий, хранимых коллегой в общей папке, в связанную книгу.

## <a name="sample-excel-file"></a>Пример Excel файла

Скачайте <a href="add-images.xlsx">add-images.xlsx</a> для готовой к использованию книги. Добавьте следующие скрипты и попробуйте пример самостоятельно!

## <a name="sample-code-copy-an-image-across-worksheets"></a>Пример кода. Скопируйте изображение в таблицах

```TypeScript
/**
 * This script transfers an image from one worksheet to another.
 */
function main(workbook: ExcelScript.Workbook)
{
  // Get the worksheet with the image on it.
  let firstWorksheet = workbook.getWorksheet("FirstSheet");

  // Get the first image from the worksheet.
  // If a script added the image, you could add a name to make it easier to find.
  let image: ExcelScript.Image;
  firstWorksheet.getShapes().forEach((shape, index) => {
    if (shape.getType() === ExcelScript.ShapeType.image) {
      image = shape.getImage();
      return;
    }
  });

  // Copy the image to another worksheet.
  image.getShape().copyTo("SecondSheet");
}
```

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>Пример кода. Добавление изображения из URL-адреса в книгу

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://raw.githubusercontent.com/OfficeDev/office-scripts-docs/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image)
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) 
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
