---
title: Добавление изображений в книгу
description: Узнайте, как использовать Office скрипты для добавления изображения в трудовую книжку и копирования его на листах.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 99c3cc2cacf6e535bdb882bb8414d23fd105be35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546039"
---
# <a name="add-images-to-a-workbook"></a>Добавление изображений в книгу

Этот пример показывает, как работать с изображениями с помощью Office скрипта в Excel.

## <a name="scenario"></a>Сценарий

Изображения помогают с брендингом, визуальной идентичностью и шаблонами. Они помогают сделать трудовую книжку больше, чем просто гигантский стол.

Первый образец копирует изображение с одного листа на другой. Это может быть использовано, чтобы поставить логотип вашей компании в том же положении на каждом листе.

Второй образец копирует изображение с URL. Это может быть использовано для копирования фотографий, которые коллега хранил в общей папке в родственную трудовую книжку.

## <a name="sample-excel-file"></a>Образец Excel файла

Скачать файл <a href="add-images.xlsx">add-images.xlsx</a> используется в этих образцах и попробовать его самостоятельно!

## <a name="sample-code-copy-an-image-across-worksheets"></a>Пример кода: Копирование изображения на листах

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>Пример кода: Добавить изображение из URL в трудовую книжку

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/images/git-octocat.png";
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
