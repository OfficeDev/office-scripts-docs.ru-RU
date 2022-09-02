---
title: Добавление изображений в книгу
description: Узнайте, как с помощью сценариев Office добавить изображение в книгу и скопировать его между листами.
ms.date: 07/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 78c7779cf4d524ed62bf8d419135863228b23d33
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572607"
---
# <a name="add-images-to-a-workbook"></a>Добавление изображений в книгу

В этом примере показано, как работать с изображениями с помощью скрипта Office в Excel.

## <a name="scenario"></a>Сценарий

Изображения помогают использовать фирменную символику, визуальное удостоверение и шаблоны. Они помогают сделать книгу больше, чем просто таблицей.

В первом примере изображение копируется с одного листа на другой. Это можно использовать для того, чтобы поместить логотип вашей компании в одинаковое положение на каждом листе.

Во втором примере изображение копируется из URL-адреса. Это можно использовать для копирования фотографий, сохраненных коллегой в общей папке, в связанную книгу.

## <a name="sample-excel-file"></a>Пример файла Excel

[ Скачайтеadd-images.xlsx](add-images.xlsx) для готовой к использованию книги. Добавьте следующие скрипты и попробуйте пример самостоятельно!

## <a name="sample-code-copy-an-image-across-worksheets"></a>Пример кода: копирование изображения на листы

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>Пример кода: добавление изображения из URL-адреса в книгу

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
  workbook.getWorksheet("WebSheet").addImage(image);
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) as string[];
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
