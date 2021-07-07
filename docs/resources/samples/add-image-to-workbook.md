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
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="371b3-103">Добавление изображений в книгу</span><span class="sxs-lookup"><span data-stu-id="371b3-103">Add images to a workbook</span></span>

<span data-ttu-id="371b3-104">В этом примере показано, как работать с изображениями с помощью Office скрипта в Excel.</span><span class="sxs-lookup"><span data-stu-id="371b3-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="371b3-105">Сценарий</span><span class="sxs-lookup"><span data-stu-id="371b3-105">Scenario</span></span>

<span data-ttu-id="371b3-106">Изображения помогают с брендингом, визуальной идентичностью и шаблонами.</span><span class="sxs-lookup"><span data-stu-id="371b3-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="371b3-107">Они помогают сделать книгу больше, чем просто гигантская таблица.</span><span class="sxs-lookup"><span data-stu-id="371b3-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="371b3-108">Первый пример копирует изображение из одного таблицы в другой.</span><span class="sxs-lookup"><span data-stu-id="371b3-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="371b3-109">Это можно использовать для того, чтобы поместить логотип вашей компании в одинаковое положение на каждом листе.</span><span class="sxs-lookup"><span data-stu-id="371b3-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="371b3-110">Второй пример копирует изображение из URL-адреса.</span><span class="sxs-lookup"><span data-stu-id="371b3-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="371b3-111">Это можно использовать для копирования фотографий, хранимых коллегой в общей папке, в связанную книгу.</span><span class="sxs-lookup"><span data-stu-id="371b3-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="371b3-112">Пример Excel файла</span><span class="sxs-lookup"><span data-stu-id="371b3-112">Sample Excel file</span></span>

<span data-ttu-id="371b3-113">Скачайте <a href="add-images.xlsx">add-images.xlsx</a> для готовой к использованию книги.</span><span class="sxs-lookup"><span data-stu-id="371b3-113">Download <a href="add-images.xlsx">add-images.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="371b3-114">Добавьте следующие скрипты и попробуйте пример самостоятельно!</span><span class="sxs-lookup"><span data-stu-id="371b3-114">Add the following scripts and try the sample yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="371b3-115">Пример кода. Скопируйте изображение в таблицах</span><span class="sxs-lookup"><span data-stu-id="371b3-115">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="371b3-116">Пример кода. Добавление изображения из URL-адреса в книгу</span><span class="sxs-lookup"><span data-stu-id="371b3-116">Sample code: Add an image from a URL to a workbook</span></span>

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
