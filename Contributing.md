# <a name="contribute-to-this-documentation"></a>Улучшение этой документации

Благодарим вас за интерес к этой документации!

* [Как внести свой вклад](#ways-to-contribute)
* [Внесение изменений с помощью GitHub](#contribute-using-github)
* [Внесение изменений с помощью Git](#contribute-using-git)
* [Как форматировать статью с помощью Markdown](#how-to-use-markdown-to-format-your-topic)
* [Вопросы и ответы](#faq)
* [Дополнительные ресурсы](#more-resources)

## <a name="ways-to-contribute"></a>Как внести свой вклад

Вот несколько способов, которые можно внести в эту документацию:

* Чтобы внести небольшие изменения в статью, [Contribute использует GitHub](#contribute-using-github).
* Для внесения больших изменений или изменений, затрагивающих код, [Contribute использует Git](#contribute-using-git).
* Сообщите об ошибках документации, используя проблемы GitHub.
* Запросите новую документацию в [Excel для сайта веб-UserVoice](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439) .

## <a name="contribute-using-github"></a>Внесение изменений с помощью GitHub

Используйте GitHub для участия в этой документации без необходимости клонировать репозиторий на Рабочий стол. Это самый простой способ создания запроса на включение внесенных изменений в этом репозитории. Используйте этот метод для внесения незначительных изменений, которые не затрагивают изменения кода.

**Примечание**. при использовании этого метода вы можете вносить изменения в одну статью за раз.

### <a name="to-contribute-using-github"></a>Для участия с использованием GitHub

1. Найдите нужную статью на сайте GitHub.
2. Когда вы входите в статью в GitHub, входите в систему GitHub (получите бесплатную учетную запись " [присоединиться к GitHub](https://github.com/join)").
3. Выберите **значок карандаша** (отредактируйте файл в разделе этого проекта) и внесите изменения в окно **<>редактирование файла** .
4. Прокрутите окно вниз и введите описание.
5. Выберите вариант **предложить изменение файла**>**создать запрос на включение внесенных**изменений.

Теперь вы успешно отправили запрос на включение внесенных изменений. Запросы на включение внесенных изменений обычно проверяются в течение 10 рабочих дней.

## <a name="contribute-using-git"></a>Внесение изменений с помощью Git

Используйте Git для внесения изменений существенные, таких как:

* Сопутствующий код.
* Вклад изменений, влияющих на значение.
* Внесение больших изменений в текст.
* Добавление новых разделов.

### <a name="to-contribute-using-git"></a>Чтобы внести изменения с помощью Git

1. Если у вас нет учетной записи GitHub, настройте ее на сайте [GitHub](https://github.com/join).
2. После создания учетной записи установите Git на своем компьютере. Выполните действия, описанные в руководстве по [настройке Git] .
3. Чтобы отправить запрос на включение внесенных изменений с помощью Git, выполните действия, описанные в разделе [Использование GitHub, Git и этого репозитория](#use-github-git-and-this-repository).
4. Вам будет предложено подписать лицензионное соглашение для участника, если вы:

    * Участник группы Microsoft Open Technologies.
    * Сотрудник, который не работает в корпорации Майкрософт.

В качестве члена сообщества необходимо подписать лицензионное соглашение (CLA), прежде чем вы сможете вносить большие отправки в проект. Вам нужно только выполнить и послать документацию только один раз. Внимательно просмотрите документ. Может потребоваться подпись вашего работодателя.

При подписывании CLA не предоставляются права на сохранение в основном репозитории, но это означает, что Teams для разработчиков Office и публикации контента для разработчиков Office сможет просматривать и утверждать ваши публикации. Вы кредитуется на отправку.

Запросы на включение внесенных изменений обычно проверяются в течение 10 рабочих дней.

## <a name="use-github-git-and-this-repository"></a>Использование GitHub, Git и этого репозитория

**Примечание**: большинство сведений в этом разделе можно найти в статьях [справки GitHub] .  Если вы знакомы с Git и GitHub, перейдите к разделу " **участие и редактирование контента** ", в котором описывается характер кода/содержимого этого репозитория.

### <a name="to-set-up-your-fork-of-the-repository"></a>Настройка разветвления репозитория

1. Чтобы добавлять информацию в этот проект, настройте учетную запись GitHub. Если это еще не сделано, перейдите на страницу [GitHub](https://github.com/join) и сделайте это сейчас.
2. Установите Git на своем компьютере. Выполните действия, описанные в руководстве по [настройке Git] .
3. Создайте свое ответвление этого репозитория. Для этого нажмите кнопку **ветвления** в верхней части страницы.
4. Скопируйте разветвление на компьютер. Для этого откройте Bash Git. Введите в командной строке следующую команду:

        git clone https://github.com/<your user name>/<repo name>.git

    Затем создайте ссылку на корневой репозиторий с помощью следующих команд:

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Поздравляем! Вы настроили свой репозиторий. Вам не потребуется повторять эти действия.

### <a name="contribute-and-edit-content"></a>Добавление и редактирование содержимого

Чтобы сделать процесс обработки как можно более плавным, выполните указанные ниже действия.

#### <a name="to-contribute-and-edit-content"></a>Добавление и изменение контента

1. Создайте новую ветвь.
2. Добавьте или отредактируйте содержимое.
3. Отправьте запрос на включение внесенных изменений в основной репозиторий.
4. Удалите ветвь.

**Важно!** Ограничьте каждую ветвь одной статьей или статьей для упрощения рабочего процесса и снижения вероятности конфликтов слияния. Содержимое, подходящее для новой ветви, включает в себя:

* Новая статья.
* Редактирования правописания и грамматики.
* Применение одного изменения форматирования в большом наборе статей (например, применение нового нижнего колонтитула для авторских прав).

#### <a name="to-create-a-new-branch"></a>Создание новой ветви

1. Откройте Bash Git.
2. В командной строки Git Bash введите `git pull upstream master:<new branch name>`. При этом создается локальная ветвь, которая копируется из последней главной ветви OfficeDev.
3. В командной строки Git Bash введите `git push origin <new branch name>`. Это оповещает GitHub о новой ветви. Новая ветвь появится в вашем ответвлении репозитория на сайте GitHub.
4. В командной строки Git Bash введите `git checkout <new branch name>` команду, чтобы переключиться на новую ветвь.

#### <a name="add-new-content-or-edit-existing-content"></a>Добавление или редактирование содержимого

Вы можете перейти к репозиторию на компьютере с помощью проводника. Файлы репозитория находятся в `C:\Users\<yourusername>\<repo name>`.

Чтобы изменить файлы, откройте их в выбранном Вами редакторе и измените их. Чтобы создать новый файл, используйте редактор выбора и сохраните новый файл в соответствующем расположении в локальной копии репозитория. Во время работы сохраните работу часто.

Файлы в `C:\Users\<yourusername>\<repo name>` — это рабочая копия новой ветви, созданной в локальном репозитории. Изменение файлов в этой папке не влияет на локальный репозиторий, пока вы не сохраните изменение. Чтобы сохранить изменение в локальном репозитории, введите следующие команды в GitBash:

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

Команда `add` добавляет изменения в область промежуточного хранения для подготовки к сохранению в репозитории. Период после `add` команды указывает, что необходимо разместить все добавленные или измененные файлы, а также рекурсивно проверять вложенные папки. (Если вы не хотите сохранять все изменения, вы можете добавить определенные файлы. Вы также можете отменить изменения. Для получения справки введите команду `git add -help` или `git status`.)

Команда `commit` применяет промежуточные изменения в репозитории. Переключатель `-m` означает, что вы предоставляете комментарий Commit в командной строке. Параметры – v и – a можно опустить. Параметр-v предназначен для подробного вывода команды, а параметр-a — для того, что вы уже выполнили команду Add.

Вы можете зафиксировать их несколько раз во время работы или вы можете выполнить их один раз.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Отправка запроса на включение внесенных изменений в основной репозиторий

Завершив работу и приготовьтесь к объединению с главным репозиторием, выполните указанные ниже действия.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Чтобы отправить запрос на включение внесенных изменений в основной репозиторий

1. В командной строки Git Bash введите `git push origin <new branch name>`. `origin` — это репозиторий GitHub, с которого вы скопировали свой локальный репозиторий. Эта команда передает текущее состояние новой ветви, в том числе все изменения, сохраненные на предыдущих шагах, в вашем ответвлении GitHub.
2. На сайте GitHub перейдите к новой ветви в своем ответвлении.
3. Нажмите кнопку **запрос на включение внесенных изменений** в верхней части страницы.
4. Убедитесь, что базовая ветвь `OfficeDev/<repo name>@master` — и ветвь head `<your username>/<repo name>@<branch name>`.
5. Нажмите кнопку **обновить диапазон фиксации** .
6. Добавьте название в запрос на включение внесенных изменений и опишите все изменения, которые вы вносите.
7. Отправьте запрос на включение внесенных изменений.

Один из администраторов сайта будет обрабатывать запрос на включение внесенных изменений. Запрос на включение внесенных изменений будет указан на<repo name> OfficeDevе или на сайте, на котором возникли проблемы. Если запрос на включение внесенных изменений принят, проблема будет устранена.

#### <a name="create-a-new-branch-after-merge"></a>Создание ветви после объединения

После успешного объединения ветви (то есть запрос на включение внесенных изменений принимается) не продолжайте работать в этой локальной ветви. Это может привести к конфликтам слияния, если вы отправите другой запрос на включение внесенных изменений. Чтобы выполнить другое обновление, создайте новую локальную ветвь из успешной Объединенной ветви, а затем удалите начальную локальную ветвь.

Например, если локальная ветвь X успешно объединена в главную ветвь OfficeDev/Office-Script-Master, и вы хотите внести дополнительные обновления объединяемого контента. Создайте новую локальную ветвь x2 из главной ветви OfficeDev/Office-Scripts-Master. Для этого откройте GitBash и выполните следующие команды:

    cd office-scripts-docs
    git pull upstream master:X2
    git push origin X2

Теперь у вас есть локальные копии (в новой локальной ветви) работы, которые вы отправили в ветке X. Кроме того, ветвь x2 содержит все работающие другие авторы, которые могут быть объединены, поэтому если работа зависит от работы других пользователей (например, общие изображения), она доступна в новой ветви. Вы можете убедиться в том, что предыдущая рабочая и совместная работы находятся в филиале, выполнив проверку новой ветви...

    git checkout X2

…и проверив ее содержимое. ( `checkout` Команда обновляет файлы в `C:\Users\<yourusername>\office-scripts-docs` текущем состоянии ветви x2.) После извлечения новой ветви вы можете вносить изменения в контент и зафиксировать их как обычно. Однако во избежание работы в объединенной ветви (X) рекомендуем удалить ее (см. раздел **Удаление ветви** ниже).

#### <a name="delete-a-branch"></a>Удаление ветви

После успешного объединения изменений в основной репозиторий удалите используемую ветвь, так как она больше не нужна.  Любую дополнительную работу следует выполнить в новой ветви.  

#### <a name="to-delete-a-branch"></a>Удаление ветви

1. В командной строки Git Bash введите `git checkout master`. Это гарантирует, что вы не работаете в удаляемой ветви (это не допускается).
2. Затем введите `git branch -d <branch name>`в командной строке команду. При этом ветвь будет удалена только в том случае, если она была успешно объединена с вышестоящим репозиторием. (Вы можете переопределить это правило с помощью флага `–D`, но для начала убедитесь, что это необходимо.)
3. Наконец, введите в командной строке команду `git push origin :<branch name>` (с пробелом перед двоеточием и без пробела после него).  При этом ветвь будет удалена из вашего ответвления GitHub.  

Поздравляем, вы успешно участвовали в проекте.

## <a name="how-to-use-markdown-to-format-your-topic"></a>Как форматировать статью с помощью Markdown

### <a name="markdown"></a>Markdown

Во всех статьях в этом репозитории используется Markdown. Полное введение (и пример синтаксиса) можно найти на сайте [ДАРИНГ фиребалл-Markdown].

## <a name="faq"></a>Вопросы и ответы

### <a name="how-do-i-get-a-github-account"></a>Как создать учетную запись GitHub?

Заполните форму на странице [Join GitHub](https://github.com/join), чтобы создать бесплатную учетную запись GitHub.

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Где можно найти лицензионное соглашение с участником?

Вам будет автоматически отправлено уведомление, если для включения внесенных вами изменений вам нужно подписать лицензионное соглашение с участником (CLA).

**Вам нужно подписать лицензионное соглашение с участником (CLA), прежде чем вы сможете вносить большие изменения в этот проект**. Вам нужно заполнить и отправить документ всего один раз. Внимательно просмотрите документ. Может потребоваться подпись вашего работодателя.

### <a name="what-happens-with-my-contributions"></a>Что происходит с моими публикациями?

Когда вы отправляете свои изменения с помощью запроса на включение внесенных изменений, наша команда будет уведомлена и будет проверять запрос на включение внесенных изменений. Вы получите уведомления о запросе на включение внесенных изменений от GitHub; Вы также можете уведомить кого-либо от нашей команды, если нам нужна дополнительная информация. Если запрос на включение внесенных изменений утвержден, мы будем обновлять документацию. Мы зарезервированием права на редактирование отправок для юридических, стилей, ясности или других проблем.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Можно ли стать утверждающим для запросов на получение запроса GitHub этого репозитория?

В настоящее время у внешних участников не разрешается утверждать запросы на включение внесенных изменений в этом репозитории.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>Как скоро будет получен ответ на запрос на изменение?

Запросы на включение внесенных изменений обычно проверяются в течение 10 рабочих дней.

## <a name="more-resources"></a>Дополнительные ресурсы

* Чтобы узнать больше о Markdown, перейдите на сайт Markdown Creator [ДАРИНГ фиребалл].
* Чтобы узнать больше об использовании Git и GitHub, сначала ознакомьтесь со статьей [GitHub Help].

[GitHub Home]: http://github.com
[Справка GitHub]: http://help.github.com/
[Настройка Git]: https://help.github.com/articles/set-up-git/
[ДАРИНГ Фиребалл — Markdown]: http://daringfireball.net/projects/markdown/
[ДАРИНГ Фиребалл]: http://daringfireball.net/