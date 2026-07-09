# PassportContract

Локальное Windows-приложение для рабочего процесса:

**паспорт PDF/фото → OCR → ручная проверка → заполненный договор DOCX**

Интернет при обработке паспорта не используется. В программу не встроена отправка файлов, аналитика или облачное хранилище.

## Важно перед загрузкой на GitHub

Не открывай архив в 7-Zip/WinRAR и не перетаскивай файлы прямо из окна архива: вложенная папка `.github/workflows` может расплющиться, и GitHub не увидит сборку.

1. Нажми по ZIP правой кнопкой.
2. Выбери **Извлечь все**.
3. Открой обычную распакованную папку `passport-contract`.
4. Внутри должны быть:

```text
.github/
  workflows/
    build-windows.yml
app.py
requirements.txt
template_gph.docx
README.md
```

5. На GitHub нажми **Add file → Upload files**.
6. Перетащи из обычной распакованной папки все пять элементов, включая `.github`.
7. Нажми **Commit changes**.

После загрузки файл должен находиться именно по адресу:

```text
.github/workflows/build-windows.yml
```

## Сборка EXE

1. Открой вкладку **Actions**.
2. Слева выбери **Build Windows EXE**.
3. Нажми **Run workflow → Run workflow**.
4. Подожди завершения зелёной галочкой.
5. Открой завершённый запуск.
6. В разделе **Artifacts** скачай `PassportContract-windows`.

После распаковки артефакта запускай:

```text
PassportContract.exe
```

Не выноси один EXE из полученной папки: рядом лежит локальный OCR Tesseract.

## Использование

1. Выбери паспорт в PDF или как изображение.
2. Нажми **Распознать паспорт**.
3. Проверь и исправь поля.
4. При необходимости вручную добавь ИНН, СНИЛС, телефон и банковские реквизиты.
5. Нажми **Создать договор**.

В комплект уже добавлен подготовленный шаблон `template_gph.docx`.

## Плейсхолдеры для других шаблонов

```text
{{contract_number}}
{{contract_date}}
{{fio}}
{{passport_series_number}}
{{issued_by}}
{{issue_date}}
{{department_code}}
{{birth_date}}
{{registration_address}}
{{phone}}
{{inn}}
{{snils}}
{{bank_name}}
{{correspondent_account}}
{{settlement_account}}
{{bik}}
```

OCR неизбежно ошибается. Перед использованием готового договора нужно сверить реквизиты с оригиналом паспорта.
