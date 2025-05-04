# Excel to SQL Converter

Программа для конвертации Excel файлов в SQL скрипты создания базы данных.

## Требования
- .NET 6.0 или выше
- Visual Studio 2022 или выше

## NuGet пакеты
```
Install-Package ExcelDataReader
Install-Package ExcelDataReader.DataSet
Install-Package System.Text.Encoding.CodePages
```

## Структура проекта
```
ExcelProcessor/
├─── MainForm.cs         # Главная форма приложения
├─── ExcelFileProcessor.cs # Основной класс для обработки Excel файлов
├── Program.cs           # Точка входа в приложение
└── README.md           # Документация
```

## Функциональность
- Выбор Excel файлов
- Копирование файлов в рабочую директорию
- Конвертация в SQL скрипт
- Транслитерация русских имен в латиницу
- Автоопределение типов данных
- Обработка специальных типов (телефоны, даты, ИНН)
