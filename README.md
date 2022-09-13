# API EXCEL

Метод, позвляющий создать таблицу в MS Excel и записать туда данные из json


***/excel*** доступ к api excel

___POST___

_/excel_ - Добавляет данные в excel

*Parameters*
json - данные вида:

```
{
      "data" : [{"col1": [1, 2, 3]}, {"col2": ["q", "w", "e"]}, {"col3": 1}]
}
```

Responses 201 успешно
Данные по умолчанию записываются в таблицу _random.xlsx_

