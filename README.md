# csv-toxls

Преобразование файла .csv в .xlsx

```bash
csv-toxls [flags] filename
flags:
 -w : кодировка windows-1251 (по умолчанию - utf-8)
 -q : поля csv в кавычках (по умолчанию - без кавычек)
 -c : разделитель полей - запятая(,) (по умолчанию - точка с запятой(;) )
 -e : допустимы пустые строки (по умолчанию - пустые строки игнорируются)
 -n : numeric поля (по умолчанию - все поля строковые), в формате -n:1:3:6 (1, 3, 6 - номера полей)
 -h : показать заголовок с нумерацие полей
```
