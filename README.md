# Sapp
PHP обработчик Excel файлов, позволяет сопоставлять информацию в строках какого либо столбца и переносить нужные соседние ячейки в целевой(ые) файлы.

1. В корневой каталог помещаем таблицу (эталон), из которой берется столбик для поиска совпадений, и столбик, из которого будет заноситься информация в целевую таблицу(ы).
2. В папку "files" помещаем целевые(ой) файлы(л), в которые(ый) будем переносить информацию из эталона.
3. Открываем файл pack.php с помощью notepad++ или любого другого редактора кода и вносим необходимые изменения в порядковые номера колонок и меняем имя эталонного файла при необходимости.
4. Запускаем сервер и открываем браузер.
5. Если вы не меняли наименования каталогов и папок, тогда вводите в строку браузера адрес http://sapp/pack.php и ждете пару секунд до завершения процесса обработки. Продолжительность обработки может меняться в зависимости от размера файлов (например, обработка файла в 12 250 строк занимает всего 5 секунд).
