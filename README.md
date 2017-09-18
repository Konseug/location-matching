# location-matching

Предыстория

Хорошая знакомая К. работает аналитиком по продажам в очень крупной компании.
Эжедневно появляются десятки новых торговых точек по всей стране, и в обязанности К. входит понять,
в зону ответственности какого из ≈4000 торговых представителей входит каждая из точек и добавить ее к исполнителю в БД на рабочем сервере.
Это рутинная часть работы К., и появилась мысль ее автоматизировать.

О работе программы

3 файла (без интерфейса, в нем нет необходимости):
1) xlsx_read:
- Чтение из xlsx-файлов списка новых торговых точек и работающих супервайзоров;
- сохранение списков в pickle-файл.
2) geolocate:
- Чтение из pickle-файла списка списка новых торговых точек и списка супервайзоров;
- поиск по Яндекс-картам координат новых торговых точек;
- поиск ближайшего к каждой торговой точке супервайзора;
- сохранение списка из пар точка-супервизор в xlsx-файле для возможности просмотра и ручного редактирования.
3) fill_db:
- Чтение списка пар (торговая точка, супервайзор) из xlsx-файла;
- занесение этого списка соответсвий на удаленный сервер с помощью Selenium webdriver;
- составление файла отчета внесенных правок.

Реальный адрес рабочего сервера, логин и пароль сотрудника заменены на "******".
