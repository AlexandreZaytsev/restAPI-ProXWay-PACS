# restAPI-ProxWay-PACS
**Интеграция** оборудования **СКУД** (PACS) (Система Контроля Управления Доступом (Physical Access Control System))
**ProxWay** по **REST API** на **VBscript** (VBA Excel)

Производитель ProxWay https://proxway-ble.ru/   
- софт - proxway-ip 3.057.7055
- оборудование - контроллер PW-400v3 + считыватели PW-Mini Multi BLE
  
для парсинга JSON использвана библиотека VBA-JSON v2.3.1 Tim Hall - https://github.com/VBA-tools/VBA-JSON

транспорт Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
***
типовой алгоритм
> http:// host name :40001/json/Authenticate  
1. авторизация на web сервисе (login) 
> http:// host name :40001/json/EmployeeGetList  
2. получение id (token) пользователя по базовым полям   
-- ФИО (синтаксис запроса аналогичный trasact sql like '%) и/или    
-- табельный номер  
-- дополнительным пользовательским полям
> http:// host name :40001/json/EventGetList  
3. запрос событий на дату (фильтр из п.п. 2)
парсинг JSON через JsonConverter.bas  
*(парсинг очень медленный (на словарях и коллекциях) - обязательно задавайте критерии фильтрации в запросе)*

> http:// host name :40001/json/Logout
4. отключение от web сервиса (logout)

в примере есть - отчет для сотрудников зарегистрированных в системе и тестовая процедура получения первого и последнего прохода сотрудника на дату
