Attribute VB_Name = "pwPACS"
' Physical Access Control System (Система Управления Контролем Доступа)
' Для Сервера СКУД РПК (развернут на локальной машине секретаря)
' Используется совместно с парсером JSON JsonConverter.vb (' tools/references = Mivrosoft Scripting Runtime)
' Вся работа по протоколу REST API ProxWay
' http://localhost:40001/json/help

Option Explicit

Const srvHost = "localhost" ' "rpk-342"  ' "localhost" ' "rpk-342" '"localhost" '"rpk-342" ' "localhost" '"rpk-342"      'сервер
Const srvLogint = "web_admin"               'логин
Const srvPassword = "ric"                   'пароль

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'попытка авторизации на сервере
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' возврат - строка UserSID или если неудачно - пустая строка
Private Function connectRestApi() 'string
  Dim pass, pointHostName, req, ret, json
   ret = ""
'   pass = UCase(GetHash(UCase(GetHash(UCase(GetHash(srvPassword) + "F593B01C562548C6B7A31B30884BDE53")))))
   pass = pwHash.MD5_string(pwHash.MD5_string(pwHash.MD5_string(srvPassword) + "F593B01C562548C6B7A31B30884BDE53"))

 '-------------авторизация
   pointHostName = "Authenticate"
   req = "{" & _
         """PasswordHash"":""" & pass & """, " & _
         """UserName"":""" & srvLogint & """" & _
         "}"
   ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0)
   If Len(ret) > 0 Then
     Set json = pwJsonConverter.ParseJSON(ret)
     ret = json("UserSID")
     Set json = Nothing
   End If
  connectRestApi = ret
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'получить внутренний id пользователя СКУД ProxWay
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' idRic - id хвостик Лоции или Табельный номер пользователя в ProxWay
' userName - ФИО пользоватея (допускается использование маски через символ %) скорее всего в формате SQL для функции Like
' idExcel - id пользователя в служебной базе РПК - учет рабочего времени
' возврат - id (Token) пользователя системы ProxWay - или в случае неудачи - пустая строка
Private Function GetUserIdProxWayByName(idRic, userName, idExcel) 'infoArr
  Dim info, req, ret, json, UserSID, pointHostName, usersInfo, userInfo, i
  Dim pwUserID, count, msg
  
   msg = ""
   pwUserID = ""
   UserSID = connectRestApi()
   If Len(UserSID) > 0 Then
 
 '-------------список пользователей
     pointHostName = "EmployeeGetList"
     req = "{" & _
           """Language"":""ru"", " & _
           """UserSID"":""" & UserSID & """, " & _
           """SubscriptionEnabled"":true, " & _
           """Limit"":0, " & _
           """StartToken"":0, " & _
           """AdditionalFieldsRequired"":true, " & _
           """Name"":""" & userName & """, " & _
           "}"

'           """DepartmentToken"":0, "
'           """DepartmentUsed"":true, " &
'           """HideDismissed"":true, " &
     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0)
  
     Set json = pwJsonConverter.ParseJSON(ret)
     count = json("Employee").count

     Select Case count
       Case 0
         msg = "совпадений не обнаружено"
       Case Else
         If count > 1 Then
           msg = "обнаружено более одного пользователя" & vbCrLf & "будет использован последний"
         End If
         
         Dim item As Object                ' Reference to each JSON Object found in the "data" property.
         Dim rows As VBA.Collection        ' Reference to each JSON Object found in the "data" property's JSON Array.
         Dim row  As Long                  ' Number of rows in the "data" property's JSON Array.
         Dim data As Scripting.Dictionary  ' Reference to a JSON Object in the "data" property's JSON Array.
     
         i = 0
         For Each userInfo In json("Employee")
           pwUserID = userInfo("Token")                         'id юзера
              
           If Len(idRic) > 0 Then                               'если нужна проверка на id сотрудника из Лоции (табельный номер в ProxWay)
             If idRic <> CStr(userInfo("EmployeeNumber")) Then
               pwUserID = ""
               msg = "совпадений по табельному номеру сотрудника" & vbCrLf & "не обнаружено"
               Exit For
             End If
           End If
       
           If Len(idExcel) > 0 Then                             'если нужна проверка на idExcel сотрудника из Excel базы учета рабочего времени (доп поля в ProxWay)
             Set rows = userInfo("AdditionalFields")            'получить коллекцию
             If rows.count > 0 Then
               For row = 1 To rows.count Step 1                   'пройти по коллекции
                 Set item = rows.item(row)                        'получить объект из коллекции (JSON Array)
                 Set data = item                                  'преобразовать его в словарь
                 If data.Items(0) = "id базы учета рабочего времени (Excel)" Then 'проверить конкретное поле
                   If idExcel <> CStr(data.Items(1)) Then
                     pwUserID = ""
                     msg = "совпадений не обнаружено"
                     Exit For
                   End If
                 End If
               Next row
             Else
               pwUserID = ""
               msg = "совпадений не обнаружено" & vbCrLf & "в базе ProxWay не настроены дополнительные поля сотрудника"
               Exit For
             End If
           End If
           i = i + 1
         Next userInfo
     End Select
     If Len(msg) > 0 Then
       MsgBox "Ошибка сопоставления пользователя" & vbCrLf & vbCrLf & _
              " по параметрам запроса" & vbCrLf & _
              "   ФИО сотрудника        : '" & userName & "'" & vbCrLf & _
              "   id сотрудника из Лоции: '" & idRic & "'" & vbCrLf & _
              "   id сотрудника из Excel: '" & idExcel & "'" & vbCrLf & _
              vbCrLf & msg, vbOKCancel + vbInformation, "Пользователи СКУД ProwWay"
     End If
     
    '-------------выход
     pointHostName = "Logout"
     req = "{""UserSID"":""" & UserSID & """}"
     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0)
   
     Set json = Nothing

   End If
   GetUserIdProxWayByName = pwUserID
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'запрос к серверу на счет входа выхода конкретного сотрудника в/из офиса, и получение ответа
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' pwIdUser - id пользователя ProxWay
' findDate - день запроса в формате "уууу.mm.dd" (поиск будет произведен на указанную дату в диапазоне времени от 00:00:00 до 23:59:59)
' возврат - одномерный массив - первое значение - время первого входа (если есть), второе значение - время последнего выхода (если есть)
Function CheckPointPWTime(pwIdUser, findDateTime)
  Dim res(1), pass, i, pointHostName, pwDataTime
  Dim url, req, ret, json, UserSID
  Dim usersInfo As Variant
  Dim userInfo As Dictionary
    
   res(0) = ""                      'время первого входа
   res(1) = ""                      'время последнего выхода

   UserSID = connectRestApi()
   If Len(UserSID) > 0 Then

 '-------------не знаю что это такое
 '    pointHostName = "AdditionalEventFieldGetList"
 '    req = "{" & _
 '          """Language"":""ru"", " & _
 '          """UserSID"":""" & UserSID & """, " & _
 '          """SubscriptionEnabled"":true, " & _
 '          """Limit"":0, " & _
 '          """StartToken"":0, " & _
 '          "}"
 '    ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 1)

 '-------------прочитать журнал событий
     pointHostName = "EventGetList"
'     url = "http://" & srvHost & ":40001/json/EventGetListV2"
     req = "{" & _
           """Language"":""ru"", " & _
           """UserSID"":""" & UserSID & """, " & _
           """SubscriptionEnabled"":false, " & _
           """Limit"":0, " & _
           """StartToken"":0, " & _
           """Employees"":[" & pwIdUser & "], " & _
           """IssuedFrom"":""" & "\/Date(" & CStr(pwUTC.ConvertToUnixTimeStamp(findDateTime & " 00:00:00", 3)) & ")\/" & """, " & _
           """IssuedTo"":""" & "\/Date(" & CStr(pwUTC.ConvertToUnixTimeStamp(findDateTime & " 23:59:59", 3)) & ")\/" & """, " & _
           "}"
'           """Employees"":[], "  'массив id сотрудников [1785, 1809] Ечина и Зайцев
'     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0) '0-не печатать логи
     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 1) '1- печатать логи
   
     Set json = pwJsonConverter.ParseJSON(ret)
'     MsgBox json("Event").count
   
     ReDim usersInfo(json("Event").count, 6)
     i = 0
     For Each userInfo In json("Event")
       If Len(CStr(userInfo("CardCode"))) > 0 Then                      'если событие прохода по ключу
         ' =userInfo("CardCode")                                        'код карты
'         Select Case CInt(userInfo("Sender")("Token"))
'           Case 5356                                                    ':5356,"Name":"Офис РПК - вход"
'           Case 5357                                                    ':5357,"Name":"Офис РПК - выход"
'         End Select
         
         Select Case CStr(userInfo("Message")("Name"))                  'время события в системе
           Case "Вход совершен"                                         'или "Вход разрешен"
             pwDataTime = CDate(pwUTC.parseJSONdate(userInfo("Issued"), utcOffset))
             If res(0) = "" Then
               res(0) = pwDataTime
             ElseIf pwDataTime < CDate(res(0)) Then
               res(0) = pwDataTime
             End If
           Case "Выход совершен"                                        'или "Выход разрешен"
             pwDataTime = CDate(pwUTC.parseJSONdate(userInfo("Issued"), utcOffset))
             If res(1) = "" Then
               res(1) = pwDataTime
             ElseIf pwDataTime > CDate(res(1)) Then
               res(1) = pwDataTime
             End If
                       
'           usersInfo(i, 0) = userInfo("Token")                      'id события
'           usersInfo(i, 1) = utc.parseJSONdate(userInfo("Issued"), utcOffset)      'время события
'           usersInfo(i, 2) = userInfo("User")("Token")              'id юзера
'           usersInfo(i, 3) = userInfo("User")("EmployeeNumber")     'табельный номер юзера
'           usersInfo(i, 4) = userInfo("User")("Name")               'имя юзера
'           usersInfo(i, 5) = userInfo("Message")("Token")           'id события
'           usersInfo(i, 5) = userInfo("Message")("Name")            'наименование события
         End Select
       End If
       i = i + 1
     Next userInfo

 '-------------выход
     pointHostName = "Logout"
     req = "{""UserSID"":""" & UserSID & """}"
     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0)
   End If
  CheckPointPWTime = res
End Function

Sub printUserList()
  Dim info, req, ret, json, UserSID, pointHostName, usersInfo, userInfo, i
  
   UserSID = connectRestApi()
   If Len(UserSID) > 0 Then
   
 '-------------список пользователей
     pointHostName = "EmployeeGetList"
     req = "{" & _
           """Language"":""ru"", " & _
           """UserSID"":""" & UserSID & """, " & _
           """SubscriptionEnabled"":true, " & _
           """Limit"":0, " & _
           """StartToken"":0, " & _
           """AdditionalFieldsRequired"":true, " & _
           "}"

'           """DepartmentToken"":0, "
'           """Name"":""String content"", "
'           """DepartmentUsed"":true, " &
'           """HideDismissed"":true, " &
     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0)
  
     Set json = pwJsonConverter.ParseJSON(ret)
 
     ReDim usersInfo(json("Employee").count, 8)
     i = 0
     For Each userInfo In json("Employee")
     'userInfo("AdditionalFields")("Name")
       usersInfo(i, 0) = userInfo("Token")                        'id юзера
       usersInfo(i, 1) = userInfo("Name")                         'имя юзера
       usersInfo(i, 2) = userInfo("CardCount")                    'количество пропусков юзера
       usersInfo(i, 3) = userInfo("DepartmentToken")              'id группы юзера
       usersInfo(i, 4) = userInfo("DepartmentName")               'имя группы юзера
       usersInfo(i, 5) = userInfo("Post")                         'должность юзера
       usersInfo(i, 6) = userInfo("Email")                        'Email юзера
       usersInfo(i, 7) = userInfo("EmployeeNumber")               'табельный номер юзера
       usersInfo(i, 8) = pwUTC.parseJSONdate(userInfo("LastModified"), utcOffset)  'дата изменения данных
       i = i + 1
     Next userInfo
   
    '-------------выход
     pointHostName = "Logout"
     req = "{""UserSID"":""" & UserSID & """}"
     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0)
   
     Set json = Nothing
'вывести на лист
     Dim row, col
     row = 5
     col = 2
     With Worksheets(1)
       For i = 0 To UBound(usersInfo, 1)
         .Cells(i + row, col + 0).Value = usersInfo(i, 0)
         .Cells(i + row, col + 1).Value = usersInfo(i, 1)
         .Cells(i + row, col + 2).Value = usersInfo(i, 2)
         .Cells(i + row, col + 3).Value = usersInfo(i, 3)
         .Cells(i + row, col + 4).Value = usersInfo(i, 4)
         .Cells(i + row, col + 5).Value = usersInfo(i, 5)
         .Cells(i + row, col + 6).Value = usersInfo(i, 6)
         .Cells(i + row, col + 7).Value = usersInfo(i, 7)
         .Cells(i + row, col + 8).Value = usersInfo(i, 8)
       Next
     End With
   End If
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'получить внутренний id пользователя СКУД ProxWay
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' idRic - id хвостик Лоции или Табельный номер пользователя в ProxWay
' userName - ФИО пользоватея (допускается использование маски через символ %) скорее всего в формате SQL для функции Like
' idExcel - id пользователя в служебной базе РПК - учет рабочего времени
' passTime - запрашиваемая дата прохода
' возврат - массив дат вход/выход
Function cheskPassProxWay(idRic, userName, idExcel, passTime) ' array
  Dim pwUserID, timeArr
  
  pwUserID = GetUserIdProxWayByName("", userName, idExcel)
  If Len(pwUserID) > 0 Then
    timeArr = CheckPointPWTime(pwUserID, passTime)
  Else
    ReDim timeArr(1)
    timeArr(0) = ""
    timeArr(1) = ""
  End If
 cheskPassProxWay = timeArr
End Function
