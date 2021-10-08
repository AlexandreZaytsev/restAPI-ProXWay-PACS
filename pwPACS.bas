Attribute VB_Name = "pwPACS"
' Physical Access Control System (������� ���������� ��������� �������)
' ��� ������� ���� ��� (��������� �� ��������� ������ ���������)
' ������������ ��������� � �������� JSON JsonConverter.vb (' tools/references = Mivrosoft Scripting Runtime)
' ��� ������ �� ��������� REST API ProxWay
' http://localhost:40001/json/help

Option Explicit

Const srvHost = "localhost" ' "rpk-342"  ' "localhost" ' "rpk-342" '"localhost" '"rpk-342" ' "localhost" '"rpk-342"      '������
Const srvLogint = "web_admin"               '�����
Const srvPassword = "ric"                   '������

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'������� ����������� �� �������
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' ������� - ������ UserSID ��� ���� �������� - ������ ������
Private Function connectRestApi() 'string
  Dim pass, pointHostName, req, ret, json
   ret = ""
'   pass = UCase(GetHash(UCase(GetHash(UCase(GetHash(srvPassword) + "F593B01C562548C6B7A31B30884BDE53")))))
   pass = pwHash.MD5_string(pwHash.MD5_string(pwHash.MD5_string(srvPassword) + "F593B01C562548C6B7A31B30884BDE53"))

 '-------------�����������
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
'�������� ���������� id ������������ ���� ProxWay
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' idRic - id ������� ����� ��� ��������� ����� ������������ � ProxWay
' userName - ��� ����������� (����������� ������������� ����� ����� ������ %) ������ ����� � ������� SQL ��� ������� Like
' idExcel - id ������������ � ��������� ���� ��� - ���� �������� �������
' ������� - id (Token) ������������ ������� ProxWay - ��� � ������ ������� - ������ ������
Private Function GetUserIdProxWayByName(idRic, userName, idExcel) 'infoArr
  Dim info, req, ret, json, UserSID, pointHostName, usersInfo, userInfo, i
  Dim pwUserID, count, msg
  
   msg = ""
   pwUserID = ""
   UserSID = connectRestApi()
   If Len(UserSID) > 0 Then
 
 '-------------������ �������������
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
         msg = "���������� �� ����������"
       Case Else
         If count > 1 Then
           msg = "���������� ����� ������ ������������" & vbCrLf & "����� ����������� ���������"
         End If
         
         Dim item As Object                ' Reference to each JSON Object found in the "data" property.
         Dim rows As VBA.Collection        ' Reference to each JSON Object found in the "data" property's JSON Array.
         Dim row  As Long                  ' Number of rows in the "data" property's JSON Array.
         Dim data As Scripting.Dictionary  ' Reference to a JSON Object in the "data" property's JSON Array.
     
         i = 0
         For Each userInfo In json("Employee")
           pwUserID = userInfo("Token")                         'id �����
              
           If Len(idRic) > 0 Then                               '���� ����� �������� �� id ���������� �� ����� (��������� ����� � ProxWay)
             If idRic <> CStr(userInfo("EmployeeNumber")) Then
               pwUserID = ""
               msg = "���������� �� ���������� ������ ����������" & vbCrLf & "�� ����������"
               Exit For
             End If
           End If
       
           If Len(idExcel) > 0 Then                             '���� ����� �������� �� idExcel ���������� �� Excel ���� ����� �������� ������� (��� ���� � ProxWay)
             Set rows = userInfo("AdditionalFields")            '�������� ���������
             If rows.count > 0 Then
               For row = 1 To rows.count Step 1                   '������ �� ���������
                 Set item = rows.item(row)                        '�������� ������ �� ��������� (JSON Array)
                 Set data = item                                  '������������� ��� � �������
                 If data.Items(0) = "id ���� ����� �������� ������� (Excel)" Then '��������� ���������� ����
                   If idExcel <> CStr(data.Items(1)) Then
                     pwUserID = ""
                     msg = "���������� �� ����������"
                     Exit For
                   End If
                 End If
               Next row
             Else
               pwUserID = ""
               msg = "���������� �� ����������" & vbCrLf & "� ���� ProxWay �� ��������� �������������� ���� ����������"
               Exit For
             End If
           End If
           i = i + 1
         Next userInfo
     End Select
     If Len(msg) > 0 Then
       MsgBox "������ ������������� ������������" & vbCrLf & vbCrLf & _
              " �� ���������� �������" & vbCrLf & _
              "   ��� ����������        : '" & userName & "'" & vbCrLf & _
              "   id ���������� �� �����: '" & idRic & "'" & vbCrLf & _
              "   id ���������� �� Excel: '" & idExcel & "'" & vbCrLf & _
              vbCrLf & msg, vbOKCancel + vbInformation, "������������ ���� ProwWay"
     End If
     
    '-------------�����
     pointHostName = "Logout"
     req = "{""UserSID"":""" & UserSID & """}"
     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0)
   
     Set json = Nothing

   End If
   GetUserIdProxWayByName = pwUserID
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'������ � ������� �� ���� ����� ������ ����������� ���������� �/�� �����, � ��������� ������
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' pwIdUser - id ������������ ProxWay
' findDate - ���� ������� � ������� "����.mm.dd" (����� ����� ���������� �� ��������� ���� � ��������� ������� �� 00:00:00 �� 23:59:59)
' ������� - ���������� ������ - ������ �������� - ����� ������� ����� (���� ����), ������ �������� - ����� ���������� ������ (���� ����)
Function CheckPointPWTime(pwIdUser, findDateTime)
  Dim res(1), pass, i, pointHostName, pwDataTime
  Dim url, req, ret, json, UserSID
  Dim usersInfo As Variant
  Dim userInfo As Dictionary
    
   res(0) = ""                      '����� ������� �����
   res(1) = ""                      '����� ���������� ������

   UserSID = connectRestApi()
   If Len(UserSID) > 0 Then

 '-------------�� ���� ��� ��� �����
 '    pointHostName = "AdditionalEventFieldGetList"
 '    req = "{" & _
 '          """Language"":""ru"", " & _
 '          """UserSID"":""" & UserSID & """, " & _
 '          """SubscriptionEnabled"":true, " & _
 '          """Limit"":0, " & _
 '          """StartToken"":0, " & _
 '          "}"
 '    ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 1)

 '-------------��������� ������ �������
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
'           """Employees"":[], "  '������ id ����������� [1785, 1809] ����� � ������
'     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0) '0-�� �������� ����
     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 1) '1- �������� ����
   
     Set json = pwJsonConverter.ParseJSON(ret)
'     MsgBox json("Event").count
   
     ReDim usersInfo(json("Event").count, 6)
     i = 0
     For Each userInfo In json("Event")
       If Len(CStr(userInfo("CardCode"))) > 0 Then                      '���� ������� ������� �� �����
         ' =userInfo("CardCode")                                        '��� �����
'         Select Case CInt(userInfo("Sender")("Token"))
'           Case 5356                                                    ':5356,"Name":"���� ��� - ����"
'           Case 5357                                                    ':5357,"Name":"���� ��� - �����"
'         End Select
         
         Select Case CStr(userInfo("Message")("Name"))                  '����� ������� � �������
           Case "���� ��������"                                         '��� "���� ��������"
             pwDataTime = CDate(pwUTC.parseJSONdate(userInfo("Issued"), utcOffset))
             If res(0) = "" Then
               res(0) = pwDataTime
             ElseIf pwDataTime < CDate(res(0)) Then
               res(0) = pwDataTime
             End If
           Case "����� ��������"                                        '��� "����� ��������"
             pwDataTime = CDate(pwUTC.parseJSONdate(userInfo("Issued"), utcOffset))
             If res(1) = "" Then
               res(1) = pwDataTime
             ElseIf pwDataTime > CDate(res(1)) Then
               res(1) = pwDataTime
             End If
                       
'           usersInfo(i, 0) = userInfo("Token")                      'id �������
'           usersInfo(i, 1) = utc.parseJSONdate(userInfo("Issued"), utcOffset)      '����� �������
'           usersInfo(i, 2) = userInfo("User")("Token")              'id �����
'           usersInfo(i, 3) = userInfo("User")("EmployeeNumber")     '��������� ����� �����
'           usersInfo(i, 4) = userInfo("User")("Name")               '��� �����
'           usersInfo(i, 5) = userInfo("Message")("Token")           'id �������
'           usersInfo(i, 5) = userInfo("Message")("Name")            '������������ �������
         End Select
       End If
       i = i + 1
     Next userInfo

 '-------------�����
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
   
 '-------------������ �������������
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
       usersInfo(i, 0) = userInfo("Token")                        'id �����
       usersInfo(i, 1) = userInfo("Name")                         '��� �����
       usersInfo(i, 2) = userInfo("CardCount")                    '���������� ��������� �����
       usersInfo(i, 3) = userInfo("DepartmentToken")              'id ������ �����
       usersInfo(i, 4) = userInfo("DepartmentName")               '��� ������ �����
       usersInfo(i, 5) = userInfo("Post")                         '��������� �����
       usersInfo(i, 6) = userInfo("Email")                        'Email �����
       usersInfo(i, 7) = userInfo("EmployeeNumber")               '��������� ����� �����
       usersInfo(i, 8) = pwUTC.parseJSONdate(userInfo("LastModified"), utcOffset)  '���� ��������� ������
       i = i + 1
     Next userInfo
   
    '-------------�����
     pointHostName = "Logout"
     req = "{""UserSID"":""" & UserSID & """}"
     ret = GetRestData("http://" & srvHost & ":40001/json/", pointHostName, req, 0)
   
     Set json = Nothing
'������� �� ����
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
'�������� ���������� id ������������ ���� ProxWay
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' idRic - id ������� ����� ��� ��������� ����� ������������ � ProxWay
' userName - ��� ����������� (����������� ������������� ����� ����� ������ %) ������ ����� � ������� SQL ��� ������� Like
' idExcel - id ������������ � ��������� ���� ��� - ���� �������� �������
' passTime - ������������� ���� �������
' ������� - ������ ��� ����/�����
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
