Dim curfolder As String
Dim curfolpath As String
Dim pathfound As Boolean
Dim oldval As String
Public outlooksent As Boolean
Public deep As Integer
Public curdeep As Integer

Sub test(r As Range)
Dim path As String
Dim pathA, pathB As String

If Not Intersect(r, ActiveSheet.Range("A4:N39")) Is Nothing Then
If Intersect(r, ActiveSheet.Range("E34:N39")) Is Nothing Then
    If r.Height < 18 Then
    oldval = Trim(r.Cells(1, 1).Value)
    If Not Trim(r.Cells(1, 1).Value) = "" Then
    
        pathA = "C:\Users\" & Environ$("username") & ThisWorkbook.Worksheets("Settings").Range("B1").Value & "\" & Trim(r.Cells(1, 1).Value) & "\"
        pathB = "C:\Users\" & Environ$("username") & ThisWorkbook.Worksheets("Settings").Range("B2").Value & "\" & Trim(r.Cells(1, 1).Value) & "*.*"
        
        If r.Cells(1, 1).Comment Is Nothing Then

        
        pathfound = False
        curfolder = Trim(r.Cells(1, 1).Value)
        curfolpath = ""
        

        If Not Dir(pathB, vbDirectory) = "" Then
            pathfound = True
            curfolpath = "C:\Users\" & Environ$("username") & ThisWorkbook.Worksheets("Settings").Range("B2").Value & "\" & Dir(pathB, vbDirectory)
        End If
        

        If pathfound = False Then
        curdeep = 1
        deep = ThisWorkbook.Worksheets("Settings").Range("B6").Value
        Call findFolder
        End If
        
        If pathfound = False Then
        path = "C:\Users\" & Environ$("username") & ThisWorkbook.Worksheets("Settings").Range("B5").Value & "\" & Trim(r.Cells(1, 1).Value) & "\"
        MkDir path
        
        r.Cells(1, 1).AddComment ThisWorkbook.Worksheets("Settings").Range("B5").Value & "\" & Trim(r.Cells(1, 1).Value) & "\"
        r.Cells(1, 1).Comment.Visible = False
        
        Shell "cmd /C start """" /max """ & path & """", vbHide
        Else
        
        curfolpath = Mid(curfolpath, InStr(1, curfolpath, Trim(ThisWorkbook.Worksheets("Settings").Range("B7").Value)))
        r.Cells(1, 1).AddComment curfolpath
        r.Cells(1, 1).Comment.Visible = False
        Shell "cmd /C start """" /max """ & "C:\Users\" & Environ$("username") & curfolpath & """", vbHide
        End If
        
        Else
        
            If Not Dir("C:\Users\" & Environ$("username") & r.Cells(1, 1).Comment.Text, vbDirectory) = "" Then
                Shell "cmd /C start """" /max """ & "C:\Users\" & Environ$("username") & r.Cells(1, 1).Comment.Text & """", vbHide
            Else
            
                pathfound = False
                curfolder = Trim(r.Cells(1, 1).Value)
                curfolpath = ""

                If Not Dir(pathB, vbDirectory) = "" Then
                    pathfound = True
                    curfolpath = "C:\Users\" & Environ$("username") & ThisWorkbook.Worksheets("Settings").Range("B2").Value & "\" & Dir(pathB, vbDirectory)
                End If
                    

                
                If pathfound = False Then
                    curdeep = 1
                    deep = ThisWorkbook.Worksheets("Settings").Range("B6").Value
                    Call findFolder
                End If
                    
                    
                If pathfound = False Then
                path = "C:\Users\" & Environ$("username") & ThisWorkbook.Worksheets("Settings").Range("B5").Value & "\" & Trim(r.Cells(1, 1).Value) & "\"
                MkDir path
                
                r.Cells(1, 1).Comment.Text ThisWorkbook.Worksheets("Settings").Range("B5").Value & "\" & Trim(r.Cells(1, 1).Value) & "\"
                
                Shell "cmd /C start """" /max """ & path & """", vbHide
                Else
                
                curfolpath = Mid(curfolpath, InStr(1, curfolpath, Trim(ThisWorkbook.Worksheets("Settings").Range("B7").Value)))
                r.Cells(1, 1).Comment.Text curfolpath
                Shell "cmd /C start """" /max """ & "C:\Users\" & Environ$("username") & r.Cells(1, 1).Comment.Text & """", vbHide
                End If
          
            End If
        End If
    
    End If
    End If
End If
End If
End Sub

Sub findFolder()

    Dim searchFolderName As String
    searchFolderName = "C:\Users\" & Environ$("username") & ThisWorkbook.Worksheets("Settings").Range("B1").Value & "\"

    Dim FileSystem As Object

    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    

    doFolder FileSystem.getFolder(searchFolderName)
    
    

End Sub

Sub doFolder(Folder)

    Dim subFolder
    On Error Resume Next
    For Each subFolder In Folder.subfolders

        If InStr(1, Split(subFolder, "\")(UBound(Split(subFolder, "\"))), curfolder) > 0 Then
            curfolpath = subFolder
            
            pathfound = True
        Exit Sub
        End If
        
        If curdeep < deep Then
        curdeep = curdeep + 1
        doFolder subFolder
        End If
        
    Next subFolder
    curdeep = curdeep - 1
End Sub


Sub createevent(title As String, s As Date)
Dim olapp As Outlook.Application
    Set olapp = Outlook.Application

Dim olNS As Outlook.Namespace
  Dim objOwner As Outlook.Recipient
  Dim newCalFolder As Outlook.Folder
  Dim olAppt As Outlook.AppointmentItem
  Set olNS = olapp.GetNamespace("MAPI")
  Set objOwner = olNS.CreateRecipient(ThisWorkbook.Worksheets("Settings").Range("B4").Value)
    objOwner.Resolve

 If objOwner.Resolved Then

 If DuplicateEvent(s, title) = True Then
 outlooksent = True
 Exit Sub
 End If
 
 Set newCalFolder = olNS.GetSharedDefaultFolder(objOwner, olFolderCalendar)


 Set olAppt = newCalFolder.Items.Add(olAppointmentItem)
     With olAppt
 
        .AllDayEvent = True
        .start = s
        .End = DateAdd("d", 1, s)
        .Subject = title

        .RequiredAttendees = ThisWorkbook.Worksheets("Settings").Range("B3").Value
        '.OptionalAttendees =
        .MeetingStatus = olMeeting
        .Categories = "Projects"
        .Body = title
        
        .ReminderSet = True
        .ReminderMinutesBeforeStart = "900"
            

        .Send
    End With
    outlooksent = True
    Else
    outlooksent = False
 End If
 End Sub
 
Sub ClearProj(r As Range, s As Date)

If Not Intersect(r, ActiveSheet.Range("A4:N39")) Is Nothing Then
If Intersect(r, ActiveSheet.Range("E34:N39")) Is Nothing Then
    If r.Height < 18 Then
    If Trim(r.Cells(1, 1).Value) = "" Then
        If Not r.Cells(1, 1).Comment Is Nothing Then
            r.Cells(1, 1).Comment.Delete
        End If
        DeleteEvent s
    End If
    End If
End If
End If
End Sub


Sub DeleteEvent(start As Date)
Dim olapp As Outlook.Application
    Set olapp = Outlook.Application

Dim olNS As Outlook.Namespace
  Dim objOwner As Outlook.Recipient

  Set olNS = olapp.GetNamespace("MAPI")
  Set objOwner = olNS.CreateRecipient(ThisWorkbook.Worksheets("Settings").Range("B4").Value)
    objOwner.Resolve

 If objOwner.Resolved Then

 Set newCalFolder = olNS.GetSharedDefaultFolder(objOwner, olFolderCalendar)


 For Each olAppt In newCalFolder.Items
     With olAppt

     If .start = start And .Categories = "Projects" Then
     olAppt.Delete
     MsgBox "Deleted the event from calender"
     End If
       
    End With
 Next
 Else
 MsgBox "Error While Delete the event from calender"
 End If
 End Sub
 
Function DuplicateEvent(start As Date, title As String) As Boolean
    Dim olapp As Outlook.Application
    Set olapp = Outlook.Application

    Dim olNS As Outlook.Namespace
  Dim objOwner As Outlook.Recipient

  Set olNS = olapp.GetNamespace("MAPI")
  Set objOwner = olNS.CreateRecipient(ThisWorkbook.Worksheets("Settings").Range("B4").Value)
    objOwner.Resolve

 If objOwner.Resolved Then

 Set newCalFolder = olNS.GetSharedDefaultFolder(objOwner, olFolderCalendar)


 For Each olAppt In newCalFolder.Items
     With olAppt

     If .start = start And .Categories = "Projects" And .Subject = title Then
     DuplicateEvent = True
     Exit Function
     End If
       
    End With
 Next
 Else
 DuplicateEvent = False
 End If
 End Function
Sub Event_trigger(target As Range)
Dim theDate As Date
'Dim monthnumber As Integer
'Dim daynumber As Integer
'monthnumber = Month(DateValue("1 " & ActiveSheet.Name & " 2020"))
'monthnumber = Format(monthnumber, "00")


'Put this code from here to - New Update 27/3/2022
'<<<
If Not Intersect(target, ActiveSheet.Range("A4:N39")) Is Nothing Then
If Intersect(target, ActiveSheet.Range("E34:N39")) Is Nothing Then
Else
Exit Sub
End If
Else
Exit Sub
End If
'>>> here

Dim l As Integer
l = target.Row
Select Case l
Case 1 To 9
theDate = ActiveSheet.Cells(4, target.Column).Value
Case 10 To 15
theDate = ActiveSheet.Cells(10, target.Column).Value
Case 16 To 21
theDate = ActiveSheet.Cells(16, target.Column).Value
Case 22 To 27
theDate = ActiveSheet.Cells(22, target.Column).Value
Case 28 To 33
theDate = ActiveSheet.Cells(28, target.Column).Value
Case 34 To 39
theDate = ActiveSheet.Cells(34, target.Column).Value
Case Else
End Select

'theDate = DateSerial(Year(Now), monthnumber, daynumber)
Dim tmp As String
tmp = ActiveSheet.Cells(target.Row, target.Column).Value
If tmp = "" Then
ClearProj target, theDate
Else
createevent target.Value, theDate
If outlooksent = True Then
MsgBox "Added to the Calender"
Else
MsgBox "Error While adding to the Calender"
End If
End If
End Sub

Sub Initial_ADD()
Dim theDate As Date
Dim j, k As Integer
Dim l As Integer
Dim target As Range
'Dim monthnumber As Integer
'Dim daynumber As Integer
'monthnumber = Month(DateValue("1 " & ActiveSheet.Name & " 2020"))
'monthnumber = Format(monthnumber, "00")
k = 4


For j = 1 To 7

Do While k < 40

Set target = ActiveSheet.Cells(k, 1 + 2 * (j - 1))


If k < 10 And ActiveSheet.Cells(4, target.Column).Value = "" Then
GoTo skipday
End If

If k > 27 And ActiveSheet.Cells(28, target.Column).Value = "" Then
GoTo skipday
End If

If k > 33 And ActiveSheet.Cells(34, target.Column).Value = "" Then
GoTo skipday
End If

If Not Intersect(target, ActiveSheet.Range("E34:N39")) Is Nothing Then
GoTo skipday
End If


l = target.Row
Select Case l
Case 1 To 9
theDate = ActiveSheet.Cells(4, target.Column).Value
Case 10 To 15
theDate = ActiveSheet.Cells(10, target.Column).Value
Case 16 To 21
theDate = ActiveSheet.Cells(16, target.Column).Value
Case 22 To 27
theDate = ActiveSheet.Cells(22, target.Column).Value
Case 28 To 33
theDate = ActiveSheet.Cells(28, target.Column).Value
Case 34 To 39
theDate = ActiveSheet.Cells(34, target.Column).Value
Case Else
End Select

If theDate < Date Then
GoTo skipday
End If

'theDate = DateSerial(Year(Now), monthnumber, daynumber)
If Not target.Value = "" Then
outlooksent = False
createevent target.Value, theDate
custom_wait 0.02
End If


skipday:
k = k + 1
Set target = ActiveSheet.Cells(k, 1 + 2 * (j - 1))

If Not target.Height < 18 Then
k = k + 1
End If

Loop
k = 4
Next
End Sub
