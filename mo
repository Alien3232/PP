'************************************************************************************************* /
'Macro Created by Jujubi
'*************************************************************************************************

Sub Messenger_email()
iLastRow1 = Worksheets("EmpDetails").Range("A1").End(xlDown).Row
Set WshShell = CreateObject("WScript.Shell")
For I = 2 To iLastRow1
If Worksheets("EmpDetails").Cells(I, "X").Value = "N" Then
 Dim mail As Object
 Dim OA As Object
 Dim message1 As String
    
    Set OA = CreateObject("Outlook.Application")
    Set mail = OA.CreateItem(0)
    
    message1 = "Hello," & vbNewLine & _
            "This is an auto generated email" & vbNewLine & _
             "Message Body" & vbNewLine & _
                "- <<Sender Name >>"
    
    With mail
    a = Worksheets("EmpDetails").Cells(I, "P").Value
    
        .To = a
        .CC = ""
        .BCC = ""
        .Subject = "Messenger: <<Email Subject>>"
        .Body = message1
     .Display
    End With
    WshShell.SendKeys "%{s}"
    On Error GoTo 0
    Set mail = Nothing
    Set OA = Nothing
    Else
  End If
     Next
End Sub
