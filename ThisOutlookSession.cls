Public Duty As Variant

Private WithEvents Items As Outlook.Items

Private Sub Application_Startup()
  Dim Ns As Outlook.NameSpace
  Set Ns = Application.GetNamespace("MAPI")
  Set objOwner = Ns.CreateRecipient("shared_mailbox_name")
  objOwner.Resolve
 Set Items = Ns.GetSharedDefaultFolder(objOwner, olFolderInbox).Items
 '===================================================
Set Connection = CreateObject("ADODB.Connection")
Dim recordSet As ADODB.recordSet, fld As ADODB.Field
Dim arr As Variant
Connection.Open "DSN=duty_excel"
Set records = New ADODB.recordSet
records.Open "select * from [" & MonthName(Month(Date)) & "$Print_Area]", Connection
i = 0
my = 0
    If Not records.EOF Then
        Do While Not records.EOF
            If InStr(records.Fields(2).Value, "FIO") > 0 Then
                my = i
                idx = 6
                Do While idx < 36
                    If InStr(records.Fields(idx).Value, "1") > 0 Then
                        If IsEmpty(Duty) Then
                            Duty = Array(idx - 5 & "." & Month(Date) & "." & Year(Date))  'Make the Variant an array with a single element
                        Else
                            ReDim Preserve Duty(UBound(Duty) + 1) 'Add next array element
                            Duty(UBound(Duty)) = idx - 5 & "." & Month(Date) & "." & Year(Date)
                            'Assign the array element
                        End If
                    End If
                    idx = idx + 1
                Loop
            End If
            'Next
            records.MoveNext 'next record
            i = i + 1
        Loop
    End If
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
  On Error Resume Next
  If TypeOf Item Is Outlook.MailItem Then
    For Each duty_date In Duty
    'sadsa
    If DateDiff("d", Date, duty_date) = 0 Then
        If Mail.UnRead Then
          Debug.Print Item.Subject
          Set i = Item.Forward
          ' i.Display
          i.Recipients.Add "forwadr@email"
          i.Send
        End If
    End If
  Next
  End If
End Sub

