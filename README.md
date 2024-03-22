Here's sample Inventory Tracker and Business Journal that help business keep track of client events and account for Inventory.
VBA was used to automate some steps and make the User Interface more fluid and simple. 
![image](https://github.com/j9ao29/ExcelDev/assets/55525806/737c18c2-4682-4788-a681-b46fd9218e48)

Here a syntax for logging the events/inventory
![image](https://github.com/j9ao29/ExcelDev/assets/55525806/43e1e5f4-0142-4f49-a435-35129cc35f6b)

Private Sub CommandButton2_Click()
Sheet1.Activate

Dim eve As Worksheet
Dim iRow As Long

Set Inv = ThisWorkbook.Sheets("Event")
iRow = [Counta(Event!A:A)] + 1

With Inv
.Cells(iRow, 1) = TextBox9.Value
.Cells(iRow, 2) = TextBox8.Value
.Cells(iRow, 3) = TextBox10.Value
.Cells(iRow, 3) = TextBox11.Value
.Cells(iRow, 3) = TextBox12.Value
End With

TextBox9.Value = vbNullString
TextBox8.Value = vbNullString
TextBox10.Value = vbNullString
TextBox11.Value = vbNullString
TextBox12.Value = vbNullString

End Sub

Private Sub CommandButton1_Click()
Sheet2.Activate

Dim Inv As Worksheet
Dim iRow As Long

Set Inv = ThisWorkbook.Sheets("Inventory")
iRow = [Counta(Inventory!A:A)] + 1

With Inv

.Cells(iRow, 1) = TextBox4.Value
.Cells(iRow, 2) = TextBox7.Value
.Cells(iRow, 3) = TextBox6.Value
End With

TextBox4.Value = vbNullString
TextBox7.Value = vbNullString
TextBox6.Value = vbNullString

End Sub

Private Sub CommandButton3_Click()

If CommandButton3.Caption = ">" Then
UserForm2.Width = 430
CommandButton3.Caption = "<"
Else
UserForm2.Width = 710
CommandButton3.Caption = ">"
End If


End Sub

Private Sub CommandButton4_Click()
Sheet3.Activate

Dim Journal As Worksheet
Dim icount As Long

Set Journal = ThisWorkbook.Sheets("Journal")
icount = [Counta(Journal!A:A)] + 1

With Journal

.Cells(icount, 1) = TextBox1.Value
.Cells(icount, 2) = TextBox2.Value
.Cells(icount, 3) = TextBox3.Value
End With

TextBox1.Value = vbNullString
TextBox2.Value = vbNullString
TextBox3.Value = vbNullString
End Sub

Private Sub Frame4_Click()

End Sub

Private Sub Image6_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Image6_Click()
Sheet1.Activate

Dim eve As Worksheet
Dim iRow As Long

Set Inv = ThisWorkbook.Sheets("Event")
iRow = [Counta(Event!A:A)] + 1

With Inv
.Cells(iRow, 1) = TextBox9.Value
.Cells(iRow, 2) = TextBox8.Value
.Cells(iRow, 3) = TextBox10.Value
.Cells(iRow, 3) = TextBox11.Value
.Cells(iRow, 3) = TextBox12.Value
End With

TextBox9.Value = vbNullString
TextBox8.Value = vbNullString
TextBox10.Value = vbNullString
TextBox11.Value = vbNullString
TextBox12.Value = vbNullString

End Sub


