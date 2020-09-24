Attribute VB_Name = "Dform"
' This Statement remove case sensitivity when comparing text "A"="a"
Option Compare Text

'*****************************************************************************
' This code was written for VB Planet. You may use this code freely. However
' displaying any of this code on a webpage or distributing it in
' uncompiled form is strictly prohibited. Thank you.
'          2000 VB Planet
'*****************************************************************************

'@===============================================================================
' CreateForm:
'   Creates a Dynamic Form
' Parameters:
'   Formtag = Will set the forms .tag property to identify this form.
'   Caption = Will set the forms .caption property which is the title of the form
'================================================================================
Sub CreateForm(ByVal FormTag As String, ByVal Caption As String)
 ' We create a new form using FORM1 as the template
 Dim Xfrm As New frmNote
 ' We set the .tag property of the form, this will help to identify it later
 Xfrm.Tag = FormTag
 ' Set the caption, this will appear in the title bar of the new form
 Xfrm.Caption = Caption
 ' Lastly we show the form (Make .Visible property True)
 Xfrm.Show
End Sub
'@===============================================================================
' RemoveForm:
'   Removes/unloads a form
' Parameters:
'   FormNumber = Specifies number of form in forms() collection.
' Note: If Formnumber is -1 the last form in the collection is removed
'===============================================================================
Public Sub RemoveForm(Optional ByVal FormNumber As Integer = -1)
' we check the value of FORMNUMBER Parameter
' if >-1 we remove that form number in the FORMS() collection
 If FormNumber > -1 Then
  Unload Forms(FormNumber)
 ' else we remove the last form in the collection
 Else
  ' The unload statement will remove the form from memory
  ' However form code still resides in memory
  Unload Forms(Forms.Count - 1)
 End If
End Sub


