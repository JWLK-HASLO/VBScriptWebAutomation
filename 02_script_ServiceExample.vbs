' Writer : JWLK
' GitHub : https://github.com/JWLK-HASLO
' Email : jewels@haslo.co

loginId = "[ID-String]"
loginPw = "[PW-String]#"

Dim objIE : Set objIE = CreateObject("InternetExplorer.Application")
Dim WshShell : Set WshShell = WScript.CreateObject ("WScript.Shell")

Function waitIE(obj)
     Do While obj.Busy = True Or obj.ReadyState <> 4
          WScript.Sleep 100
     Loop
End Function

objIE.Navigate "[Site-URL]"
objIE.Visible = True
WshShell.AppActivate objIE 'Click IE Browser < This Need To Tab Control

With objIE.Document
  .getElementByid("ltbInputUsername").Value = loginId
  .getElementByid("ltbInputPassword").Value = loginPw
  .getElementByid("btnLogIn").Click
End With
waitIE(objIE)
'Top Operator Manu Click
objIE.Document.getElementByid("menuitem-operator").Click
waitIE(objIE)

'AddOperator Button Click
objIE.Document.getElementByid("btnAddOperator").Click
waitIE(objIE)

'Logon Information Input
With objIE.Document
     .getElementByid("ltbInputOperator_Username").Value = "userid_test"
     .getElementByid("ltbInputOperator_PasswordDevice").Value = "userpw_test"
     .getElementByid("ltbInputOperator_Barcode").Value = "device_barcode"
End With
waitIE(objIE)

'Find All Button & Save to objButtonNodeList (Type object)
Dim objButtonNodeList : Set objButtonNodeList = objIE.Document.getElementsByTagName("Button")
Dim numButtonNodeList
numButtonNodeList = objButtonNodeList.length

'Looping Using Statement 'For' 
' & Find Attribute 'data-id' Because "Home department" Section Dropdown Menu Button have Attribute 'data-id'
' & Find Not Null object  => Only Button having Attribute 'data-id'
For i = 0 to numButtonNodeList - 1
     data = objButtonNodeList.item(i).getAttribute("data-id") 
     If Not IsNull(data) Then
          'Compare to 'data-id' String & Using Action depends on Each Section
          'StrComp(A-String, B-String) is A-String = B-String Equel  
          ' => return '0', So Using 'Not' operator front of StrComp for Output Result '1'

          'First Section of Home department :: Hospital
          If Not StrComp(data, "lddInputHospitals") Then 
               objButtonNodeList.item(i).Click
               WScript.Sleep 100
               WshShell.SendKeys "{TAB}"
               WScript.Sleep 100
               WshShell.SendKeys "{ENTER}"
          'Second Section of Home department :: Department
          Else If Not StrComp(data, "lddInputHomeDepartments") Then
               objButtonNodeList.item(i).Click
               WScript.Sleep 100
               WshShell.SendKeys "Office"
               WScript.Sleep 100
               WshShell.SendKeys "{ENTER}"
          'Third Section of Home department :: Role
          Else If Not StrComp(data, "lddInputAccessModel_SelectedRole") Then
               objButtonNodeList.item(i).Click
               WScript.Sleep 100
               WshShell.SendKeys "{ENTER}"
          End If
     End If
Next
