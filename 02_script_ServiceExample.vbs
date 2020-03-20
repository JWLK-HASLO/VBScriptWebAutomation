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
     .getElementByid("ltbInputOperator_Username").Value = "userid_test20"
     .getElementByid("ltbInputOperator_PasswordDevice").Value = "userpw_test20"
     .getElementByid("ltbInputOperator_Barcode").Value = "device_barcode20"
End With
waitIE(objIE)


Sub dropDownButton()
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
                    WScript.Sleep 100
               'Second Section of Home department :: Department
               ElseIf Not StrComp(data, "lddInputHomeDepartments") Then               
                    objButtonNodeList.item(i).Click
                    WScript.Sleep 100
                    WshShell.SendKeys "office"
                    WScript.Sleep 100
                    WshShell.SendKeys "{ENTER}"
                    WScript.Sleep 100
               'Third Section of Home department :: Role
               ElseIf Not StrComp(data, "lddInputAccessModel_SelectedRole") Then
                    objButtonNodeList.item(i).Click
                    WScript.Sleep 100
                    WshShell.SendKeys "{ENTER}"
                    WScript.Sleep 100
               End If
          End If
     Next
End Sub

Call dropDownButton()
WScript.Sleep 100


'Save Button Click
objIE.Document.getElementByid("btnSave").Click
waitIE(objIE)


Sub dropDownButtonOneMore()
     'Find All Button & Save to objButtonNodeList (Type object) 
     'One More => For 'Access Right' Section Hospital DropBox
     Dim objButtonNodeListOneMore : Set objButtonNodeListOneMore = objIE.Document.getElementsByTagName("Button")
     Dim numButtonNodeListOneMore
     numButtonNodeListOneMore = objButtonNodeListOneMore.length

     'Looping Using Statement 'For' 
     ' & Find Attribute 'data-id' For "Access rights" Section Dropdown Menu Button
     ' & Find Only Hospital Button  'data-id' Attribute is "lddInputHospitals"
     For j = 0 to numButtonNodeListOneMore - 1
          dataOneMore = objButtonNodeListOneMore.item(j).getAttribute("data-id") 
          If Not IsNull(dataOneMore) Then
               'Section of Home department &  Access rights :: Hospital
               If Not StrComp(dataOneMore, "lddInputHospital") Then
                    objButtonNodeListOneMore.item(j).Click
                    WScript.Sleep 100
                    WshShell.SendKeys "{TAB}"
                    WScript.Sleep 100
                    WshShell.SendKeys "{ENTER}"
                    WScript.Sleep 100
               End If
          End If
     Next
End Sub

Call dropDownButtonOneMore()
WScript.Sleep 100


'Add Button Click
objIE.Document.getElementByid("btnAdd").Click
waitIE(objIE)

'Devices Script Command Add => Later


'Save Button Click
objIE.Document.getElementByid("btnSave").Click
waitIE(objIE)

'Top Operator Manu Click
objIE.Document.getElementByid("menuitem-operator").Click
waitIE(objIE)

'All Operators Check Button Click
objIE.Document.getElementByid("allOperatorsChk").Click
waitIE(objIE)

'Push Operators Button Click
objIE.Document.getElementByid("btnPushAllOperators").Click
waitIE(objIE)

'Push Operators OK Button Click
objIE.Document.getElementByid("btnOk").Click
waitIE(objIE)