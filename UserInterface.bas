Private identifier As String
Private line as Integer

Public Function createForm(title As String)

    identifier = "userform"
    line = 1

    Set Form = ThisWorkbook.VBProject.VBComponents.Add(3)
    Set Label = Form.Designer.Controls.Add("Forms.Label.1")

    With Form
        .properties("Name") = identifier
        .properties("Caption") = title
        .properties("Width") = 400
        .properties("Height") = 50
        .properties("BackColor") = RGB(245, 245, 245)
    End With

    With Label
        .top = 0
        .left = 0
        .width = 390
        .height = 1
        .backColor = RGB(220, 220, 220)
    End With
 
End Function

Public Function addDivider()

    Set Form = ThisWorkbook.VBProject.VBComponents(identifier)
    Set Label = Form.Designer.Controls.Add("Forms.Label.1")
    
    With Label
        .top = Form.properties("Height") - 25 + 10
        .left = 0
        .width = 390
        .height = 1
        .backColor = RGB(220, 220, 220)
    End With
    
    Form.properties("Height") = Form.properties("Height") + Label.height + 15 + 20
 
End Function

Public Function addLabel(content As String)

    Set Form = ThisWorkbook.VBProject.VBComponents(identifier)
    Set Label = Form.Designer.Controls.Add("Forms.Label.1")
    
    With Label
        .caption = content
        .top = Form.properties("Height") - 25
        .left = 20
        .width = 350
    End With
    
    Form.properties("Height") = Form.properties("Height") + Label.height + 5
 
End Function

Public Function addInput()

    Set Form = ThisWorkbook.VBProject.VBComponents(identifier)
    Set text = Form.Designer.Controls.Add("Forms.TextBox.1")
    
    With text
        .top = Form.properties("Height") - 25
        .left = 20
        .width = 350
        .height = 100
    End With
    
    Form.properties("Height") = Form.properties("Height") + text.height + 15

End Function

Public Function addDropdown(ParamArray items() As Variant)

    Set Form = ThisWorkbook.VBProject.VBComponents(identifier)
    Set Combo = Form.Designer.Controls.Add("Forms.ComboBox.1")
    
    With Combo
        .name = "combobox"
        .top = Form.properties("Height") - 25
        .left = 20
        .width = 350
    End With

    Form.codemodule.insertlines line, "Private Sub Combo_Initialize()"
    line = line + 1
    Dim item as variant
    For Each item in items
        Form.codemodule.insertlines line, "Me.Controls(""combobox"").AddItem """ & item & """"
        line = line + 1
    Next item
    Form.codemodule.insertlines line, "End Sub"
    line = line + 1
    
    Form.properties("Height") = Form.properties("Height") + Combo.height + 15

End Function

Public Function addList(ParamArray items() As Variant)

    Set Form = ThisWorkbook.VBProject.VBComponents(identifier)
    Set List = Form.Designer.Controls.Add("Forms.ListBox.1")
    
    With List
        .name = "listbox"
        .top = Form.properties("Height") - 25
        .left = 20
        .width = 350
    End With
    
    Form.codemodule.insertlines line, "Private Sub List_Initialize()"
    line = line + 1
    Dim item as variant
    For Each item in items
        Form.codemodule.insertlines line, "Me.Controls(""listbox"").AddItem """ & item & """"
        line = line + 1
    Next item
    Form.codemodule.insertlines line, "End Sub"
    line = line + 1

    Form.properties("Height") = Form.properties("Height") + List.height + 15

End Function

Public Function addToggle(content As String)

    Set Form = ThisWorkbook.VBProject.VBComponents(identifier)
    Set Label = Form.Designer.Controls.Add("Forms.Label.1")
    Set Check = Form.Designer.Controls.Add("Forms.CheckBox.1")

    With Label
        .caption = content
        .top = Form.properties("Height") - 25 + 3
        .left = 20
        .width = 300
    End With

    With Check
        .top = Form.properties("Height") - 25
        .left = 358
    End With
    
    Form.properties("Height") = Form.properties("Height") + Check.height + 5

End Function

Public Function addAttach(action As String)

    Set Form = ThisWorkbook.VBProject.VBComponents(identifier)
    Set Button = Form.Designer.Controls.Add("Forms.CommandButton.1")
    
    With Button
        .name = "button2"
        .caption = "Vedhæft" & Space(8) & ChrW(8593)
        .top = Form.properties("Height") - 25
        .left = 300
        .width = 70
    End With

    Form.codemodule.insertlines line, "Private Sub button2_Click()"
    line = line + 1
    Form.codemodule.insertlines line, action
    line = line + 1
    Form.codemodule.insertlines line, "End Sub"
    line = line + 1
    
    Form.properties("Height") = Form.properties("Height") + Button.height + 15
 
End Function

Public Function addActions(action As String)

    Set Form = ThisWorkbook.VBProject.VBComponents(identifier)
    Set SecondaryButton = Form.Designer.Controls.Add("Forms.CommandButton.1")
    Set PrimaryButton = Form.Designer.Controls.Add("Forms.CommandButton.1")
    
    With SecondaryButton
        .name = "button3"
        .caption = ChrW(8592) & Space(6) & "Annuller"
        .top = Form.properties("Height") - 25
        .left = 220
        .width = 70
    End With

    With PrimaryButton
        .name = "button4"
        .caption = "Bekræft" & Space(6) & ChrW(8594)
        .top = Form.properties("Height") - 25
        .left = 300
        .width = 70
    End With

    Form.codemodule.insertlines line, "Private Sub button3_Click()"
    line = line + 1
    Form.codemodule.insertlines line, "Unload Me"
    line = line + 1
    Form.codemodule.insertlines line, "End Sub"
    line = line + 1
    Form.codemodule.insertlines line, "Private Sub button4_Click()"
    line = line + 1
    Form.codemodule.insertlines line, action
    line = line + 1
    Form.codemodule.insertlines line, "End Sub"
    line = line + 1
    
    Form.properties("Height") = Form.properties("Height") + SecondaryButton.height + 15
 
End Function

Public Function renderForm()

    Set Form = ThisWorkbook.VBProject.VBComponents(identifier)
    Form.properties("Height") = Form.properties("Height") + 10

    Form.codemodule.insertlines line, "Private Sub UserForm_Initialize()"
    line = line + 1
    Form.codemodule.insertlines line, "Combo_Initialize"
    line = line + 1
    Form.codemodule.insertlines line, "List_Initialize"
    line = line + 1
    Form.codemodule.insertlines line, "End Sub"

    VBA.UserForms.Add(identifier).Show
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(identifier)
 
End Function