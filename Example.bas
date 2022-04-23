Sub Example()

    Dim Form As New UserInterface

    With Form
        .createForm     "Window"
        .addLabel       "Label label label label label label label label."
        .addInput
        .AddAttach      "msgbox (""File attached!"")"
        .addDivider
        .addDropdown    "Apple", "Banana", "Clementine"
        .addToggle      "Description"
        .addDivider
        .addList        "Apple", "Banana", "Clementine"
        .addActions     "msgbox (""Confirmed!"")"
        .renderForm
    End With

End Sub
