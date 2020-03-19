VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents conLabel As MSForms.Label
Attribute conLabel.VB_VarHelpID = -1
Private WithEvents conComboBox As MSForms.ComboBox
Attribute conComboBox.VB_VarHelpID = -1
Private WithEvents conCommandButton As MSForms.CommandButton
Attribute conCommandButton.VB_VarHelpID = -1

Property Get myDate() As Date
    With UsFormDateControl
        myDate = CDate(.Caption)
    End With
End Property

Public Sub ReceiveLabel(ByVal reLabel As MSForms.Label)
    Set conLabel = reLabel
End Sub

Public Sub ReceiveComboBox(ByVal reComboBox As MSForms.ComboBox)
    Set conComboBox = reComboBox
End Sub

Public Sub ReceiveCommandButton(ByVal reCommandButton As MSForms.CommandButton)
    Set conCommandButton = reCommandButton
End Sub

Private Sub conComboBox_Change()
    With UsFormDateControl
        .AddLabel_Day DateSerial(.Controls("ComboBoxYear"), .Controls("ComboBoxMonth"), Day(.Caption))
    End With
End Sub

Private Sub conCommandButton_Click()
    Dim iTmp As Integer
    With UsFormDateControl
        Select Case conCommandButton.Name
            Case "Year-"
                iTmp = .Controls("ComboBoxYear").Value
                If iTmp <> 1900 Then .Controls("ComboBoxYear").Value = iTmp - 1
            Case "Year+"
                iTmp = .Controls("ComboBoxYear").Value
                If iTmp <> 2999 Then .Controls("ComboBoxYear").Value = iTmp + 1
            Case "Month-"
                iTmp = .Controls("ComboBoxMonth").Value
                .Controls("ComboBoxMonth").Value = IIf(iTmp - 1 Mod 12, iTmp - 1, 12)
            Case "Month+"
                iTmp = .Controls("ComboBoxMonth").Value
                .Controls("ComboBoxMonth").Value = IIf(iTmp Mod 12, iTmp + 1, 1)
        End Select
    End With
End Sub

Private Sub conLabel_Click()
    Dim sTmp As String
    
    With UsFormDateControl
        .Caption = CDate(Replace(conLabel.Name, "LabelDay", ""))
        
        sTmp = .sLabelName
     
        conLabel.BackColor = RGB(0, 100, 250)
        If sTmp <> conLabel.Name And Len(sTmp) > 0 Then
            On Error Resume Next
            .Controls(sTmp).BackColor = RGB(230, 230, 230)
            On Error GoTo 0
        End If
        .sLabelName = conLabel.Name
        
        '如果选中其他月份的标签，重置日期。
        If Month(.Caption) <> Val(.Controls("ComboBoxMonth").Value) Then
            .AddLabel_Day myDate
            .Controls("ComboBoxMonth").Value = Month(.Caption)
        End If
    End With
    
    If TypeName(Selection) = "Range" Then
        Selection = myDate
    End If
    
    Unload UsFormDateControl
End Sub
