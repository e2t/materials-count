VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Количество деталей"
   ClientHeight    =   9705.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18315
   OleObjectBlob   =   "MainForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnCancel_Click()
    ExitApp
End Sub

Private Sub btnRun_Click()
    FilterAndPrint
    Me.txtFilter.SetFocus
End Sub

Private Sub chkCurDir_Click()
    ResearchMaterials
End Sub

Private Sub lstDeps_Change()
    If Me.lstDeps.ListIndex >= 0 Then
        ShowWhereIsPartUsed Me.lstDeps.ListIndex
    End If
End Sub

Private Sub lstParts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.lstParts.ListIndex >= 0 Then
        OpenPart Me.lstParts.ListIndex
    End If
End Sub

Private Sub txtFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FilterAndPrint
        Me.txtFilter.SetFocus
        KeyCode = 0
    End If
End Sub

Private Sub UserForm_Initialize()
    '''The resize must be first!
    'Me.Width = MaximizedWidth
    'Me.Height = MaximizedHeight
    '''The resize must be first!
    
    Me.lstDeps.ColumnWidths = ";70;70"
    Me.txtFilter.SetFocus
End Sub

Private Sub UserForm_Resize()
    Me.btnCancel.Top = Me.Height - 51

    Me.btnCancel.Left = Me.Width - 81

    Me.lstDeps.Width = Me.Width - 15
    Me.lstDeps.Height = Me.Height - 61
End Sub
