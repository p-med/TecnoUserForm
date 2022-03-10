VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Primary 
   Caption         =   "Carga de Datos - TF Paracel"
   ClientHeight    =   11790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18075
   OleObjectBlob   =   "Primary.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Primary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Paulo Medina
'2022

'*************************************************************************************************************************
'**********************************************FUNCIONES******************************************************************
'*************************************************************************************************************************

'Definicion de funcion. MsgBox time out
Private Declare PtrSafe Function CustomTimeOffMsgBox Lib "user32" Alias "MessageBoxTimeoutA" ( _
            ByVal xHwnd As LongPtr, _
            ByVal xText As String, _
            ByVal xCaption As String, _
            ByVal xMsgBoxStyle As VbMsgBoxStyle, _
            ByVal xwlange As Long, _
            ByVal xTimeOut As Long) _
    As Long

'***********************************************************************************************************************
'*******************************************FUNCION DE BOTONES**********************************************************
'***********************************************************************************************************************

Private Sub btn_Guardar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

'Definicion de variables
Dim ws As Worksheet
Dim wbLocal As Workbook
Dim wbMaster As Workbook
Dim addnew As Range
Dim f As Range
Dim r As Range
Dim tbl As ListObject, tRow As ListRow, tCol As ListColumn

'Definicion de elementos
Set wbLocal = ThisWorkbook
'Set wbMaster = Workbooks.Open("C:\Users\rosem\OneDrive - Smithsonian Institution\Documents\Cae\Avances_Paracel.xlsx")
Set ws = wbLocal.Worksheets("Datos")
Set tbl = ws.ListObjects("Data")
'ListObjects ("Data")

'Datos de input boxes
Set tRow = tbl.ListRows.Add
With tRow
    .Range(1).Value = fecha.Value
    .Range(2).Value = Actividad.Value
    .Range(4).Value = Avance.Value
    .Range(8).Value = Parcela.Value
    .Range(9).Value = TM.Value
    If Me.OptTerminado.Value Then .Range(10).Value = "Terminado"
    If Me.OptionButton1.Value Then .Range(10).Value = "En curso"
    
End With
    
Call CustomTimeOffMsgBox(0, "Datos cargados exitosamente.", "Carga de Datos - TF", vbInformation, 0, 1000)

    
End Sub

Private Sub btn_Limpiar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Boton de Limpiar - Limpia todos los valores menos fecha y Estado
Me.Parcela.Value = ""
Me.TM.Value = ""
Me.Avance.Value = ""
Me.Actividad.Value = ""

End Sub


Private Sub btnAbtInactive_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub imgSettingsActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Boton de AJUSTES - Activo
    frmCD.Visible = False
    frmAjustes.Visible = True
End Sub

Private Sub btnVisualActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Boton de VISUALIZACIONES - Activo
    btnInactiveViz.Visible = False
    btnVisualInactive.Visible = False
    btnCDInactive.Visible = True
    CDInactive.Visible = True
    VizOnClick.Visible = True
    OnClickCD.Visible = False
    AjInactive.Visible = True
    AjustesOnClick.Visible = False

    'ThisWorkbook.RefreshAll
    
End Sub

Private Sub AjustesActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Boton de AJUSTES - Activo

    AjustesOnClick.Visible = True
    AjInactive.Visible = False
    btnCDInactive.Visible = True
    CDInactive.Visible = True
    VizOnClick.Visible = False
    OnClickCD.Visible = False
    btnInactiveViz.Visible = True
    frmAjustes.Visible = True
    
End Sub

Private Sub btnCDActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Boton de CARGA DE DATOS - Activo
    btnInactiveViz.Visible = True
    btnVisualInactive.Visible = True
    btnCDInactive.Visible = False
    CDInactive.Visible = False
    VizOnClick.Visible = False
    OnClickCD.Visible = True
    frmAjustes.Visible = False
    AjustesOnClick.Visible = False
    AjInactive.Visible = True

    
End Sub

Private Sub btnAdvSetActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Application.Visible = True
    
End Sub

'******************************************************************************************************************************
'**************************************************HOVER EFFECTS***************************************************************
'******************************************************************************************************************************

Private Sub btnVisualInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Boton de VISUALIZACIONES - Hover effect
    btnVisualInactive.Visible = False
End Sub

Private Sub btnCDInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Boton de CARGA DE DATOS - Hover effect
    btnCDInactive.Visible = False
End Sub

Private Sub AjustesInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Boton de AJUSTES - Hover effect
    AjustesInactive.Visible = False
End Sub

Sub btnInactiveClean_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make LIMPIAR Button appear Green when hovered on

    btnInactive.Visible = True
    btnInactiveClean.Visible = False

End Sub

Sub btnInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make GUARDAR Button appear Green when hovered on

  btnInactive.Visible = False
  btnInactiveClean.Visible = True

End Sub

Sub imgSettingsInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make SETTINGS Button appear Green when hovered on

    imgSettingsActive.Visible = True
    imgSettingsInactive.Visible = False

End Sub

'frmAjustes
'******************************************************************************************************************************

Sub btnAbtInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make ABOUT Button appear Green when hovered on
    btnAbtInactive.Visible = False
End Sub

Sub btnAdvSettInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make ADVANCE SETTINGS Button appear Green when hovered on
    btnAdvSettInactive.Visible = False
End Sub

Sub btnHelpInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make HELP Button appear Green when hovered on
    btnHelpInactive.Visible = False
End Sub

'******************************************************************************************************************************

Sub frmSideMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

    btnInactive.Visible = True
    btnInactiveClean.Visible = True
    
    imgSettingsInactive.Visible = True
    btnVisualInactive.Visible = True
    
    btnCDInactive.Visible = True
    
    AjustesInactive.Visible = True
End Sub

Sub frmCD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

    btnInactive.Visible = True
    btnInactiveClean.Visible = True
    imgSettingsInactive.Visible = True
    btnVisualInactive.Visible = True
    AjustesInactive.Visible = True
End Sub

Sub frmAjustes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnAbtInactive.Visible = True
    btnAdvSettInactive.Visible = True
    btnHelpInactive.Visible = True
End Sub

'******************************************************************************************************************************
'***********************************************USER FORM INITIALIZE AND MISC**************************************************
'******************************************************************************************************************************

Private Sub UserForm_Initialize()
  'Carga de datos activa al iniciar
  OnClickCD.Visible = True
  CDInactive.Visible = False
  btnCDInactive.Visible = False
  'Fecha del dia como default
  Dim dtToday As Date
  dtToday = Date
  fecha.Value = dtToday
  VizOnClick.Visible = False
  AjustesOnClick.Visible = False
  'Frame CARGA DE DATOS como default at initialize
  frmAjustes.Visible = False
  

End Sub

Private Sub fecha_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    'Fecha cargada
    On Error Resume Next
    Me.fecha = CDate(Me.fecha)

End Sub


