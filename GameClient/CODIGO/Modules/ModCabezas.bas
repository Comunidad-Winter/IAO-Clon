Attribute VB_Name = "modCabezas"
Option Explicit
Public MinEleccion As Integer, MaxEleccion As Integer
Public Actual As Integer
 
Sub DameOpciones()
 
Dim i As Integer
 
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = Head_Range(HUMANO).mStart
                MaxEleccion = Head_Range(HUMANO).mEnd
                MinEleccion = Head_Range(HUMANO).mStart
            Case "Elfo"
                Actual = Head_Range(ELFO).mStart
                MaxEleccion = Head_Range(ELFO).mEnd
                MinEleccion = Head_Range(ELFO).mStart
            Case "Drow"
                Actual = Head_Range(ElfoOscuro).mStart
                MaxEleccion = Head_Range(ElfoOscuro).mEnd
                MinEleccion = Head_Range(ElfoOscuro).mStart
            Case "Enano"
                Actual = Head_Range(Enano).mStart
                MaxEleccion = Head_Range(Enano).mEnd
                MinEleccion = Head_Range(Enano).mStart
            Case "Gnomo"
                Actual = Head_Range(Gnomo).mStart
                MaxEleccion = Head_Range(Gnomo).mEnd
                MinEleccion = Head_Range(Gnomo).mStart

            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = Head_Range(HUMANO).fStart
                MaxEleccion = Head_Range(HUMANO).fEnd
                MinEleccion = Head_Range(HUMANO).fStart
            Case "Elfo"
                Actual = Head_Range(ELFO).fStart
                MaxEleccion = Head_Range(ELFO).fEnd
                MinEleccion = Head_Range(ELFO).fStart
            Case "Drow"
                Actual = Head_Range(ElfoOscuro).fStart
                MaxEleccion = Head_Range(ElfoOscuro).fEnd
                MinEleccion = Head_Range(ElfoOscuro).fStart
            Case "Enano"
                Actual = Head_Range(Enano).fStart
                MaxEleccion = Head_Range(Enano).fEnd
                MinEleccion = Head_Range(Enano).fStart
            Case "Gnomo"
                Actual = Head_Range(Gnomo).fStart
                MaxEleccion = Head_Range(Gnomo).fEnd
                MinEleccion = Head_Range(Gnomo).fStart

            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
End Select
 
frmCrearPersonaje.HeadView.Cls
Call DrawGrhToHdc(frmCrearPersonaje.HeadView.hdc, HeadData(Actual).Head(3).GrhIndex, 5, 5)
frmCrearPersonaje.HeadView.Refresh
End Sub



Public Sub DrawGrhToHdc(ByVal desthDC As Long, Grh As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer)
 
On Error GoTo Err
 
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    Dim grh_index As Integer
   
    grh_index = Grh
 
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).NumFrames = 0 Then Exit Sub
 
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If
 
    src_x = GrhData(grh_index).sX
    src_y = GrhData(grh_index).sY
    src_width = GrhData(grh_index).pixelWidth
    src_height = GrhData(grh_index).pixelHeight
           
    hdcsrc = CreateCompatibleDC(desthDC)
 
    file_path = App.path & "\resources\graphics\" & GrhData(grh_index).FileNum & ".bmp"
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
   
    'BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
   TransparentBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, src_width, src_height, RGB(0, 0, 0)
 
   Call DeleteObject(SelectObject(hdcsrc, PrevObj))
   DeleteDC hdcsrc
 
Err:
If Err.number = 481 Then MsgBox "Imposible cargar recurso error: " & "481"
 
End Sub
