VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMacrosLibrosYRevistas 
   Caption         =   "Macros para estandarizar libros y artículos"
   ClientHeight    =   7230
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   6255
   OleObjectBlob   =   "FormMacrosLibrosYRevistas.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FormMacrosLibrosYRevistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
' Macros creadas por Edward Aníbal Vásquez.
' Contacto: edward_avg@hotmail.com
' GitHub: edanvagu

Private Sub CommandButtonCitasYBibliografia_Click()
    ConvertirCitasYBibliografiaATextoEstaticoT
    MsgBox ("Se convirtieron citas y bibliografía de campo dinámico a texto estático.")
End Sub

Private Sub CommandButtonComillonesPorDobles_Click()
    ConvertirComillonesAComillasT
    MsgBox ("Se convirtieron las comillas españolas por inglesas.")
End Sub

Private Sub CommandButtonCornisas_Click()
    EliminarEncabezadosYPiesDePaginaT
    MsgBox ("Se eliminaron los textos de los encabezados y los pies de página.")
End Sub

Private Sub CommandButtonCreaEstilos_Click()
    CrearEstilosT
    MsgBox ("Se crearon los estilos de niveles de titulación y citas indentadas.")
End Sub

Private Sub CommandButtonFormato_Click()
    AplicarFormatoBasicoT
    MsgBox ("El formato del documento se estableció con interlineado de 1.5 líneas, 8 ptos de espacio después de cada párrafo, letra Times New Roman y márgenes izquierdo/derecho de 3 cm, y superior/inferior de 2.5 cm.")
End Sub

Private Sub CommandButtonGuionesARayas_Click()
    ConvertirGuionesARayasT
    MsgBox ("Se convirtieron los guiones por rayas.")
End Sub

Private Sub CommandButtonInsertarPorcentaje_Click()
    InsertarEspacioFinoPorcentajeT
    MsgBox ("Se insertaron (y reemplazaron espacios normales por) espacios finos entre números y símbolos de porcentaje.")
End Sub

Private Sub CommandButtonNotaAntesPuntuacion_Click()
    MoverLlamadoANotaAPieDePaginaAntesDeSignoDePuntuacionT
    MsgBox ("Se movieron las referencias a notas al pie de página antes de los signos de puntuación.")
End Sub

Private Sub CommandButtonRectasPorTipograficas_Click()
    ConvertirComillasRectasPorTipograficasT
    MsgBox ("Se convirtieron las comillas rectas por tipográficas.")
End Sub

Private Sub CommandButtonDoblesPorComillones_Click()
    ConvertirComillasAComillonesT
    MsgBox ("Se convirtieron las comillas inglesas por españolas.")
End Sub

Private Sub CommandButtonResaltaRayas_Click()
    ResaltarRayasT
    MsgBox ("Se resaltaron las rayas.")
End Sub

Private Sub CommandButtonSiglasAVersalitas_Click()
    ConvertirSiglasAVersalitasT
End Sub

Private Sub CommandButtonEspaciosDobles_Click()
    QuitarEspaciosSobrantesT
    MsgBox ("Se eliminaron los espacios que sobraban.")
End Sub

Private Sub CommandButtonSaltosParrafosDobles_Click()
    QuitarSaltosDeParrafoSobrantesT
    MsgBox ("Se eliminaron los saltos de párrafo que sobraban.")
End Sub

Private Sub CommandButtonTabulaciones_Click()
    QuitarTabulacionesT
    MsgBox ("Se eliminaron las tabulaciones.")
End Sub

Private Sub CommandButtonControlContenido_Click()
    EliminarControlesDeContenidoT
    MsgBox ("Se eliminaron los controles de contenido.")
End Sub

Private Sub CommandButtonPorcentaje_Click()
    QuitarEspacioPorcentajeT
    MsgBox ("Se eliminaron los espacios entre número y símbolo de porcentaje.")
End Sub

Private Sub CommandButtonNotaDespuesPuntuacion_Click()
    MoverLlamadoANotaAPieDePaginaDespuesDeSignoDePuntuacionT
    MsgBox ("Se movieron las referencias a notas al pie de página después de los signos de puntuación.")
End Sub

Private Sub CommandButtonResaltaComillas_Click()
    ResaltarComillasT
    MsgBox ("Se resaltaron las comillas.")
End Sub

Private Sub CommandButtonResaltaCursivas_Click()
    ResaltarCursivasT
    MsgBox ("Se resaltaron las cursivas.")
End Sub

Private Sub CommandButtonResaltaSuperindices_Click()
    ResaltarSuperindicesT
    MsgBox ("Se resaltaron los superíndices.")
End Sub

Private Sub CommandButtonResaltaVersalitas_Click()
    ResaltarVersalitasT
    MsgBox ("Se resaltaron las versalitas.")
End Sub

Private Sub CommandButtonNivelesTitulacion_Click()
    InsertarMarcasNivelesTitulacionT
End Sub

Private Sub CommandButtonCitasIndentadas_Click()
    InsertarMarcasCitasIndentadasT
End Sub

Private Sub CommandButtonURL_Click()
    ActivarHipervinculosT
    MsgBox ("Se habilitaron los hipervínculos, convirtiendo texto plano en enlaces clicables.")
End Sub

Private Sub CommandButtonVersalitasAMayusculas_Click()
    ConvertirVersalitasAMayusculasT
    MsgBox ("Se convirtieron en mayúsculas las versalitas que estaban en el rango seleccionado, y se les quitó el resaltado (si lo tenían).")
End Sub

Private Sub Label25_Click()

End Sub

Private Sub Label33_Click()

End Sub

Private Sub Label61_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()
    Me.MultiPage1.Value = 0
End Sub
Private Sub InsertarMarcasCitasIndentadasT()

    Dim strQuoteStart As String
    Dim strQuoteEnd As String
    Dim intNumberOfQuotes As Integer
    Dim styleName As String
    styleName = "CitasIndentadas-Macros"
    
    On Error Resume Next
    Dim styleExists As Boolean
    styleExists = Not ActiveDocument.Styles(styleName) Is Nothing
    On Error GoTo 0

    If Not styleExists Then
        MsgBox "Primero debe ejecutar la macro ""Crear estilos marcas diagramación"".", vbExclamation
        Exit Sub
    End If
    
    ActiveDocument.Range(0, 0).Select
    
    QuitarSaltosDeParrafoSobrantesT
    LimpiarMarcasCitasIndentadasT
    
    intNumberOfQuotes = 0
    strQuoteStart = "[INICIO DE CITA] "
    strQuoteEnd = " [FIN DE CITA]"

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.style = ActiveDocument.Styles(styleName)
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
    End With
    Do While (True)
        If Selection.Find.Execute = True Then
            Selection.InsertBefore strQuoteStart
            Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.InsertAfter strQuoteEnd
            Selection.MoveRight unit:=wdCharacter, Count:=1
            Selection.EndOf
            intNumberOfQuotes = intNumberOfQuotes + 1
        Else
            Exit Do
        End If
    Loop
    
    AplicarColorMarcasCitasIndentadasT
    
    MsgBox ("Se agregaron etiquetas a " & intNumberOfQuotes & " citas indentadas")

    LimpiarFormatoT
    
End Sub
Private Sub LimpiarMarcasCitasIndentadasT()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[INICIO DE CITA] "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " [FIN DE CITA]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
End Sub
Private Sub AplicarColorMarcasCitasIndentadasT()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorRed
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .Text = "[INICIO DE CITA]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorRed
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .Text = "[FIN DE CITA]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub InsertarMarcasNivelesTitulacionT()

    Dim strArrayTLevels
    Dim strTLevel As String
    Dim intNumberOfT As Integer
    Dim TLevel As Variant
    Dim i As Integer
    Dim styleExists As Boolean
    
    strArrayTLevels = Array("T1-Macros", "T2-Macros", "T3-Macros", "T4-Macros", "T5-Macros", "T1-espanol-Macros", "T1-ingles-Macros", "T1-portugues-Macros")
    
    For Each TLevel In strArrayTLevels
        On Error Resume Next
        styleExists = Not ActiveDocument.Styles(TLevel) Is Nothing
        On Error GoTo 0

        If Not styleExists Then
            MsgBox "Primero debe ejecutar la macro ""Crear estilos marcas diagramación"".", vbExclamation
            Exit Sub
        End If
    Next TLevel
    
    ActiveDocument.Range(0, 0).Select
    
    QuitarSaltosDeParrafoSobrantesT
    LimpiarMarcasNivelesTitulacionT
    
    intNumberOfT = 0

    For i = 0 To UBound(strArrayTLevels)
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.style = ActiveDocument.Styles(strArrayTLevels(i))
        strTLevel = "[" & Mid(strArrayTLevels(i), 1, Len(strArrayTLevels(i)) - Len("-Macros")) & "] "
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
        End With
        Do While (True)
            If Selection.Find.Execute = True Then
                Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
                Selection.InsertBefore strTLevel
                Selection.MoveRight unit:=wdCharacter, Count:=1
                Selection.EndOf
                intNumberOfT = intNumberOfT + 1
            Else
                Exit Do
            End If
        Loop
        Selection.HomeKey wdStory
        
    Next i
   
    ResaltarMarcasNivelesTitulacionT
    
    MsgBox ("Se agregaron etiquetas a " & intNumberOfT & " títulos")

    LimpiarFormatoT

End Sub
Private Sub LimpiarMarcasNivelesTitulacionT()
    
    Dim strArrayTLevels
    Dim i As Integer
    
    strArrayTLevels = Array("[T1] ", "[T2] ", "[T3] ", "[T4] ", "[T5] ", "[T1-espanol] ", "[T1-ingles] ", "[T1-portugues] ")
    
    For i = 0 To UBound(strArrayTLevels)
    
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = strArrayTLevels(i)
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Next i
    
End Sub
Private Sub ResaltarMarcasNivelesTitulacionT()


    Dim strArrayTLevels
    Dim strHighlightcolor
    Dim i As Integer
    
    strArrayTLevels = Array("[T1]", "[T2]", "[T3]", "[T4]", "[T5]", "[T1-espanol]", "[T1-ingles]", "[T1-portugues]")
    strHighlightcolor = wdBrightGreen
    
    For i = 0 To UBound(strArrayTLevels)
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Options.DefaultHighlightColorIndex = strHighlightcolor
        Selection.Find.Replacement.Highlight = True
        With Selection.Find
            .Text = strArrayTLevels(i)
            .Replacement.Text = ""
            .Replacement.Font.Bold = True
            .Replacement.Font.Italic = False
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Next i

End Sub
Private Sub QuitarSaltosDeParrafoSobrantesT()

    QuitarEspaciosSobrantesT
    
    Dim ListSeparator As String
    
    ListSeparator = Application.International(wdListSeparator)
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^013){2" & ListSeparator & "}"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    LimpiarFormatoT
    
End Sub
Private Sub QuitarEspaciosSobrantesT()

    Dim rng As Range
    Dim ListSeparator As String
    
    ListSeparator = Application.International(wdListSeparator)

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "( ){2" & ListSeparator & "}"
                .Replacement.Text = "\1"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    QuitarEspacioInicioParrafoT
    QuitarEspacioFinalParrafoT
    
    LimpiarFormatoT
    
End Sub
Private Sub QuitarEspacioInicioParrafoT()

    ActiveDocument.Range(0, 0).Select

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^013) "
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub QuitarEspacioFinalParrafoT()

    ActiveDocument.Range(0, 0).Select

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " (^013)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub QuitarTabulacionesT()

    Dim rng As Range

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "^t"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .Execute Replace:=wdReplaceAll
            End With

            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub ResaltarCursivasT()

    Dim rng As Range

    Options.DefaultHighlightColorIndex = wdTurquoise

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Font.Italic = True
                .Replacement.ClearFormatting
                .Replacement.Highlight = True
                .Text = ""
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub ResaltarSuperindicesT()

    Dim rng As Range

    Options.DefaultHighlightColorIndex = wdPink

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Font.Superscript = True
                .Replacement.ClearFormatting
                .Replacement.Highlight = True
                .Text = ""
                .Replacement.Text = "^&"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "[°ºª\*]"
                .Replacement.Text = "^&"
                .Replacement.Highlight = True
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = True
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            
            
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub ConvertirSiglasAVersalitasT()

    Dim rng As Range
    Dim intNumberOfSmallCaps As Integer
    Dim ListSeparator As String

    intNumberOfSmallCaps = 0
    ListSeparator = Application.International(wdListSeparator)

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Text = "[A-Z]{2" & ListSeparator & "}"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = True
            End With

            While rng.Find.Execute
                rng.Case = wdLowerCase
                rng.Font.SmallCaps = True
                intNumberOfSmallCaps = intNumberOfSmallCaps + 1
                rng.Collapse Direction:=wdCollapseEnd
            Wend

            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng

    MsgBox ("Se convirtieron " & intNumberOfSmallCaps & " siglas, números romanos o mayúsculas de dos o más letras a versalitas en todo el documento.")

    LimpiarFormatoT

End Sub
Private Sub ConvertirVersalitasAMayusculasT()
    Dim selectedRange As Range
    Dim charac As Range

    If Selection.Type = wdSelectionIP Then
        MsgBox "No se tiene ningún texto seleccionado.", vbExclamation
        Exit Sub
    End If

    Set selectedRange = Selection.Range

    For Each charac In selectedRange.Characters
        If charac.Font.SmallCaps Then
            charac.Font.SmallCaps = False
            If Not charac.Font.Superscript Then
                charac.Text = UCase(charac.Text)
            End If
            charac.HighlightColorIndex = wdNoHighlight
        End If
    Next charac
End Sub
Private Sub ResaltarVersalitasT()

    Dim rng As Range

    Options.DefaultHighlightColorIndex = wdYellow

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                    .ClearFormatting
                    .Font.SmallCaps = True
                    .Font.AllCaps = False
                    .Replacement.ClearFormatting
                    .Replacement.Highlight = True
                    .Text = ""
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=wdReplaceAll
            End With
            
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub EliminarControlesDeContenidoT()
    Dim i As Integer
    Dim rng As Range
    Dim cc As ContentControl
    Dim controlCount As Integer
    
    For i = 1 To 10
        controlCount = 0

        For Each rng In ActiveDocument.StoryRanges
            Do
                For Each cc In rng.ContentControls
                    cc.Delete
                    controlCount = controlCount + 1
                Next cc

                Set rng = rng.NextStoryRange
            Loop While Not rng Is Nothing
        Next rng

        If controlCount = 0 Then Exit For
    Next i
        
End Sub
Private Sub MoverLlamadoANotaAPieDePaginaDespuesDeSignoDePuntuacionT()

    ActiveDocument.Range(0, 0).Select

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(^2)([\.\,\;\:\!\?])"
        .Replacement.Text = "\2\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    LimpiarFormatoT
    
End Sub
Private Sub MoverLlamadoANotaAPieDePaginaAntesDeSignoDePuntuacionT()

    ActiveDocument.Range(0, 0).Select

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "([\.\,\;\:\!\?])(^2)"
        .Replacement.Text = "\2\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    LimpiarFormatoT
    
End Sub
Private Sub ConvertirComillasRectasPorTipograficasT()

    Dim originalAutoFormat As Boolean
    Dim rng As Range

    originalAutoFormat = Options.AutoFormatAsYouTypeReplaceQuotes

    Options.AutoFormatAsYouTypeReplaceQuotes = True

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = """"
                .Replacement.Text = """"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "'"
                .Replacement.Text = "'"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    Options.AutoFormatAsYouTypeReplaceQuotes = originalAutoFormat
    
    LimpiarFormatoT
    
End Sub
Private Sub ConvertirComillasAComillonesT()
    
    Dim originalAutoFormat As Boolean
    Dim rng As Range
    
    originalAutoFormat = Options.AutoFormatAsYouTypeReplaceQuotes

    Options.AutoFormatAsYouTypeReplaceQuotes = True
    
    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = """"
                .Replacement.Text = """"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    Options.AutoFormatAsYouTypeReplaceQuotes = False

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = ChrW(8220)
                .Replacement.Text = "«"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = ChrW(8221)
                .Replacement.Text = "»"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    Options.AutoFormatAsYouTypeReplaceQuotes = originalAutoFormat
    
    LimpiarFormatoT
    
End Sub
Private Sub ConvertirComillonesAComillasT()
    
    Dim rng As Range
  
    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "«"
                .Replacement.Text = """"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "»"
                .Replacement.Text = """"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub ResaltarComillasT()

    Dim rng As Range
    Dim quote As Variant

    Options.DefaultHighlightColorIndex = wdTeal

    Dim quotes As Variant
    quotes = Array("""", "'", "«", "»")

    For Each rng In ActiveDocument.StoryRanges
        Do
            For Each quote In quotes
                With rng.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = quote
                    .Replacement.Text = ""
                    .Replacement.Highlight = True
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=wdReplaceAll
                End With
            Next quote
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub QuitarEspacioPorcentajeT()

    Dim rng As Range

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "([0-9]) (%)"
                .Replacement.Text = "\1\2"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "([0-9])^s(%)"
                .Replacement.Text = "\1\2"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub InsertarEspacioFinoPorcentajeT()

    Dim rng As Range

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "([0-9]) (%)"
                .Replacement.Text = "\1^s\2"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "([0-9])(%)"
                .Replacement.Text = "\1^s\2"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub ConvertirGuionesARayasT()

    Dim rng As Range

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "([ ^013])[-–]"
                .Replacement.Text = "\1—"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With

            With rng.Find
                .Text = "[-–]([ .,;])"
                .Replacement.Text = "—\1"
                .Execute Replace:=wdReplaceAll
            End With
            
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub ResaltarRayasT()

    Dim rng As Range

    Options.DefaultHighlightColorIndex = wdRed

    For Each rng In ActiveDocument.StoryRanges
        Do
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Replacement.Highlight = True
                .Text = "—"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .Execute Replace:=wdReplaceAll
            End With
            
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    LimpiarFormatoT
    
End Sub
Private Sub ConvertirCitasYBibliografiaATextoEstaticoT()

    Dim rngStory As Range
    Dim fld As Field
    Dim oRng As Range

    For Each rngStory In ActiveDocument.StoryRanges
        Set oRng = rngStory.Duplicate

        Do
            For Each fld In oRng.Fields
                If fld.Type = wdFieldCitation Then
                    fld.Select
                    Set oRng = Selection.Range
                    oRng.Start = oRng.Start - 1
                    oRng.End = oRng.End + 1
                    oRng.Text = fld.Result
                End If
            Next fld
            
            Set rngStory = rngStory.NextStoryRange
            If Not rngStory Is Nothing Then
                Set oRng = rngStory.Duplicate
            End If
        Loop While Not rngStory Is Nothing
    Next rngStory
    
    Dim rng As Range

    For Each rng In ActiveDocument.StoryRanges
        Do
            rng.Fields.Unlink
            Set rng = rng.NextStoryRange
            
        Loop While Not rng Is Nothing
    Next rng
    
    ActivarHipervinculosT
    LimpiarFormatoT

End Sub
Private Sub ActivarHipervinculosT()

    ActiveDocument.Range(0, 0).Select

    Dim originalAutoFormatApplyHeadings As Boolean
    Dim originalAutoFormatApplyLists As Boolean
    Dim originalAutoFormatApplyBulletedLists As Boolean
    Dim originalAutoFormatApplyOtherParas As Boolean
    Dim originalAutoFormatReplaceQuotes As Boolean
    Dim originalAutoFormatReplaceSymbols As Boolean
    Dim originalAutoFormatReplaceOrdinals As Boolean
    Dim originalAutoFormatReplaceFractions As Boolean
    Dim originalAutoFormatReplacePlainTextEmphasis As Boolean
    Dim originalAutoFormatReplaceHyperlinks As Boolean
    Dim originalAutoFormatPreserveStyles As Boolean
    Dim originalAutoFormatPlainTextWordMail As Boolean
    
    With Options
        originalAutoFormatApplyHeadings = .AutoFormatApplyHeadings
        originalAutoFormatApplyLists = .AutoFormatApplyLists
        originalAutoFormatApplyBulletedLists = .AutoFormatApplyBulletedLists
        originalAutoFormatApplyOtherParas = .AutoFormatApplyOtherParas
        originalAutoFormatReplaceQuotes = .AutoFormatReplaceQuotes
        originalAutoFormatReplaceSymbols = .AutoFormatReplaceSymbols
        originalAutoFormatReplaceOrdinals = .AutoFormatReplaceOrdinals
        originalAutoFormatReplaceFractions = .AutoFormatReplaceFractions
        originalAutoFormatReplacePlainTextEmphasis = .AutoFormatReplacePlainTextEmphasis
        originalAutoFormatReplaceHyperlinks = .AutoFormatReplaceHyperlinks
        originalAutoFormatPreserveStyles = .AutoFormatPreserveStyles
        originalAutoFormatPlainTextWordMail = .AutoFormatPlainTextWordMail
        
        .AutoFormatApplyHeadings = False
        .AutoFormatApplyLists = False
        .AutoFormatApplyBulletedLists = False
        .AutoFormatApplyOtherParas = False
        .AutoFormatReplaceQuotes = False
        .AutoFormatReplaceSymbols = False
        .AutoFormatReplaceOrdinals = False
        .AutoFormatReplaceFractions = False
        .AutoFormatReplacePlainTextEmphasis = False
        .AutoFormatReplaceHyperlinks = True
        .AutoFormatPreserveStyles = False
        .AutoFormatPlainTextWordMail = False
    End With
    
    Selection.Range.AutoFormat
    
        With Options
        .AutoFormatApplyHeadings = originalAutoFormatApplyHeadings
        .AutoFormatApplyLists = originalAutoFormatApplyLists
        .AutoFormatApplyBulletedLists = originalAutoFormatApplyBulletedLists
        .AutoFormatApplyOtherParas = originalAutoFormatApplyOtherParas
        .AutoFormatReplaceQuotes = originalAutoFormatReplaceQuotes
        .AutoFormatReplaceSymbols = originalAutoFormatReplaceSymbols
        .AutoFormatReplaceOrdinals = originalAutoFormatReplaceOrdinals
        .AutoFormatReplaceFractions = originalAutoFormatReplaceFractions
        .AutoFormatReplacePlainTextEmphasis = originalAutoFormatReplacePlainTextEmphasis
        .AutoFormatReplaceHyperlinks = originalAutoFormatReplaceHyperlinks
        .AutoFormatPreserveStyles = originalAutoFormatPreserveStyles
        .AutoFormatPlainTextWordMail = originalAutoFormatPlainTextWordMail
    End With

End Sub
Private Sub CrearEstilosT()

    Dim oStyle As style
    Dim strArrayStyles
    Dim styleName As Variant
    Dim styleExist As Boolean
    
    strArrayStyles = Array("T1-Macros", "T2-Macros", "T3-Macros", "T4-Macros", "T5-Macros", "T1-espanol-Macros", "T1-ingles-Macros", "T1-portugues-Macros", "CitasIndentadas-Macros")
    
    For Each styleName In strArrayStyles
        On Error Resume Next
        styleExist = Not ActiveDocument.Styles(styleName) Is Nothing
        On Error GoTo 0
        
        If Not styleExist Then
            Set oStyle = ActiveDocument.Styles.Add(Name:=styleName, Type:=wdStyleTypeParagraph)
            With oStyle
                If oStyle = "CitasIndentadas-Macros" Then
                    .Font.Size = 10
                    .Font.Bold = False
                    .ParagraphFormat.LeftIndent = CentimetersToPoints(2.5)
                Else
                    .Font.Size = 12
                    .Font.Bold = True
                    .ParagraphFormat.LeftIndent = CentimetersToPoints(0)
                End If
                .Font.Name = "Times New Roman"
                .Font.Color = wdBlack
                .ParagraphFormat.RightIndent = CentimetersToPoints(0)
                .ParagraphFormat.SpaceBefore = 0
                .ParagraphFormat.SpaceBeforeAuto = False
                .ParagraphFormat.SpaceAfter = 8
                .ParagraphFormat.SpaceAfterAuto = False
                .ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
                .ParagraphFormat.Alignment = wdAlignParagraphJustify
                .ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
                .ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
                .ParagraphFormat.CharacterUnitLeftIndent = 0
                .ParagraphFormat.CharacterUnitRightIndent = 0
                .ParagraphFormat.CharacterUnitFirstLineIndent = 0
                .ParagraphFormat.LineUnitBefore = 0
                .ParagraphFormat.LineUnitAfter = 0
                .NoSpaceBetweenParagraphsOfSameStyle = False
                .QuickStyle = True
                .NextParagraphStyle = "Normal"
            End With
        End If
    Next styleName
    
End Sub
Private Sub EliminarEncabezadosYPiesDePaginaT()

    Dim sec As Section
    Dim header As HeaderFooter
    Dim footer As HeaderFooter
        
    For Each sec In ActiveDocument.Sections
        For Each header In sec.Headers
            header.Range.Delete
        Next header
        For Each footer In sec.Footers
            footer.Range.Delete
        Next footer
    Next sec
    
End Sub
Private Sub AplicarFormatoBasicoT()

    Dim rngStory As Range
    
    With ActiveDocument.Content
        .Font.Name = "Times New Roman"
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceBeforeAuto = False
        .ParagraphFormat.SpaceAfter = 8
        .ParagraphFormat.SpaceAfterAuto = False
        .ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
        .ParagraphFormat.WidowControl = True
        .ParagraphFormat.KeepWithNext = False
        .ParagraphFormat.KeepTogether = False
        .ParagraphFormat.NoLineNumber = False
        .ParagraphFormat.Hyphenation = True
        .ParagraphFormat.MirrorIndents = False
        .ParagraphFormat.TextboxTightWrap = wdTightNone
        .ParagraphFormat.CollapsedByDefault = False
        .PageSetup.TopMargin = CentimetersToPoints(2.5)
        .PageSetup.BottomMargin = CentimetersToPoints(2.5)
        .PageSetup.LeftMargin = CentimetersToPoints(3)
        .PageSetup.RightMargin = CentimetersToPoints(3)
    End With

    For Each rngStory In ActiveDocument.StoryRanges
        Do
            rngStory.Font.Name = "Times New Roman"
            Set rngStory = rngStory.NextStoryRange
        Loop While Not rngStory Is Nothing
    Next

    LimpiarFormatoT
    
End Sub
Private Sub LimpiarFormatoT()

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
End Sub
