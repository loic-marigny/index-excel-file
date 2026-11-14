Attribute VB_Name = "ModuleSommaire"
Option Explicit

Sub CréerSommaire()

    Dim wsSommaire As Worksheet
    Dim ws As Worksheet
    Dim i As Integer
    Dim ligne As Integer

    Application.ScreenUpdating = False

    ' Supprimer l'ancien sommaire s'il existe
    On Error Resume Next
    Set wsSommaire = ThisWorkbook.Worksheets("Sommaire")
    If Not wsSommaire Is Nothing Then
        Application.DisplayAlerts = False
        wsSommaire.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Créer un nouveau sommaire
    Set wsSommaire = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
    wsSommaire.Name = "Sommaire"

    ' Titre du sommaire
    With wsSommaire.Range("A1")
        .Value = "Sommaire du fichier"
        .Font.Bold = True
        .Font.Size = 16
    End With

    ligne = 3 ' Commencer à la ligne 3 pour laisser de l'espace

    ' Parcours de tous les onglets pour les ajouter au sommaire
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Sommaire" Then
            With wsSommaire
                .Hyperlinks.Add Anchor:=.Cells(ligne, 1), _
                                Address:="", _
                                SubAddress:="'" & ws.Name & "'!A1", _
                                TextToDisplay:=ws.Name
                ' Stylisation
                .Cells(ligne, 1).Font.Size = 12
                .Cells(ligne, 1).Font.Underline = xlUnderlineStyleSingle
                .Cells(ligne, 1).Font.Color = RGB(0, 102, 204)
            End With
            ligne = ligne + 1
        End If
    Next ws

    ' Ajuster la largeur de la colonne
    wsSommaire.Columns("A:A").AutoFit

    Application.ScreenUpdating = True
    MsgBox "Onglet Sommaire créé avec succès !", vbInformation

End Sub
