VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BarreProgressionEtConsole 
   Caption         =   "Progression du traitement"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12165
   OleObjectBlob   =   "BarreProgressionEtConsole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BarreProgressionEtConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const cstNbLigConsole = 15
Const cstNbCarLigne = 100
Dim iTauxPrecedent As Integer

' Afficher la barre de progression ( à 0%) et la console (sans texte)
Public Sub Afficher()
    ' Afficher le formulaire
    Me.Show 0
    ' Réduire la barre au minimum et afficher 0%
    iTauxPrecedent = -1
    Call Actualiser(0)
    ' Effacer le contenu de la console
    Call EffacerTexte
End Sub

' Fermer le formulaire
Public Sub Masquer()
    Unload Me
End Sub

' Actualiser la barre de progression
Public Sub Actualiser(iTaux As Integer)

    ' Si variation du taux alors on rafraichît la barre de progression
    If iTaux <> iTauxPrecedent Then
        iTauxPrecedent = iTaux
        ' Pourcentage affiché en noir (lorsque le fond est blanc) puis en blanc (sur fond bleu)
        If iTaux < 53 Then textePourcentage.ForeColor = vbBlack Else textePourcentage.ForeColor = vbWhite
        ' Modifier la longueur de la barre
        ImgBarreProgression.Width = (iTaux * textePourcentage.Width) / 100
        ' Modifier le pourcentage affiché
        textePourcentage = iTaux & " %"
        ' Rafraichir le formulaire
        DoEvents
    End If
    
End Sub

' Actualiser le texte affiché dans la console en concaténant le nouveau texte en paramètre
Public Sub AfficherTexte(sTexte As String)

    Dim sTabTxt() As String, iLigne As Integer, iNbLig As Integer, sTxtConsole As String
    
    ' Découpe le texte par ligne à partir du contenu affiché dans la console et du texte à ajouter
    sTabTxt = Split(Console & IIf(Console = "", "", vbCrLf) & sTexte, vbCrLf)
    
    ' Seules les 15 dernières lignes peuvent être affichées
    iNbLig = 0
    ' Texte à afficher dans la console
    sTxtConsole = ""

    ' Concatène les 15 (cstNbLigConsole) dernières lignes du texte à afficher
    For iLigne = UBound(sTabTxt) To LBound(sTabTxt) Step -1
        ' Calcule le nb de lignes nécessaires par blocs de texte
        iNbLig = iNbLig + 1 + (Len(sTabTxt(iLigne)) - 1) \ cstNbCarLigne
        ' Si l'ajout du bloc ne dépasse pas la limite d'affichage de la console
        If iNbLig <= cstNbLigConsole Then
            sTxtConsole = sTabTxt(iLigne) & IIf(sTxtConsole = "", "", vbCrLf) & sTxtConsole
        Else
            ' Limite atteinte, on cesse de concaténer des blocs de texte
            Exit For
        End If
    Next iLigne
    ' Réactualisation de la console
    Console = sTxtConsole
    DoEvents
    
End Sub

Public Sub EffacerTexte()
    Console = ""
    DoEvents
End Sub

' Clic sur le bouton "Fermer"
Private Sub Fermeture_Click()
    Call Masquer
End Sub
