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

' Afficher la barre de progression ( � 0%) et la console (sans texte)
Public Sub Afficher()
    ' Afficher le formulaire
    Me.Show 0
    ' R�duire la barre au minimum et afficher 0%
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

    ' Si variation du taux alors on rafraich�t la barre de progression
    If iTaux <> iTauxPrecedent Then
        iTauxPrecedent = iTaux
        ' Pourcentage affich� en noir (lorsque le fond est blanc) puis en blanc (sur fond bleu)
        If iTaux < 53 Then textePourcentage.ForeColor = vbBlack Else textePourcentage.ForeColor = vbWhite
        ' Modifier la longueur de la barre
        ImgBarreProgression.Width = (iTaux * textePourcentage.Width) / 100
        ' Modifier le pourcentage affich�
        textePourcentage = iTaux & " %"
        ' Rafraichir le formulaire
        DoEvents
    End If
    
End Sub

' Actualiser le texte affich� dans la console en concat�nant le nouveau texte en param�tre
Public Sub AfficherTexte(sTexte As String)

    Dim sTabTxt() As String, iLigne As Integer, iNbLig As Integer, sTxtConsole As String
    
    ' D�coupe le texte par ligne � partir du contenu affich� dans la console et du texte � ajouter
    sTabTxt = Split(Console & IIf(Console = "", "", vbCrLf) & sTexte, vbCrLf)
    
    ' Seules les 15 derni�res lignes peuvent �tre affich�es
    iNbLig = 0
    ' Texte � afficher dans la console
    sTxtConsole = ""

    ' Concat�ne les 15 (cstNbLigConsole) derni�res lignes du texte � afficher
    For iLigne = UBound(sTabTxt) To LBound(sTabTxt) Step -1
        ' Calcule le nb de lignes n�cessaires par blocs de texte
        iNbLig = iNbLig + 1 + (Len(sTabTxt(iLigne)) - 1) \ cstNbCarLigne
        ' Si l'ajout du bloc ne d�passe pas la limite d'affichage de la console
        If iNbLig <= cstNbLigConsole Then
            sTxtConsole = sTabTxt(iLigne) & IIf(sTxtConsole = "", "", vbCrLf) & sTxtConsole
        Else
            ' Limite atteinte, on cesse de concat�ner des blocs de texte
            Exit For
        End If
    Next iLigne
    ' R�actualisation de la console
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
