VERSION 5.00
Begin VB.Form frmProperties 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Propriétés du noeud"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAide 
      Caption         =   "&Aide"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox chkPosForcee 
      Caption         =   "&Position forcée"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtLegende 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lbly 
      Caption         =   "Y"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblx 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "URL :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Légende"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmProperties
'Ajouté par C.Dutoit le 10 Août 2000

Option Explicit

Dim NoeudModifie As Long


Public Sub EditerNoeud(index As Long)
    If index < 0 Or index > UBound(Arbre) Then
        MsgBox "Indice non valide", vbCritical, "Erreur lors de l'édition du noeud..."
        Hide
        Exit Sub
    End If
    
    
    
    NoeudModifie = index
    
    'Afficher légende et URL
    txtLegende = Arbre(index).Legende
    txtURL = Arbre(index).URL
    
    'Afficher (x,y) si paramètres forcés, rien sinon
    If Arbre(index).PositionForcee Then
        chkPosForcee.Value = vbChecked
        EtatControlesXY (True)
        txtX = Arbre(index).x
        txtY = Arbre(index).y
    Else
        chkPosForcee.Value = vbUnchecked
        EtatControlesXY (False)
    End If
    
    Me.Show vbModal, frmMDI
End Sub 'EditerNoeud

Private Sub chkPosForcee_Click()
    EtatControlesXY (chkPosForcee.Value = vbChecked)
End Sub 'chkPosForcee_Click


'Modifie l'état des contrôles gérant les propriétés X et Y
Sub EtatControlesXY(Etat As Boolean)
    txtX.Enabled = Etat
    txtY.Enabled = Etat
    lblx.Enabled = Etat
    lbly.Enabled = Etat
    
    If Not Etat Then txtX = ""
    If Not Etat Then txtY = ""
End Sub 'EtatControlesXY


Private Sub cmdAide_Click()
    frmMDI.cmDlgHlp.HelpFile = App.Path & "\GMindmap.hlp"
    frmMDI.cmDlgHlp.HelpCommand = cdlHelpContext
    frmMDI.cmDlgHlp.HelpContext = 1 '"Edition_Paramètres"
    frmMDI.cmDlgHlp.ShowHelp  ' Afficher l'aide
End Sub

'Décharger la feuille
Private Sub cmdAnnuler_Click()
    'Annulation après création d'un noeud
    If Arbre(NoeudModifie).Legende = "" And Arbre(NoeudModifie).x = 0 And Arbre(NoeudModifie).y = 0 Then
        SupprimerNoeud (NoeudModifie)
    End If
    Unload Me
End Sub 'cmdAnnuler_Click


'Enregistrer les modifications et décharger la feuille
Private Sub cmdOK_Click()
    'Enregistrer les nouvelles propriétés dans l'arbre
    With Arbre(NoeudModifie)
        .Legende = txtLegende
        .URL = txtURL
        If chkPosForcee.Value = vbChecked Then
            .x = txtX
            .y = txtY
            .PositionForcee = True
        Else
            .PositionForcee = False
        End If
    End With
    
    'Redéfinir le titre ?
    MyApp.Modifie = True
    SetAppCaption
    
    'Mise à jour
    DessinerAllMindMap
    Unload Me
End Sub 'cmdOk_Click

