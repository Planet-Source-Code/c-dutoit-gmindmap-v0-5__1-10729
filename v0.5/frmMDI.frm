VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5745
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   9030
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmDlgHlp 
      Left            =   1800
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":0D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":13F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Nouveau mindmap"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Ouvrir un mindmap"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Enregistrer un mindmap"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Imprimer le mindmap"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Insérer un noeud"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Supprimer un noeud"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmDlgImprimer 
      Left            =   1440
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Impression d'un Mindmap"
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuFichierNouveau 
         Caption         =   "&Nouveau"
      End
      Begin VB.Menu mnuFichierOuvrir 
         Caption         =   "&Ouvrir..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFichierEnregistrer 
         Caption         =   "&Enregistrer"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFichierEnregistrerSous 
         Caption         =   "Enregistrer &sous..."
      End
      Begin VB.Menu mnuFichierSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFichierImporter 
         Caption         =   "Im&porter"
         Begin VB.Menu mnuFichierImporterTxt 
            Caption         =   "&Fichier texte..."
         End
      End
      Begin VB.Menu mnuFichierExporter 
         Caption         =   "&Exporter"
         Begin VB.Menu mnuFichierExporterTxt 
            Caption         =   "&Fichier Texte..."
         End
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFichierMEP 
         Caption         =   "&Mise en page"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFichierImprimer 
         Caption         =   "&Imprimer"
      End
      Begin VB.Menu mnuFichierSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFichierLast 
         Caption         =   "1."
         Index           =   1
      End
      Begin VB.Menu mnuFichierLast 
         Caption         =   "2."
         Index           =   2
      End
      Begin VB.Menu mnuFichierLast 
         Caption         =   "3."
         Index           =   3
      End
      Begin VB.Menu mnuFichierLast 
         Caption         =   "4."
         Index           =   4
      End
      Begin VB.Menu mnuFichierLast 
         Caption         =   "5."
         Index           =   5
      End
      Begin VB.Menu mnuFichierSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFichierQuitter 
         Caption         =   "&Quitter"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuNoeud 
      Caption         =   "&Noeud"
      Begin VB.Menu mnuNoeudInsererFils 
         Caption         =   "&Insérer un fils"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuNoeudSupprimer 
         Caption         =   "&Supprimer le noeud"
      End
      Begin VB.Menu mnuNoeudsAffNPosForcee 
         Caption         =   "&Afficher les noeuds à position forcée"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuNoeudDeforcer 
         Caption         =   "&Déforcer toutes les positions forcées"
      End
   End
   Begin VB.Menu mnuAide 
      Caption         =   "&?"
      Begin VB.Menu mnuAideIndex 
         Caption         =   "&Index"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAideContextuelle 
         Caption         =   "Aide contextuelle"
      End
      Begin VB.Menu mnuAideSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAideNetProduit 
         Caption         =   "&Téléchargement du produit"
      End
      Begin VB.Menu mnuAideSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAideAPropos 
         Caption         =   "&A propos de..."
      End
   End
   Begin VB.Menu mnuPopFrmMap 
      Caption         =   "frmMapPopMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopFrmMapInsFils 
         Caption         =   "&Insérer un fils..."
      End
      Begin VB.Menu mnuPopFrmMapSupFils 
         Caption         =   "&Supprimer un fils..."
      End
      Begin VB.Menu mnuPopFrmMapForcerPos 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmMDI
Option Explicit




Private Sub MDIForm_Load()
    'Initialisation de la feuille
    SetAppCaption
    ShowLastOpenedFiles
End Sub 'MDIForm_Load


'Afficher la boîte de dialogue "à propos de..."
Private Sub mnuAideAPropos_Click()
    frmAbout.Show vbModal
End Sub 'mnuAideAPropos_Click


'Mettre ou non l'aide contextuelle
Private Sub mnuAideContextuelle_Click()
    If frmMDI.MousePointer = 99 Then '=>pointeur normal
        frmMDI.MousePointer = 0
    Else '=>pointeur d'aide contextuelle
        frmMDI.MousePointer = 99
        frmMDI.MouseIcon = LoadPicture(App.Path & "\img\aidecontextuelle.ico")
    End If
End Sub 'mnuAideContextuelle_Click


'Afficher l'aide
Private Sub mnuAideIndex_Click()
    cmDlgHlp.HelpFile = App.Path & "\GMindmap.hlp"
    cmDlgHlp.HelpCommand = cdlHelpIndex
    cmDlgHlp.ShowHelp  ' Afficher l'aide
End Sub 'mnuAideIndex_Click


'Afficher la page du produit
Private Sub mnuAideNetProduit_Click()
    Shell "start http://www.home.ch/~spaw4758/redirect.html?gmindmapv04fr"
End Sub 'mnuAideNetProduit_Click


'Enregistrer le mindmap
Private Sub mnuFichierEnregistrer_Click()
    If MyApp.Fichier = "" Then EnregistrerSous Else SauverArbre (MyApp.Fichier)
End Sub 'mnuFichierEnregistrer_Click


'Demander le nom du fichier de destination et enregistrer le mindmap
Private Sub mnuFichierEnregistrerSous_Click()
    EnregistrerSous
End Sub 'mnuFichierEnregistrerSous_Click


'Exporter dans un fichier texte
Private Sub mnuFichierExporterTxt_Click()
    Dim Filename As String
    Filename = Exporter_DemanderNomFichier
  
    If Filename <> "" Then ExporterTexte (cmDlg.Filename) Else MsgBox "L'exportation a echoué !", vbExclamation, "Erreur à l'exportation..."
End Sub 'mnuFichierExporterTxt_Click


'Importer depuis un fichier texte
Private Sub mnuFichierImporterTxt_Click()
    Dim Filename As String
    Filename = Importer_DemanderNomFichier
  
    If Filename <> "" Then ImporterArbre (cmDlg.Filename) Else MsgBox "L'importation a echoué !", vbExclamation, "Erreur à l'importation..."
End Sub 'mnuFichierImporterTxt_Click

Private Sub mnuFichierImprimer_Click()
    'DessinerAllMindMap
    'frmMap.BackColor = RGB(255, 255, 255)
    
    ImprimerMindmap
End Sub 'mnuFichierImprimer_Click


Private Sub mnuFichierLast_Click(index As Integer)
   Dim Filename As String
   Dim IndiceFenetre As Integer
   
   'Lire le nom du fichier dans la base de registres
   Filename = GetSetting(App.EXEName, "Options", "File" & index)
   
   'Ouvrir le fichier
   If FileExist(Filename) Then
       'Enregistrer si modifie
        If MyApp.Modifie Then
            Select Case MsgBox("Voulez-vous enregistrer le fichier actuel ? ", vbYesNoCancel, "Nouveau fichier")
                Case vbYes: If Not EnregistrerSous Then Exit Sub
                Case vbNo:  'null
                Case vbCancel: Exit Sub
            End Select
        End If
       ChargerArbre (Filename)
   End If
End Sub 'mnuFichierLast_Click

Private Sub mnuFichierNouveau_Click()
    'Enregistrer si modifie
    If MyApp.Modifie Then
        Select Case MsgBox("Voulez-vous enregistrer le fichier actuel ? ", vbYesNoCancel, "Nouveau fichier")
            Case vbYes: If Not EnregistrerSous Then Exit Sub
            Case vbNo:  'null
            Case vbCancel: Exit Sub
        End Select
    End If
    
    NouveauFichier
End Sub 'mnuFichierNouveau_Click



Private Sub mnuFichierOuvrir_Click()
    'Enregistrer si modifie
    If MyApp.Modifie Then
        Select Case MsgBox("Voulez-vous enregistrer le fichier actuel ? ", vbYesNoCancel, "Nouveau fichier")
            Case vbYes: If Not EnregistrerSous Then Exit Sub
            Case vbNo:  'null
            Case vbCancel: Exit Sub
        End Select
    End If
    
    'Demander quel fichier ouvrir
    Dim Fichier
    Fichier = Ouvrir_DemanderNomFichier
    If Fichier <> "" Then
        'Ouvrir le fichier
        ChargerArbre (Fichier)
        DessinerAllMindMap
    End If
End Sub 'mnuFichierOuvrir_Click



'Quitter le programme
Private Sub mnuFichierQuitter_Click()
    If MyApp.Modifie Then
        Select Case MsgBox("Voulez-vous enregistrer le fichier actuel ? ", vbYesNoCancel, "Nouveau fichier")
            Case vbYes: If Not EnregistrerSous Then Exit Sub
            Case vbNo:  'null
            Case vbCancel: Exit Sub
        End Select
    End If

    Unload frmMap
    Unload frmAbout
    Unload frmMDI
    Unload frmProperties
End Sub 'mnuFichierQuitter_Click




'Déforcer toutes les positions forcées
Private Sub mnuNoeudDeforcer_Click()
    Dim i As Long
    
    For i = 0 To UBound(Arbre)
        Arbre(i).PositionForcee = False
    Next i
    DessinerAllMindMap
End Sub 'mnuNoeudDeforcer_Click



Sub mnuNoeudInsererFils_Click()
    CreerFils (NoeudSelectionne)
    frmProperties.EditerNoeud (UBound(Arbre))
End Sub 'mnuNoeudInsererFils_Click




'Supprimer le noeud en cours
Private Sub mnuNoeudSupprimer_Click()
    SupprimerNoeud (NoeudSelectionne)
        'Définir le titre de la fenêtre principale + ...
    If Not MyApp.Modifie Then
        MyApp.Modifie = True
        SetAppCaption
    End If
End Sub 'mnuNoeudSupprimer_Click



'Demander le nom du fichier et enregistrer. true en sortie si tout s'est bien passé
Function EnregistrerSous() As Boolean
    Dim Filename As String
    Filename = EnregistrerSous_DemanderNomFichier
  
    If Filename <> "" Then SauverArbre (cmDlg.Filename)
End Function 'EnregistrerSous


'Demander le nom de fichier pour la procédure Enregistrer sous
'Note : il parait qu'il ne faut pas abréger les noms !
Function EnregistrerSous_DemanderNomFichier() As String
    On Error GoTo suite
    cmDlg.DialogTitle = "Enregistrement du mindmap..."
    cmDlg.Filter = "Fichiers G-Mindmap (*.gmm)|*.gmm|Tous les fichiers (*.*)|*.*"
    cmDlg.ShowSave
    EnregistrerSous_DemanderNomFichier = cmDlg.Filename
    
    Exit Function
    
suite: 'Traitement des erreurs (bouton annuler !)
    EnregistrerSous_DemanderNomFichier = ""
End Function 'EnregistrerSous_DemanderNomFichier



'Demander le nom de fichier pour la procédure Ouvrir
'Note : il parait qu'il ne faut pas abréger les noms !
Function Ouvrir_DemanderNomFichier() As String
    On Error GoTo suite
    cmDlg.DialogTitle = "Ouverture d'un mindmap..."
    cmDlg.Filter = "Fichiers G-Mindmap (*.gmm)|*.gmm|Tous les fichiers (*.*)|*.*"
    cmDlg.ShowOpen
    Ouvrir_DemanderNomFichier = cmDlg.Filename
    
    Exit Function
    
suite: 'Traitement des erreurs (bouton annuler !)
    Ouvrir_DemanderNomFichier = ""
End Function 'EnregistrerSous_DemanderNomFichier



'Demander le nom de fichier pour la procédure Exporter
'Note : il parait qu'il ne faut pas abréger les noms !
Function Exporter_DemanderNomFichier() As String
    On Error GoTo suite
    cmDlg.DialogTitle = "Exportation du mindmap..."
    cmDlg.Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
    cmDlg.ShowSave
    Exporter_DemanderNomFichier = cmDlg.Filename
    
    Exit Function
    
suite: 'Traitement des erreurs (bouton annuler !)
    Exporter_DemanderNomFichier = ""
End Function 'Exporter_DemanderNomFichier


'Demander le nom de fichier pour la procédure Importer
'Note : il parait qu'il ne faut pas abréger les noms !
Function Importer_DemanderNomFichier() As String
    On Error GoTo suite
    cmDlg.DialogTitle = "Importation du mindmap..."
    cmDlg.Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
    cmDlg.ShowSave
    Importer_DemanderNomFichier = cmDlg.Filename
    
    Exit Function
    
suite: 'Traitement des erreurs (bouton annuler !)
    Importer_DemanderNomFichier = ""
End Function 'Importer_DemanderNomFichier


'Forcer la position ou non
Private Sub mnuPopFrmMapForcerPos_Click()
    Arbre(NoeudSelectionne).PositionForcee = Not Arbre(NoeudSelectionne).PositionForcee
    DessinerAllMindMap
End Sub 'mnuPopFrmMapForcerPos

'Insérer un fils
Private Sub mnuPopFrmMapInsFils_Click()
    frmMDI.mnuNoeudInsererFils_Click
End Sub 'mnuPopFrmMapInsFils_Click


'Supprimer le noeud sélectionné
Private Sub mnuPopFrmMapSupFils_Click()
    SupprimerNoeud (NoeudSelectionne)
End Sub 'mnuPopFrmMapSupFils_Click


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Afficher l'aide contextuelle ?
    If frmMDI.MousePointer = 99 And Button.index <> 9 Then
        cmDlgHlp.HelpFile = App.Path & "\GMindmap.hlp"
        cmDlgHlp.HelpCommand = cdlHelpContext
        Select Case Button.index
            Case 1 To 4, 6 To 7: cmDlgHlp.HelpContext = Button.index + 99
            Case Else: cmDlgHlp.HelpContext = 9999
        End Select
        cmDlgHlp.ShowHelp
    Else 'Traîter les clicks
        Select Case Button.index
            Case 1: 'nouveau
                    mnuFichierNouveau_Click
            Case 2: 'Ouvrir
                    mnuFichierOuvrir_Click
            Case 3: 'Enregistrer
                    mnuFichierEnregistrer_Click
            Case 4: 'Imprimer
                    mnuFichierImprimer_Click
            Case 6: 'Insérer un fils
                    mnuNoeudInsererFils_Click
            Case 7: 'Supprimer un noeud
                    mnuNoeudSupprimer_Click
            Case 9: 'Aide contextuelle
                    mnuAideContextuelle_Click
        End Select
    End If
End Sub 'Toolbar1_ButtonClick
