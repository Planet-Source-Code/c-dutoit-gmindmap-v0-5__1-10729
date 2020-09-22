Attribute VB_Name = "modMain"
'modMain : Module principale, divers
'Par C.Dutoit, 2 Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit


Type TMyApp 'Données vitales pour l'application
    Fichier As String  'Nom de fichier actuel
    Modifie As Boolean 'Fichier modifié ?
End Type 'TMyApp

Global MyApp As TMyApp  'Données vitales de l'applications



'Définir le titre de la fenêtre principale
Sub SetAppCaption()
    Dim Caption As String
    Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & " "
    
    
    If MyApp.Fichier <> "" Then Caption = Caption & "[" & MyApp.Fichier & "] "
    If MyApp.Modifie Then Caption = Caption & "*"
    
    frmMDI.Caption = Caption
End Sub 'SetAppCaption


'Afficher la liste des derniers fichiers Ouverts 'Importé de Gipic v0.8
Sub ShowLastOpenedFiles()
  Dim Fichier As String
  Dim i As Integer
  
  'Afficher chaque fichier
  For i = 1 To 5
     Fichier = GetSetting(App.EXEName, "Options", "File" & i)
     
     If Fichier <> "" Then
        frmMDI.mnuFichierLast(i).Caption = "&" & i & " " & Fichier
     Else
        frmMDI.mnuFichierLast(i).Caption = "&" & i & " -"
     End If
  Next i
End Sub 'ShowLastOpenedFiles


'Ajouter un element en tete de la liste des derniers fichiers Ouverts ''Importé de Gipic v0.8
Sub AddToLastOppenedFiles(Filename As String)
  Dim Entrees(1 To 6) As String
  Dim i As Integer
  Dim j As Integer
  
  'Lire la liste des derniers fichiers ouverts
  For i = 1 To 5
     Entrees(i + 1) = GetSetting(App.EXEName, "Options", "File" & i)
  Next i
  
  'Ajouter le fichier Actuel
  Entrees(1) = Filename
  
  'Supprimer les doublons
  For i = 2 To 6
     If Entrees(i) = Filename Then
        For j = i + 1 To 6
          Entrees(j - 1) = Entrees(j)
        Next j
     End If
  Next i
  
  'Enregistrer la liste
  For i = 1 To 5
     Call SaveSetting(App.EXEName, "Options", "File" & i, Entrees(i))
  Next i
  
  'Afficher la liste
  ShowLastOpenedFiles
End Sub 'ShowLastOpenedFiles


'Vérifie si un fichier existe
Function FileExist(Filename As String) As Boolean
  On Error GoTo erreur
  Open Filename For Input As #1
  Close #1
  FileExist = True
  Exit Function
  
erreur:
  FileExist = False
End Function 'FileExist
