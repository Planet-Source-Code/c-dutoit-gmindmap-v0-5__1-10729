Attribute VB_Name = "modES"
'modES : Gestion des entrées - sorties
'Par C.Dutoit, 2 Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit


Dim Buffer As String   'Buffer de lecture du fichier

'Commencer un nouveau Mindmap
Sub NouveauFichier()
    'Créer la racine
    ReDim Arbre(0)
    Arbre(0).Legende = "SANS TITRE"
    Arbre(0).URL = ""
    Arbre(0).NbSuivants = 0
    
    'Sélectionner la racine
    NoeudSelectionne = 0
    
    'Mettre à jour l'affichage
    DessinerAllMindMap
    
    'Définir les paramètres du programme
    MyApp.Fichier = ""
    MyApp.Modifie = False
    SetAppCaption
End Sub 'Nouveau Fichier


'Format de fichier.gmm : (Texte) (exemple)
'Signature :   "GMM v1.1"
'Nb de noeuds  "113"
'puis pour chaque noeud :
'Legende , URL, PosX, PosY
'Décalage de n*4 caractères pour chaque niveau de l'arbre

'Sauvegarde d'un arbre par récursion
Private Sub SauverArbreRec(indice As Long, Indentation)
    'Sauver le noeud
    If Arbre(indice).PositionForcee Then
        Print #1, Space$(Indentation) & Arbre(indice).Legende & "," & Arbre(indice).URL & "," & _
            (Arbre(indice).x - Arbre(0).x) / frmMap.ScaleWidth * 1000 & "," & _
            (Arbre(indice).y - Arbre(0).y) / frmMap.ScaleHeight * 1000
    Else
        Print #1, Space$(Indentation) & Arbre(indice).Legende & "," & Arbre(indice).URL
    End If
    
    Dim i
    'Sauver les fils
    If Arbre(indice).NbSuivants > 0 Then
        'Sauver chaque fils
        For i = 0 To Arbre(indice).NbSuivants - 1
            SauverArbreRec Arbre(indice).Suivants(i), Indentation + 4
        Next i
    End If
End Sub 'SauverArbreRec


'Sauver un arbre
Sub SauverArbre(Filename As String)
    'Ouvrir le fichier
    Open Filename For Output Access Write As #1
    
    'Enregistrer la signature et la taille
    Print #1, "GMM v1.1"
    Print #1, UBound(Arbre)
    
    'Sauver l'arbre, récursivement
    SauverArbreRec 0, 0
    
    'Fermer le fichier
    Close #1
    
    'Ajouter à la liste des derniers fichiers ouverts
    AddToLastOppenedFiles (Filename)
    ShowLastOpenedFiles
    
    'Le fichier n'est pas modifié + enregistrement du nom du fichier +...
    MyApp.Fichier = Filename
    MyApp.Modifie = False
    SetAppCaption
End Sub 'SauverArbre


'Chargement d'un arbre, récursivement, TOREDO
Private Function ChargerArbreRec(Parent As Long, IndentationParent As Long)
  While Not EOF(1)
    'Lire l'élément
    If EOF(1) Then Exit Function
    If Buffer = "" Then Line Input #1, Buffer
    
    'Relever le nb d'indentations
    Dim NbIndent As Long
    NbIndent = NbIndentation(Buffer)
    
    
    
    'Chercher le bon endroit
    Select Case NbIndent - IndentationParent
        Case Is < 0: Exit Function 'niveau supérieur : je remonte
        Case Is = 0: Exit Function 'même niveau que le parent : je remonte d'un niveau
        Case 1: 'j 'insère (voir plus loin)
        Case Is > 1: MsgBox "Erreur lors du chargement de l'arbre.. " & vbCrLf & "Je tente cependant de continuer...", vbInformation, "Erreur au chargement..."
    End Select
    
    'Créer l'élément
    ReDim Preserve Arbre(UBound(Arbre) + 1)
    Arbre(UBound(Arbre)).Legende = GetLegende(Buffer)
    Arbre(UBound(Arbre)).URL = GetURL(Buffer)
    Arbre(UBound(Arbre)).NbSuivants = 0
    Arbre(UBound(Arbre)).x = Val(GetPosX(Buffer)) * frmMap.ScaleWidth / 1000
    Arbre(UBound(Arbre)).y = Val(GetPosY(Buffer)) * frmMap.ScaleHeight / 1000
    Arbre(UBound(Arbre)).PositionForcee = (Arbre(UBound(Arbre)).x <> -1 And Arbre(UBound(Arbre)).y <> -1)
    
    'Enregistrer le lien dans le parent '(attention : l'indice 0 existe !)
    ReDim Preserve Arbre(Parent).Suivants(Arbre(Parent).NbSuivants)
    Arbre(Parent).NbSuivants = Arbre(Parent).NbSuivants + 1
    Arbre(Parent).Suivants(Arbre(Parent).NbSuivants - 1) = UBound(Arbre)
    Buffer = ""
    
    'Insérer les fils
    ChargerArbreRec UBound(Arbre), NbIndent
  Wend
End Function 'ChargerArbreRec


'Charger un arbre
Sub ChargerArbre(Filename As String)
    Dim TempStr As String
    
    'Ouvrir le fichier
    Open Filename For Input Access Read As #1
    
    'Vérifier le format
    Input #1, TempStr
    If Left$(TempStr, 6) <> "GMM v1" Then
        MsgBox "Erreur : le fichier spécifié n'est pas au bon format !", vbCritical, "Erreur au chargement..."
        Exit Sub
    End If
    
    'Lire la taille
    Line Input #1, TempStr
    ReDim Arbre(0) 'Val(TempStr))
    
    'Lire la racine
    Line Input #1, TempStr
    Arbre(0).Legende = GetLegende(TempStr)
    Arbre(0).URL = GetURL(TempStr)
    Arbre(0).PositionForcee = False
    
    'Lire l'arbre, récursivement
    Buffer = ""
    While Not EOF(1)
        ChargerArbreRec 0, 0
    Wend
    
    'Fermer le fichier
    Close #1
    
    'Ajouter à la liste des derniers fichiers ouverts
    AddToLastOppenedFiles (Filename)
    ShowLastOpenedFiles
    
    'Définir le titre de la fenêtre; dessiner le mindmap
    MyApp.Fichier = Filename
    MyApp.Modifie = False
    SetAppCaption
    DessinerAllMindMap
End Sub 'ChargerArbre


'Retourner la légende d'une ligne <légende, URL>
Private Function GetLegende(Chaine As String) As String
    Dim pos
    pos = InStr(Chaine, ",")
    
    If pos > 0 Then
        GetLegende = LTrim$(Left$(Chaine, InStr(Chaine, ",") - 1))
    Else
        GetLegende = LTrim$(Chaine)
    End If
End Function 'GetLegende


'Retourner l'URL d'une ligne <légende, URL>
Private Function GetURL(Chaine As String) As String
    Dim pos, pos2
    pos = InStr(Chaine, ",")
    
    If pos > 0 And pos < Len(Chaine) Then
        pos2 = InStr(Chaine, ",")
        If pos2 < 1 Then pos = Len(Chaine)
        If pos2 > pos Then
            GetURL = RTrim$(Mid$(Chaine, pos + 1, pos2 - pos - 1))
        Else
            GetURL = ""
        End If
    Else
        GetURL = ""
    End If
End Function 'GetLegende


'Retourner la position X d'un chaine <légende, URL, x, y>
Private Function GetPosX(Chaine As String) As String
    Dim pos, maChaine
    maChaine = Chaine
    
    'Passer la première virgule
    pos = InStr(maChaine, ",")
    If pos <= 0 Or pos = Len(maChaine) Then
        GetPosX = "-1"
        Exit Function
    End If
    
    'Passer la seconde (virgule)
    maChaine = Right$(maChaine, Len(maChaine) - pos - 1)
    pos = InStr(maChaine, ",")
    If pos <= 0 Then
        GetPosX = "-1"
        Exit Function
    End If
    
    pos = InStr(maChaine, ",")
    If pos < 1 Then pos = Len(maChaine)
    maChaine = Left$(maChaine, pos - 1)
    GetPosX = RTrim$(maChaine)
End Function 'GetPosX


'Retourner la position Y d'un chaine <légende, URL, x, y>
Private Function GetPosY(Chaine As String) As String
    Dim pos, maChaine
    maChaine = Chaine
    
    'Passer la première virgule
    pos = InStr(maChaine, ",")
    If pos <= 0 Then
        GetPosY = "-1"
        Exit Function
    End If
    
    'Passer la seconde (virgule)
    maChaine = Right$(Chaine, Len(maChaine) - pos)
    pos = InStr(maChaine, ",")
    If pos <= 0 Then
        GetPosY = "-1"
        Exit Function
    End If
    
    'Passer la troisième (virgule)
    maChaine = Right$(maChaine, Len(maChaine) - pos)
    pos = InStr(maChaine, ",")
    If pos <= 0 Then
        GetPosY = "-1"
        Exit Function
    End If
    
    maChaine = Right$(maChaine, Len(maChaine) - pos)
    GetPosY = RTrim$(maChaine)
End Function 'GetPosY


'Compter le nombre d'indentations de 4 présent au début d'une chaine
Private Function NbIndentation(Chaine As String) As Long
    Dim i
    For i = 1 To Len(Chaine)
        If Mid$(Chaine, i, 1) <> " " Then Exit For
    Next i
    
    NbIndentation = (i - 1) / 4
End Function 'NbIndentation


'Compter le nombre de tabulation présent au début d'une chaine
Private Function NbTab(Chaine As String) As Long
    Dim i
    For i = 1 To Len(Chaine)
        If Mid$(Chaine, i, 1) <> vbTab Then Exit For
    Next i
    
    NbTab = (i - 1)
End Function 'NbTab






'Format de fichier.gmm : (Texte) (exemple)
'Signature :   "GMM v1.0"
'Nb de noeuds  "113"
'puis pour chaque noeud :
'Legende , URL,  + si position forcée : x, y
'Décalage de n*4 caractères pour chaque niveau de l'arbre

'Sauvegarde d'un arbre par récursion
Private Sub ExporterTexteRec(indice As Long, Indentation As Long)
    Dim text As String, i As Long
    text = ""
    If Indentation > 0 Then For i = 1 To Indentation: text = text & vbTab: Next i
    
    text = text & Arbre(indice).Legende
    Print #1, text
    
    'Sauver les fils
    If Arbre(indice).NbSuivants > 0 Then
        'Sauver chaque fils
        For i = 0 To Arbre(indice).NbSuivants - 1
            ExporterTexteRec Arbre(indice).Suivants(i), Indentation + 1
        Next i
    End If
End Sub 'ExporterTexteRec


'Exporter un arbre au format texte
Sub ExporterTexte(Filename As String)
    'Ouvrir le fichier
    Open Filename For Output Access Write As #1
    
    'Sauver l'arbre, récursivement
    ExporterTexteRec 0, -1
    
    'Fermer le fichier
    Close #1
    
    MsgBox "Le fichier a été exporté !", vbInformation, "Exportation réussie !"
End Sub 'ExporterTexte






'Chargement d'un arbre, récursivement, TOREDO
Private Function ImporterArbreRec(Parent As Long, IndentationParent As Long)
  While Not EOF(1)
    'Lire l'élément
    If EOF(1) Then Exit Function
    If Buffer = "" Then Line Input #1, Buffer
    
    'Relever le nb d'indentations
    Dim NbIndent As Long
    NbIndent = NbTab(Buffer)
    
    
    
    'Chercher le bon endroit
    Select Case NbIndent - IndentationParent
        Case Is < 0: Exit Function 'niveau supérieur : je remonte
        Case Is = 0: Exit Function 'même niveau que le parent : je remonte d'un niveau
        Case 1: 'j 'insère (voir plus loin)
        Case Is > 1: MsgBox "Erreur lors du chargement de l'arbre.. " & vbCrLf & "Je tente cependant de continuer...", vbInformation, "Erreur au chargement..."
    End Select
    
    'Créer l'élément
    ReDim Preserve Arbre(UBound(Arbre) + 1)
    Arbre(UBound(Arbre)).Legende = GetLegende(Right$(Buffer, Len(Buffer) - NbTab(Buffer)))
    Arbre(UBound(Arbre)).URL = GetURL(Buffer)
    Arbre(UBound(Arbre)).NbSuivants = 0
    
    'Enregistrer le lien dans le parent '(attention : l'indice 0 existe !)
    ReDim Preserve Arbre(Parent).Suivants(Arbre(Parent).NbSuivants)
    Arbre(Parent).NbSuivants = Arbre(Parent).NbSuivants + 1
    Arbre(Parent).Suivants(Arbre(Parent).NbSuivants - 1) = UBound(Arbre)
    Buffer = ""
    
    'Insérer les fils
    ImporterArbreRec UBound(Arbre), NbIndent
  Wend
End Function 'ImporterArbreRec


'Charger un arbre
Sub ImporterArbre(Filename As String)
    Dim TempStr As String
    
    'Ouvrir le fichier
    Open Filename For Input Access Read As #1
    
    ReDim Arbre(0)
    
    'Lire la racine
    Line Input #1, TempStr
    Arbre(0).Legende = GetLegende(TempStr)
    Arbre(0).URL = GetURL(TempStr)
    
    'Lire l'arbre, récursivement
    Buffer = ""
    While Not EOF(1)
        ImporterArbreRec 0, -1
    Wend
    
    'Fermer le fichier
    Close #1
End Sub 'ImporterArbre
