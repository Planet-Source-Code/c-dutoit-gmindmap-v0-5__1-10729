Attribute VB_Name = "modOperations"
'modOperations : Opérations diverses sur le Mindmap
'Par C.Dutoit, 2 Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit


'Créer un fils
Sub CreerFils(Parent As Long)
    'Redimensionner l'arbre (+1)
    ReDim Preserve Arbre(UBound(Arbre) + 1)
    
    'Créer le noeud
    Arbre(UBound(Arbre)).Legende = ""
    Arbre(UBound(Arbre)).NbSuivants = 0
    Arbre(UBound(Arbre)).URL = ""
    Arbre(UBound(Arbre)).PositionForcee = False
    Arbre(UBound(Arbre)).x = 0
    Arbre(UBound(Arbre)).y = 0
    
    'Ajouter le fils au parent
    If Arbre(Parent).NbSuivants = 0 Then
        Arbre(Parent).NbSuivants = 1
        ReDim Arbre(Parent).Suivants(0)
        Arbre(Parent).Suivants(0) = UBound(Arbre)
    Else
        ReDim Preserve Arbre(Parent).Suivants(UBound(Arbre(Parent).Suivants) + 1)
        Arbre(Parent).Suivants(UBound(Arbre(Parent).Suivants)) = UBound(Arbre)
        Arbre(Parent).NbSuivants = Arbre(Parent).NbSuivants + 1
    End If
End Sub 'CreerFils





'Supprimer un noeud
Sub SupprimerNoeud(index As Long)
    'Indice correct ?
    If index < 0 Or index > UBound(Arbre) Then
        MsgBox "Tentative de suppression d'un noeud inexistant", vbExclamation, "Erreur..."
        Exit Sub
    End If
    
    'Tentative de suppression de la racine ?
    If index = 0 Then
        MsgBox "impossible de supprimer le premier noeud !", vbExclamation, "Erreur..."
        Exit Sub
    End If
    
    'Supprimer de l'arbre
    Dim i, j
    For i = index + 1 To UBound(Arbre)
        Arbre(i - 1) = Arbre(i)
    Next i
    ReDim Preserve Arbre(UBound(Arbre) - 1)
    
    'Supprimer le lien depuis le parent
    Dim k
    Dim found As Boolean
    found = False
    For i = 0 To UBound(Arbre)
        If Arbre(i).NbSuivants > 0 Then
            For j = 0 To UBound(Arbre(i).Suivants)
                If Arbre(i).Suivants(j) = index Then 'Supprimer la référence
                    'Décaler les suivants
                    For k = j + 1 To UBound(Arbre(i).Suivants)
                        Arbre(i).Suivants(k - 1) = Arbre(i).Suivants(k)
                    Next k
                    
                    'Redimensionner l'arbre
                    If UBound(Arbre(i).Suivants) > 0 Then ReDim Preserve Arbre(i).Suivants(UBound(Arbre(i).Suivants) - 1)
                    Arbre(i).NbSuivants = Arbre(i).NbSuivants - 1
                    found = True
                End If
                If found Then Exit For
            Next j
        End If
        If found Then Exit For
    Next i
    
    'Déplacer les liens sur les indices supérieur à l'indice du noeud à supprimer
    For i = 0 To UBound(Arbre)
        If Arbre(i).NbSuivants > 0 Then
            For j = 0 To UBound(Arbre(i).Suivants)
                If Arbre(i).Suivants(j) > index Then Arbre(i).Suivants(j) = Arbre(i).Suivants(j) - 1
            Next j
        End If
    Next i
End Sub 'SupprimerNoeud




'Retourner le N° du noeud le plus proche. Dist max = largeur de "OOOO"
Function NoeudLePlusProcheXY(x As Long, y As Long) As Long
    Dim i As Long      'Variable de boucle
    Dim Dist As Long, DistTemp As Long 'Distance au point
    Dim Noeud As Long  'Noeud le plus proche
    
    'Initialisation
    Dist = frmMap.TextWidth("OOOO")
    Noeud = -1
    
    'Chercher le point le plus proche
    For i = 0 To UBound(Arbre)
        'Calculer la distance au point
        DistTemp = Sqr((Arbre(i).x - x) ^ 2 + (Arbre(i).y - y) ^ 2)
        
        'Distance plus petite ? => on enregistre le point et la distance
        If DistTemp < Dist Then
            Dist = DistTemp
            Noeud = i
        End If
    Next i
    
    'Retourner le noeud le plus proche
    NoeudLePlusProcheXY = Noeud
End Function 'NoeudLePlusProche
