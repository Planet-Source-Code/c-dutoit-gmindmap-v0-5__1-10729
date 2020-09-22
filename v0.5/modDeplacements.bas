Attribute VB_Name = "modDeplacements"
'modDeplacements : Gestion des déplacements
'Crée par C.Dutoit pour GMindMap v0.5 le 13 Août 2000
Option Explicit

Const pi = 3.1415926535

Enum TDirection
    gauche
    droite
    haut
    bas
End Enum


'Chercher le noeud le plus proche vers la droite
Function NoeudLePlusProche(x As Long, y As Long, Direction As TDirection) As Long
    'Initialisations
    Dim i
    Dim Dist As Single, DistTemp As Single
    Dim Noeud As Long
    Noeud = -1
    Dim Angle As Single
    Const Tolerance = 0.05
    
    
    'Recalculer les positions
    CalculerCoordonnees
    
    'Chercher le point le plus proche
    For i = 0 To UBound(Arbre)
        If Arbre(i).x <> x Or Arbre(i).y <> y Then
            Angle = Atn((Arbre(i).y - y) / (Arbre(i).x - x + 0.000001))
            If Arbre(i).x - x < 0 Then Angle = Angle + pi
            If Angle < 0 Then While Angle < 0: Angle = Angle + 2 * pi: Wend
            If Angle > 2 * pi Then While Angle > 2 * pi: Angle = Angle - 2 * pi: Wend
            
            
            'Bon angle ?
            If (Direction = droite And (Angle >= 7 * pi / 4 - Tolerance Or Angle <= pi / 4 + Tolerance)) Or _
               (Direction = gauche And (Angle >= 3 * pi / 4 - Tolerance And Angle <= 5 * pi / 4 + Tolerance)) Or _
               (Direction = haut And (Angle >= pi / 4 - Tolerance And Angle <= 3 * pi / 4 + Tolerance)) Or _
               (Direction = bas And (Angle >= 5 * pi / 4 - Tolerance And Angle <= 7 * pi / 4 + Tolerance)) Then
                'Calculer la distance
                If (Direction = droite Or Direction = gauche) Then
                    DistTemp = Sqr((Arbre(i).x - x) ^ 2 + (Arbre(i).y - y) ^ 2)
                Else
                    DistTemp = Sqr((Arbre(i).x - x) ^ 2 + (Arbre(i).y - y) ^ 2)
                End If
                
                'Meilleure distance que la précédente ?
                If Noeud = -1 Or Dist > DistTemp Then
                    Dist = DistTemp
                    Noeud = i
                End If
            End If
        End If
    Next i
    
    'Retourner le point
    NoeudLePlusProche = Noeud
End Function 'noeudLePlusProche


Sub SelectionnerLeNoeudADroite(x As Long, y As Long)
    Dim Noeud As Long
    Noeud = NoeudLePlusProche(x, y, droite)
    
    If Noeud > -1 Then
        NoeudSelectionne = Noeud
        DessinerAllMindMap
    End If
End Sub 'SelectionnerLeNoeudADroite


Sub SelectionnerLeNoeudAGauche(x As Long, y As Long)
    Dim Noeud As Long
    Noeud = NoeudLePlusProche(x, y, gauche)
    
    If Noeud > -1 Then
        NoeudSelectionne = Noeud
        DessinerAllMindMap
    End If
End Sub 'SelectionnerLeNoeudAGauche


Sub SelectionnerLeNoeudEnHaut(x As Long, y As Long)
    Dim Noeud As Long
    Noeud = NoeudLePlusProche(x, y, bas) 'le graphe se dessine à l'envers !
    
    If Noeud > -1 Then
        NoeudSelectionne = Noeud
        DessinerAllMindMap
    End If
End Sub 'SelectionnerLeNoeudEnHaut


Sub SelectionnerLeNoeudEnBas(x As Long, y As Long)
    Dim Noeud As Long
    Noeud = NoeudLePlusProche(x, y, haut) 'le graphe se dessine à l'envers !
    
    If Noeud > -1 Then
        NoeudSelectionne = Noeud
        DessinerAllMindMap
    End If
End Sub 'SelectionnerLeNoeudEnBas
