Attribute VB_Name = "Module2"
' Module level variables
Dim grille(1 To 20, 1 To 20) As Integer
Dim reglesRobots(1 To 400) As Integer
Dim scoresRobots(1 To 400) As Long
Dim scoresParRegle(0 To 10) As Long
Dim decisionLunatique As Integer
Dim scoreBadBoy(1 To 400) As Integer
Const x As Integer = 20
Const y As Integer = 20
Dim rancuneRobots(1 To 400) As Integer
Dim derniereInteraction(1 To 400, 1 To 400) As Integer
Dim compteurNonCooperation(1 To 400) As Integer
Dim historiqueInteractions(400, 400, 9) As Integer
Dim coopere As Integer
Dim defection As Integer
Dim randomIndex As Integer









Sub Initialiser()
    Randomize
    ' Déclaration des variables pour les indices de boucle
    Dim i As Integer, j As Integer, robotID As Integer
    
    ' Initialisation des règles aléatoires pour chaque robot
    For i = 1 To 400
        reglesRobots(i) = Int((11 * Rnd))
        scoresRobots(i) = 0 ' Initialiser les scores à 0
    Next i
    
    For i = 1 To 400
        reglesRobots(i) = Int((11 * Rnd))
        scoresRobots(i) = 0
        scoreBadBoy(i) = 0 ' Initialisation du score de badboy à 0 pour chaque robot
        rancuneRobots(i) = 0
        compteurNonCooperation(i) = 0
    Next i

    
    ' Placement des robots dans la grille et coloration en fonction des règles
    robotID = 1
    For i = 1 To 20
        For j = 1 To 20
            grille(i, j) = robotID
            Select Case reglesRobots(robotID)
                Case 0
                    Cells(i, j).Interior.Color = vbRed
                Case 1
                    Cells(i, j).Interior.Color = vbYellow
                Case 2
                    Cells(i, j).Interior.Color = vbCyan
                Case 3
                    Cells(i, j).Interior.Color = vbGreen
                Case 4
                    Cells(i, j).Interior.Color = vbMagenta
                Case 5
                    Cells(i, j).Interior.Color = vbBlue
                Case 6
                    Cells(i, j).Interior.Color = vbBlack
                Case 7
                    Cells(i, j).Interior.Color = RGB(255, 105, 180) ' Hot Pink, using the RGB function
                Case 8
                    Cells(i, j).Interior.Color = RGB(255, 165, 0)   ' Orange, using the RGB function
                Case 9
                    Cells(i, j).Interior.Color = RGB(128, 0, 128)   ' Purple, using the RGB function
                Case 10
                    Cells(i, j).Interior.Color = RGB(64, 224, 208)  ' Turquoise, using the RGB function
                
                    
            End Select
            robotID = robotID + 1
        Next j
    Next i
    
    For i = 0 To 10
    scoresParRegle(i) = 0
    Next i
    

    For i = 1 To 400
        For j = 1 To 400
            derniereInteraction(i, j) = -1 ' -1 signifie qu'aucune interaction n'a encore eu lieu
        Next j
    Next i
    
    
    
    
    ' Assuming 400 robots and 10 possible historical interactions
    For i = 0 To 400
        For j = 0 To 400
            For k = 0 To 9
                historiqueInteractions(i, j, k) = 2 ' Initialize to 2 indicating no prior interaction
            Next k
        Next j
    Next i


    coopere = 0
    defection = 0
    
    
    



End Sub

Sub SimulerInteractions()
    Dim robot1 As Integer, robot2 As Integer
    Dim i As Integer, j As Integer
    
    ' Réaliser 40 itérations d'interaction
    For i = 1 To 40
        For j = 1 To 400
            robot1 = j
            robot2 = Int((400 * Rnd) + 1)
            
            ' Assurez-vous que robot1 et robot2 sont différents
            While robot1 = robot2
                robot2 = Int((400 * Rnd) + 1)
            Wend
            
            ' Appeler la fonction qui gère l'interaction et la mise à jour des scores
            GererInteraction robot1, robot2
        Next j
    Next i
    
    ' Analyser les résultats après toutes les itérations
    AnalyserResultats
End Sub

Function DistanceEntreRobots(robot1 As Integer, robot2 As Integer) As Integer
    Dim positionRobot1 As Variant, positionRobot2 As Variant
    positionRobot1 = TrouverPosition(robot1)
    positionRobot2 = TrouverPosition(robot2)
    
    ' Calcul de la distance de Manhattan entre deux points sur la grille
    DistanceEntreRobots = Abs(positionRobot1(0) - positionRobot2(0)) + Abs(positionRobot1(1) - positionRobot2(1))
End Function
Function TrouverPosition(robot As Integer) As Variant
    Dim i As Integer, j As Integer
    For i = 1 To 20
        For j = 1 To 20
            If grille(i, j) = robot Then
                TrouverPosition = Array(i, j)
                Exit Function
            End If
        Next j
    Next i
End Function





Sub GererInteraction(robot1 As Integer, robot2 As Integer)
    

    Const REGLE_EGOISTE As Integer = 0
    Const REGLE_GENEREUX As Integer = 1
    Const REGLE_FAMILIAL As Integer = 2
    Const REGLE_LUNATIQUE As Integer = 3
    Const REGLE_SECTE As Integer = 4
    Const REGLE_PSYCHOTIQUE As Integer = 5
    Const REGLE_REPUTATION As Integer = 6
    Const REGLE_DONNANT_DONNANT As Integer = 7
    Const REGLE_PARDON As Integer = 8
    Const REGLE_ELEPHANT As Integer = 9
    Const REGLE_ELEPHANT_LUNATIQUE As Integer = 10
    
    
    
    ' Initialisation des variables
    Dim pointsGagnesOuPerdusRobot1 As Integer
    Dim pointsGagnesOuPerdusRobot2 As Integer
    Dim distance As Integer

    distance = DistanceEntreRobots(robot1, robot2)
    
    ' Réinitialiser les points gagnés ou perdus à chaque interaction
    pointsGagnesOuPerdusRobot1 = 0
    pointsGagnesOuPerdusRobot2 = 0
    
    ' Application des règles pour robot1
    Select Case reglesRobots(robot1)
    
    
    '===========================EGOISTE=======================================
        Case REGLE_EGOISTE
            ' L'égoïste ne gagne des points que si l'autre est généreux
            If reglesRobots(robot2) = REGLE_GENEREUX Then
                pointsGagnesOuPerdusRobot1 = 4
                pointsGagnesOuPerdusRobot2 = -1
                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                scoreBadBoy(robot2) = scoreBadBoy(robot2)
            End If
            
            If reglesRobots(robot2) = REGLE_EGOISTE Then
                pointsGagnesOuPerdusRobot1 = 0
                pointsGagnesOuPerdusRobot2 = 0
                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
            End If
            
            If reglesRobots(robot2) = REGLE_FAMILIAL Then
                If distance <= 2 Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                Else
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                Dim decisionLunatique As Integer
                decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                
                ' L'égoïste gagne des points seulement si le lunatique coopère.
                If decisionLunatique = 3 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                Else
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1 'AJOUTER DANS LES AUTREEES
                End If
            End If
                
            If reglesRobots(robot2) = REGLE_SECTE Then
                pointsGagnesOuPerdusRobot1 = 0
                pointsGagnesOuPerdusRobot2 = 0
                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
            End If
            
            If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                If rancuneRobots(robot2) > x Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                Else
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                End If
            End If
                    
            
            If reglesRobots(robot2) = REGLE_REPUTATION Then
                If scoreBadBoy(robot2) <= y Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                Else
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                If derniereInteraction(robot1, robot2) = -1 Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf derniereInteraction(robot1, robot2) = 0 Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                ElseIf derniereInteraction(robot1, robot2) = 1 Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_PARDON Then
                If derniereInteraction(robot1, robot2) = -1 Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                ElseIf derniereInteraction(robot1, robot2) = 1 Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
            End If
            '////////////////////////////////////////////////////////////// TO CONTINUE
            If reglesRobots(robot2) = REGLE_ELEPHANT Then
                For i = 0 To 9
                    If historiqueInteractions(robot1, robot2, i) = 1 Then
                        coopere = coopere + 1
                    ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                        defection = defection + 1
                    End If
                Next
                If coopere > defection Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf defection > coopere Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                
                Randomize
                randomIndex = Int((10 * Rnd()))
            
                
                If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                    ' If the selected interaction was cooperation, then cooperate
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                    ' If the selected interaction was defection, then defect
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
            End If



  
            
            
            
            ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
                        
                        
                        
                        
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2
            
            
            
            
 

            
            

            
            
         
             
        '===========================================GENEREUX==========================================='
            
        Case REGLE_GENEREUX
            If reglesRobots(robot2) = REGLE_GENEREUX Then
                pointsGagnesOuPerdusRobot1 = 3
                pointsGagnesOuPerdusRobot2 = 3
                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                scoreBadBoy(robot2) = scoreBadBoy(robot2)
            End If
            
            If reglesRobots(robot2) = REGLE_EGOISTE Then
                pointsGagnesOuPerdusRobot1 = -1
                pointsGagnesOuPerdusRobot2 = 4
                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
            End If
            If reglesRobots(robot2) = REGLE_FAMILIAL Then
                If distance <= 2 Then pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                Else: pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
            End If
             
            If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                
                decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                
                ' L'égoïste gagne des points seulement si le lunatique coopère.
                If decisionLunatique = 3 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                Else
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
            End If
           
            If reglesRobots(robot2) = REGLE_SECTE Then
                pointsGagnesOuPerdusRobot1 = -1
                pointsGagnesOuPerdusRobot2 = 4
                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
            End If
            
            If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                If rancuneRobots(robot2) > x Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    rancuneRobots(robot2) = rancuneRobots(robot2)
                Else
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    rancuneRobots(robot2) = rancuneRobots(robot2)
                End If
            End If
                    
            
            If reglesRobots(robot2) = REGLE_REPUTATION Then
                If scoreBadBoy(robot2) <= y Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                Else
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                If derniereInteraction(robot1, robot2) = -1 Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf derniereInteraction(robot1, robot2) = 0 Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                ElseIf derniereInteraction(robot1, robot2) = 1 Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_PARDON Then
                If derniereInteraction(robot1, robot2) = -1 Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                ElseIf derniereInteraction(robot1, robot2) = 1 Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_ELEPHANT Then
                For i = 0 To 9
                    If historiqueInteractions(robot1, robot2, i) = 1 Then
                        coopere = coopere + 1
                    ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                        defection = defection + 1
                    End If
                Next
                If coopere > defection Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf defection > coopere Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                
                Randomize
                randomIndex = Int((10 * Rnd()))
            
                
                If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                    ' If the selected interaction was cooperation, then cooperate
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                    ' If the selected interaction was defection, then defect
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_ELEPHANT Then
                For i = 0 To 9
                    If historiqueInteractions(robot1, robot2, i) = 1 Then
                        coopere = coopere + 1
                    ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                        defection = defection + 1
                    End If
                Next
                If coopere > defection Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf defection > coopere Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        
                End If
            End If
            
            If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                
                Randomize
                randomIndex = Int((10 * Rnd()))
            
                
                If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                    ' If the selected interaction was cooperation, then cooperate
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                    ' If the selected interaction was defection, then defect
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
            End If
            
            ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
                        
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2
      
                
                
                
'========================================FAMILIAL============================================='
            
        Case REGLE_FAMILIAL
            ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
            
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2
                
            
            
            '++++++++++++++++++++DISTANCE+++++++++++++++++++++++++
            If distance <= 2 Then
                    If reglesRobots(robot2) = REGLE_GENEREUX Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                    
                    If reglesRobots(robot2) = REGLE_EGOISTE Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                    If reglesRobots(robot2) = REGLE_FAMILIAL Then
                        If distance <= 2 Then pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                        Else: pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
                     
                    If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                        ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                        
                        decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                        
                        ' L'égoïste gagne des points seulement si le lunatique coopère.
                        If decisionLunatique = 3 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                        Else
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    If reglesRobots(robot2) = REGLE_SECTE Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        
                    If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                        If rancuneRobots(robot2) > x Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            rancuneRobots(robot2) = rancuneRobots(robot2)
                        Else
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            rancuneRobots(robot2) = rancuneRobots(robot2)
                        End If
                    End If
                            
                    
                    If reglesRobots(robot2) = REGLE_REPUTATION Then
                        If scoreBadBoy(robot2) <= y Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        Else
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_PARDON Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    
                    End If
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT Then
                        For i = 0 To 9
                            If historiqueInteractions(robot1, robot2, i) = 1 Then
                                coopere = coopere + 1
                            ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                defection = defection + 1
                            End If
                        Next
                        If coopere > defection Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf defection > coopere Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                        
                        Randomize
                        randomIndex = Int((10 * Rnd()))
                    
                        
                        If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                            ' If the selected interaction was cooperation, then cooperate
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                            ' If the selected interaction was defection, then defect
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
            End If
        
                    
                    '+++++++++++++++++++++++++++DISTANCE+++++++++++++++++++++++++++
            If distance > 2 Then

                If reglesRobots(robot2) = REGLE_GENEREUX Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
                
                If reglesRobots(robot2) = REGLE_EGOISTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_FAMILIAL Then
                    If distance <= 2 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
            
                
                If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                    ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
              
                    decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                    
                    ' L'égoïste gagne des points seulement si le lunatique coopère.
                    If decisionLunatique = 3 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                    
                If reglesRobots(robot2) = REGLE_SECTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                        If rancuneRobots(robot2) > x Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                        Else
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                            
                        End If
                    End If
                            
                    
                    If reglesRobots(robot2) = REGLE_REPUTATION Then
                        If scoreBadBoy(robot2) <= y Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        Else
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_PARDON Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    If reglesRobots(robot2) = REGLE_ELEPHANT Then
                        For i = 0 To 9
                            If historiqueInteractions(robot1, robot2, i) = 1 Then
                                coopere = coopere + 1
                            ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                defection = defection + 1
                            End If
                        Next
                        If coopere > defection Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf defection > coopere Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                        
                        Randomize
                        randomIndex = Int((10 * Rnd()))
                    
                        
                        If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                            ' If the selected interaction was cooperation, then cooperate
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                            ' If the selected interaction was defection, then defect
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
                    

                End If
            End If
            '=============================================LUNATIQUE========================
         
        Case REGLE_LUNATIQUE
             ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
            
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2
            
            decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
            '+++++++++++++++++++++++++++++++++++++++++++++LUNATIQUE COOPERE+++++++++++++++++++++++++++++++++++++
            If decisionLunatique = 3 Then
                    If reglesRobots(robot2) = REGLE_GENEREUX Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                    
                    If reglesRobots(robot2) = REGLE_EGOISTE Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                    End If
                    If reglesRobots(robot2) = REGLE_FAMILIAL Then
                        If distance <= 2 Then pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        Else: pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
                     
                    If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                        ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                        
                        decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                        
                        ' L'égoïste gagne des points seulement si le lunatique coopère.
                        If decisionLunatique = 3 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        Else
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                

                        End If
                        
                    If reglesRobots(robot2) = REGLE_SECTE Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        
                    End If
                    
                    If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                        If rancuneRobots(robot2) > x Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            rancuneRobots(robot2) = rancuneRobots(robot2)
                        Else
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            rancuneRobots(robot2) = rancuneRobots(robot2)
                            
                        End If
                    End If
                            
                    
                    If reglesRobots(robot2) = REGLE_REPUTATION Then
                        If scoreBadBoy(robot2) <= y Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        Else
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
                    
                    
                    If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    If reglesRobots(robot2) = REGLE_PARDON Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
            
                    If reglesRobots(robot2) = REGLE_ELEPHANT Then
                        For i = 0 To 9
                            If historiqueInteractions(robot1, robot2, i) = 1 Then
                                coopere = coopere + 1
                            ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                defection = defection + 1
                            End If
                        Next
                        If coopere > defection Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf defection > coopere Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                        
                        Randomize
                        randomIndex = Int((10 * Rnd()))
                    
                        
                        If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                            ' If the selected interaction was cooperation, then cooperate
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                            ' If the selected interaction was defection, then defect
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
            
                    
            End If
        '+++++++++++++++++++++++++++++++++++++++LUNATIQUE COOPERE PAS++++++++++++++++++++++++++++++++++++++++++++
            If decisionLunatique = 0 Then
                If reglesRobots(robot2) = REGLE_GENEREUX Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
                
                If reglesRobots(robot2) = REGLE_EGOISTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_FAMILIAL Then
                    If distance <= 2 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
            
                
                If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                    ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
              
                    decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                    
                    ' L'égoïste gagne des points seulement si le lunatique coopère.
                    If decisionLunatique = 3 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                    
                If reglesRobots(robot2) = REGLE_SECTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                        If rancuneRobots(robot2) > x Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                        Else
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                            
                        End If
                    End If
                            
                    
                    If reglesRobots(robot2) = REGLE_REPUTATION Then
                        If scoreBadBoy(robot2) <= y Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        Else
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    If reglesRobots(robot2) = REGLE_PARDON Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT Then
                        For i = 0 To 9
                            If historiqueInteractions(robot1, robot2, i) = 1 Then
                                coopere = coopere + 1
                            ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                defection = defection + 1
                            End If
                        Next
                        If coopere > defection Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf defection > coopere Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                        
                        Randomize
                        randomIndex = Int((10 * Rnd()))
                    
                        
                        If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                            ' If the selected interaction was cooperation, then cooperate
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                            ' If the selected interaction was defection, then defect
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If

            End If
            '=========================================SECTE=================================================
        Case REGLE_SECTE
        
            ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
            
            
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2

            If reglesRobots(robot2) = REGLE_SECTE Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
            End If
                
            If reglesRobots(robot2) = REGLE_GENEREUX Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
            End If
                
            If reglesRobots(robot2) = REGLE_EGOISTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
            End If
                
            If reglesRobots(robot2) = REGLE_FAMILIAL Then
                If distance <= 2 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                Else
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
            End If
                
            If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                    ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                
                decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                
                    
                    ' L'égoïste gagne des points seulement si le lunatique coopère.
                If decisionLunatique = 3 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                Else
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                        If rancuneRobots(robot2) > x Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                        Else
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                            
                        End If
                    End If
                            
                    
                    If reglesRobots(robot2) = REGLE_REPUTATION Then
                        If scoreBadBoy(robot2) <= y Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        Else
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    If reglesRobots(robot2) = REGLE_PARDON Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT Then
                        For i = 0 To 9
                            If historiqueInteractions(robot1, robot2, i) = 1 Then
                                coopere = coopere + 1
                            ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                defection = defection + 1
                            End If
                        Next
                        If coopere > defection Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf defection > coopere Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                        
                        Randomize
                        randomIndex = Int((10 * Rnd()))
                    
                        
                        If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                            ' If the selected interaction was cooperation, then cooperate
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                            ' If the selected interaction was defection, then defect
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If

            End If
                    '=====================================================PSYCHOTIQUE===============================================================
        Case REGLE_PSYCHOTIQUE
        
            ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
            
            
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2

       ' ++++++++++++++++++++++++++++++++++++++++++++++++++++ PSYCHOTIQUE COOPERE PAS++++++++++++++++++++++++++++++++++++++++++++++++++++
            If rancuneRobots(robot1) > x Then  'la rancune est supérieure à x donc il va toujours faire défection
                If reglesRobots(robot2) = REGLE_EGOISTE Then  'ils font tous les deux défection
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                    
                ElseIf reglesRobots(robot1) = REGLE_GENEREUX Then  'l'un triche et pas l'autre
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    
                ElseIf reglesRobots(robot2) = REGLE_FAMILIAL Then  'regarder la distance entre les deux pour savoir le comportement de familial
                    If distance <= 2 Then
                       pointsGagnesOuPerdusRobot1 = 4
                       pointsGagnesOuPerdusRobot2 = -1
                       scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                       scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else: pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                    End If
                    
                ElseIf reglesRobots(robot2) = REGLE_LUNATIQUE Then
                    decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                
                    
                    ' L'égoïste gagne des points seulement si le lunatique coopère.
                    If decisionLunatique = 3 Then
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                                pointsGagnesOuPerdusRobot1 = 0
                                pointsGagnesOuPerdusRobot2 = 0
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                                
                    End If
                ElseIf reglesRobots(robot2) = REGLE_SECTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                    
                ElseIf reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                    If rancuneRobots(robot2) > x Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                        rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                
                    Else
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                        
                    End If
                    
                    
            
                End If
                If reglesRobots(robot2) = REGLE_REPUTATION Then
                        If scoreBadBoy(robot2) <= y Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            
                            
                        Else
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                        End If
                End If
                
                If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    If reglesRobots(robot2) = REGLE_PARDON Then
                        If derniereInteraction(robot1, robot2) = -1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        ElseIf derniereInteraction(robot1, robot2) = 1 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                    End If
                    
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT Then
                        For i = 0 To 9
                            If historiqueInteractions(robot1, robot2, i) = 1 Then
                                coopere = coopere + 1
                            ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                defection = defection + 1
                            End If
                        Next
                        If coopere > defection Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf defection > coopere Then
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                
                        End If
                    End If
                    
                    If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                        
                        Randomize
                        randomIndex = Int((10 * Rnd()))
                    
                        
                        If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                            ' If the selected interaction was cooperation, then cooperate
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                            ' If the selected interaction was defection, then defect
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                    End If

                
            '++++++++++++++++++++++++++++++++++++++PSYCHOTIQUE COOPERE ++++++++++++++++++++++++++++++++++++++
                
            ElseIf rancuneRobots(robot1) <= x Then
                If reglesRobots(robot2) = REGLE_EGOISTE Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                    
                ElseIf reglesRobots(robot1) = REGLE_GENEREUX Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    
                    
                ElseIf reglesRobots(robot2) = REGLE_FAMILIAL Then  'regarder la distance entre les deux pour savoir le comportement de familial
                    If distance <= 2 Then
                      pointsGagnesOuPerdusRobot1 = 3
                      pointsGagnesOuPerdusRobot2 = 3
                    Else: pointsGagnesOuPerdusRobot1 = -1
                       pointsGagnesOuPerdusRobot2 = 4
                       scoreBadBoy(robot1) = scoreBadBoy(robot1)
                       scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                       rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                    End If
                    
                ElseIf reglesRobots(robot2) = REGLE_LUNATIQUE Then
                    decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                
                    
                    ' L'égoïste gagne des points seulement si le lunatique coopère.
                    If decisionLunatique = 3 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                    Else
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                                
                    End If
                ElseIf reglesRobots(robot2) = REGLE_SECTE Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                 
                ElseIf reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                    If rancuneRobots(robot2) > x Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                        
                
                    Else
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        
                        
                    End If
                
                End If
                If reglesRobots(robot2) = REGLE_REPUTATION Then
                        If scoreBadBoy(robot2) <= y Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            
                        Else
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                        End If
                End If
                
                If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                If derniereInteraction(robot1, robot2) = -1 Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                ElseIf derniereInteraction(robot1, robot2) = 0 Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    rancuneRobots(robot1) = rancuneRobots(robot1) + 1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                ElseIf derniereInteraction(robot1, robot2) = 1 Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
                
                If reglesRobots(robot2) = REGLE_PARDON Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT Then
                    For i = 0 To 9
                        If historiqueInteractions(robot1, robot2, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    If coopere > defection Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf defection > coopere Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                    
                    Randomize
                    randomIndex = Int((10 * Rnd()))
                
                    
                    If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                        ' If the selected interaction was cooperation, then cooperate
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                        ' If the selected interaction was defection, then defect
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
            
            
                
            End If

                
            End If
'===========================================================DONNANT=====================================
         
         Case REGLE_DONNANT_DONNANT
            ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
            
            
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2
         '   ++++++++++++++++++++++++++++++ DONNANT COOPERE++++++++++++++++++++++++++++++++++++
            If derniereInteraction(robot2, robot1) = -1 Or derniereInteraction(robot2, robot1) = 1 Then
                 If reglesRobots(robot2) = REGLE_GENEREUX Then
                     pointsGagnesOuPerdusRobot1 = 3
                     pointsGagnesOuPerdusRobot2 = 3
                     scoreBadBoy(robot1) = scoreBadBoy(robot1)
                     scoreBadBoy(robot2) = scoreBadBoy(robot2)
                 End If
                 
                 If reglesRobots(robot2) = REGLE_EGOISTE Then
                     pointsGagnesOuPerdusRobot1 = -1
                     pointsGagnesOuPerdusRobot2 = 4
                     scoreBadBoy(robot1) = scoreBadBoy(robot1)
                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                 If reglesRobots(robot2) = REGLE_FAMILIAL Then
                     If distance <= 2 Then pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else: pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                  
                 If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                     ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                     
                     decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                     
                     ' L'égoïste gagne des points seulement si le lunatique coopère.
                     If decisionLunatique = 3 Then
                             pointsGagnesOuPerdusRobot1 = 3
                             pointsGagnesOuPerdusRobot2 = 3
                             scoreBadBoy(robot1) = scoreBadBoy(robot1)
                             scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else
                             pointsGagnesOuPerdusRobot1 = -1
                             pointsGagnesOuPerdusRobot2 = 4
                             scoreBadBoy(robot1) = scoreBadBoy(robot1)
                             scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     End If
                 End If
                
                 If reglesRobots(robot2) = REGLE_SECTE Then
                     pointsGagnesOuPerdusRobot1 = -1
                     pointsGagnesOuPerdusRobot2 = 4
                     scoreBadBoy(robot1) = scoreBadBoy(robot1)
                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                 
                 If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                     If rancuneRobots(robot2) > x Then
                         pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                         rancuneRobots(robot2) = rancuneRobots(robot2)
                     Else
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                         rancuneRobots(robot2) = rancuneRobots(robot2)
                     End If
                 End If
                         
                 
                 If reglesRobots(robot2) = REGLE_REPUTATION Then
                     If scoreBadBoy(robot2) <= y Then
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else
                         pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     End If
                 End If
                 
                 If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                     If derniereInteraction(robot1, robot2) = -1 Then
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     ElseIf derniereInteraction(robot1, robot2) = 0 Then
                         pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     ElseIf derniereInteraction(robot1, robot2) = 1 Then
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     End If
                 End If
                 If reglesRobots(robot2) = REGLE_PARDON Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                 End If
                 
                 If reglesRobots(robot2) = REGLE_ELEPHANT Then
                    For i = 0 To 9
                        If historiqueInteractions(robot1, robot2, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    If coopere > defection Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf defection > coopere Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                    End If
                 End If
                
                 If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                    
                    Randomize
                    randomIndex = Int((10 * Rnd()))
                
                    
                    If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                        ' If the selected interaction was cooperation, then cooperate
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                        ' If the selected interaction was defection, then defect
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                 End If
                
            
         '+++++++++++++++++++++++++++++++++++++++++ DONNANT COOPERE PAS ++++++++++++++++++++++++++++++++
            ElseIf derniereInteraction(robot2, robot1) = 0 Then
                If reglesRobots(robot2) = REGLE_GENEREUX Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
                
                If reglesRobots(robot2) = REGLE_EGOISTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_FAMILIAL Then
                    If distance <= 2 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                    ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                    
                    decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                    
                    ' L'égoïste gagne des points seulement si le lunatique coopère.
                    If decisionLunatique = 3 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1 'AJOUTER DANS LES AUTREEES
                    End If
                End If
                    
                If reglesRobots(robot2) = REGLE_SECTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                    If rancuneRobots(robot2) > x Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                    Else
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                    End If
                End If
                        
                
                If reglesRobots(robot2) = REGLE_REPUTATION Then
                    If scoreBadBoy(robot2) <= y Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_PARDON Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT Then
                    For i = 0 To 9
                        If historiqueInteractions(robot1, robot2, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    If coopere > defection Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf defection > coopere Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                    
                    Randomize
                    randomIndex = Int((10 * Rnd()))
                
                    
                    If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                        ' If the selected interaction was cooperation, then cooperate
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                        ' If the selected interaction was defection, then defect
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
            
            
           
            End If
        '=====================================================REPUTATION===========================================
         ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
            
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2
            '+++++++++++++++++++++++++++++ REPUTATION COOPERE PAS+++++++++++++++++++++++++++++++++
            If scoreBadBoy(robot2) > y Then
                 If reglesRobots(robot2) = REGLE_GENEREUX Then
                     pointsGagnesOuPerdusRobot1 = 4
                     pointsGagnesOuPerdusRobot2 = -1
                     scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                     scoreBadBoy(robot2) = scoreBadBoy(robot2)
                 End If
                 
                 If reglesRobots(robot2) = REGLE_EGOISTE Then
                     pointsGagnesOuPerdusRobot1 = 0
                     pointsGagnesOuPerdusRobot2 = 0
                     scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                 If reglesRobots(robot2) = REGLE_FAMILIAL Then
                     If distance <= 2 Then pointsGagnesOuPerdusRobot1 = 4
                         pointsGagnesOuPerdusRobot2 = -1
                         scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else: pointsGagnesOuPerdusRobot1 = 0
                         pointsGagnesOuPerdusRobot2 = 0
                         scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    
                 End If
                  
                 If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                     ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                     
                     decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                     
                     ' L'égoïste gagne des points seulement si le lunatique coopère.
                     If decisionLunatique = 3 Then
                             pointsGagnesOuPerdusRobot1 = 4
                             pointsGagnesOuPerdusRobot2 = -1
                             scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                             scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else
                             pointsGagnesOuPerdusRobot1 = 0
                             pointsGagnesOuPerdusRobot2 = 0
                             scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                             scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     End If
                 End If
                
                 If reglesRobots(robot2) = REGLE_SECTE Then
                     pointsGagnesOuPerdusRobot1 = 0
                     pointsGagnesOuPerdusRobot2 = 0
                     scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                 
                 If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                     If rancuneRobots(robot2) > x Then
                         pointsGagnesOuPerdusRobot1 = 0
                         pointsGagnesOuPerdusRobot2 = 0
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                         rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                     Else
                         pointsGagnesOuPerdusRobot1 = 4
                         pointsGagnesOuPerdusRobot2 = -1
                         scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                         rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                     End If
                 End If
                         
                 
                 If reglesRobots(robot2) = REGLE_REPUTATION Then
                     If scoreBadBoy(robot2) <= y Then
                         pointsGagnesOuPerdusRobot1 = 4
                         pointsGagnesOuPerdusRobot2 = -1
                         scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else
                         pointsGagnesOuPerdusRobot1 = 0
                         pointsGagnesOuPerdusRobot2 = 0
                         scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     End If
                 End If
                 
                 If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                     If derniereInteraction(robot1, robot2) = -1 Then
                         pointsGagnesOuPerdusRobot1 = 4
                         pointsGagnesOuPerdusRobot2 = -1
                         scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     ElseIf derniereInteraction(robot1, robot2) = 0 Then
                         pointsGagnesOuPerdusRobot1 = 0
                         pointsGagnesOuPerdusRobot2 = 0
                         scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     ElseIf derniereInteraction(robot1, robot2) = 1 Then
                         pointsGagnesOuPerdusRobot1 = 4
                         pointsGagnesOuPerdusRobot2 = -1
                         scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     End If
                 End If
                 
                If reglesRobots(robot2) = REGLE_PARDON Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT Then
                    For i = 0 To 9
                        If historiqueInteractions(robot1, robot2, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    If coopere > defection Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf defection > coopere Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                    
                    Randomize
                    randomIndex = Int((10 * Rnd()))
                
                    
                    If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                        ' If the selected interaction was cooperation, then cooperate
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                        ' If the selected interaction was defection, then defect
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                '++++++++++++++++++++++++++++++++++++REPUTATION COOPERE+++++++++++++++++++++++++++++++++++++
            ElseIf scoreBadBoy(robot2) <= y Then
                If reglesRobots(robot2) = REGLE_GENEREUX Then
                    pointsGagnesOuPerdusRobot1 = 3
                    pointsGagnesOuPerdusRobot2 = 3
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
                
                If reglesRobots(robot2) = REGLE_EGOISTE Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_FAMILIAL Then
                    If distance <= 2 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                    ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                    
                    decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                    
                    ' L'égoïste gagne des points seulement si le lunatique coopère.
                    If decisionLunatique = 3 Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1 'AJOUTER DANS LES AUTREEES
                    End If
                End If
                    
                If reglesRobots(robot2) = REGLE_SECTE Then
                    pointsGagnesOuPerdusRobot1 = -1
                    pointsGagnesOuPerdusRobot2 = 4
                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                    If rancuneRobots(robot2) > x Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        
                    Else
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        
                    End If
                End If
                        
                
                If reglesRobots(robot2) = REGLE_REPUTATION Then
                    If scoreBadBoy(robot2) <= y Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        
                    Else
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                
                
                If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                End If
                If reglesRobots(robot2) = REGLE_PARDON Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                End If
                If reglesRobots(robot2) = REGLE_ELEPHANT Then
                    For i = 0 To 9
                        If historiqueInteractions(robot1, robot2, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    If coopere > defection Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf defection > coopere Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                    
                    Randomize
                    randomIndex = Int((10 * Rnd()))
                
                    
                    If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                        ' If the selected interaction was cooperation, then cooperate
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                        ' If the selected interaction was defection, then defect
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
            
            
            
            
            
           
            End If
            
'===========================================================PARDON=====================================
            
         Case REGLE_PARDON
            ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
            
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2
         '   ++++++++++++++++++++++++++++++ PARDON COOPERE++++++++++++++++++++++++++++++++++++
            If derniereInteraction(robot2, robot1) = -1 Or derniereInteraction(robot2, robot1) = 1 Then
                 If reglesRobots(robot2) = REGLE_GENEREUX Then
                     pointsGagnesOuPerdusRobot1 = 3
                     pointsGagnesOuPerdusRobot2 = 3
                     scoreBadBoy(robot1) = scoreBadBoy(robot1)
                     scoreBadBoy(robot2) = scoreBadBoy(robot2)
                 End If
                 
                 If reglesRobots(robot2) = REGLE_EGOISTE Then
                     pointsGagnesOuPerdusRobot1 = -1
                     pointsGagnesOuPerdusRobot2 = 4
                     scoreBadBoy(robot1) = scoreBadBoy(robot1)
                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                 If reglesRobots(robot2) = REGLE_FAMILIAL Then
                     If distance <= 2 Then pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else: pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                  
                 If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                     ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                     
                     decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                     
                     ' L'égoïste gagne des points seulement si le lunatique coopère.
                     If decisionLunatique = 3 Then
                             pointsGagnesOuPerdusRobot1 = 3
                             pointsGagnesOuPerdusRobot2 = 3
                             scoreBadBoy(robot1) = scoreBadBoy(robot1)
                             scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else
                             pointsGagnesOuPerdusRobot1 = -1
                             pointsGagnesOuPerdusRobot2 = 4
                             scoreBadBoy(robot1) = scoreBadBoy(robot1)
                             scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     End If
                 End If
                
                 If reglesRobots(robot2) = REGLE_SECTE Then
                     pointsGagnesOuPerdusRobot1 = -1
                     pointsGagnesOuPerdusRobot2 = 4
                     scoreBadBoy(robot1) = scoreBadBoy(robot1)
                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                 
                If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                     If rancuneRobots(robot2) > x Then
                         pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                         rancuneRobots(robot2) = rancuneRobots(robot2)
                     Else
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                         rancuneRobots(robot2) = rancuneRobots(robot2)
                     End If
                End If
                         
                 
                If reglesRobots(robot2) = REGLE_REPUTATION Then
                     If scoreBadBoy(robot2) <= y Then
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else
                         pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     End If
                End If
                 
                If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                     If derniereInteraction(robot1, robot2) = -1 Then
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     ElseIf derniereInteraction(robot1, robot2) = 0 Then
                         pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     ElseIf derniereInteraction(robot1, robot2) = 1 Then
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     End If
                End If
                If reglesRobots(robot2) = REGLE_PARDON Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT Then
                    For i = 0 To 9
                        If historiqueInteractions(robot1, robot2, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    If coopere > defection Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf defection > coopere Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                    
                    Randomize
                    randomIndex = Int((10 * Rnd()))
                
                    
                    If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                        ' If the selected interaction was cooperation, then cooperate
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                        ' If the selected interaction was defection, then defect
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
            End If
         '   ++++++++++++++++++++++++++++++ PARDON COOPERE++++++++++++++++++++++++++++++++++++
            If derniereInteraction(robot2, robot1) = 0 And compteurNonCooperation(robot2) <= 1 Then
                 If reglesRobots(robot2) = REGLE_GENEREUX Then
                     pointsGagnesOuPerdusRobot1 = 3
                     pointsGagnesOuPerdusRobot2 = 3
                     scoreBadBoy(robot1) = scoreBadBoy(robot1)
                     scoreBadBoy(robot2) = scoreBadBoy(robot2)
                 End If
                 
                 If reglesRobots(robot2) = REGLE_EGOISTE Then
                     pointsGagnesOuPerdusRobot1 = -1
                     pointsGagnesOuPerdusRobot2 = 4
                     scoreBadBoy(robot1) = scoreBadBoy(robot1)
                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                 If reglesRobots(robot2) = REGLE_FAMILIAL Then
                     If distance <= 2 Then pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else: pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                  
                 If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                     ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                     
                     decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                     
                     ' L'égoïste gagne des points seulement si le lunatique coopère.
                     If decisionLunatique = 3 Then
                             pointsGagnesOuPerdusRobot1 = 3
                             pointsGagnesOuPerdusRobot2 = 3
                             scoreBadBoy(robot1) = scoreBadBoy(robot1)
                             scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else
                             pointsGagnesOuPerdusRobot1 = -1
                             pointsGagnesOuPerdusRobot2 = 4
                             scoreBadBoy(robot1) = scoreBadBoy(robot1)
                             scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     End If
                 End If
                
                 If reglesRobots(robot2) = REGLE_SECTE Then
                     pointsGagnesOuPerdusRobot1 = -1
                     pointsGagnesOuPerdusRobot2 = 4
                     scoreBadBoy(robot1) = scoreBadBoy(robot1)
                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                 End If
                 
                 If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                     If rancuneRobots(robot2) > x Then
                         pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                         rancuneRobots(robot2) = rancuneRobots(robot2)
                     Else
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                         rancuneRobots(robot2) = rancuneRobots(robot2)
                     End If
                 End If
                         
                 
                 If reglesRobots(robot2) = REGLE_REPUTATION Then
                     If scoreBadBoy(robot2) <= y Then
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     Else
                         pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     End If
                 End If
                 
                 If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                     If derniereInteraction(robot1, robot2) = -1 Then
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     ElseIf derniereInteraction(robot1, robot2) = 0 Then
                         pointsGagnesOuPerdusRobot1 = -1
                         pointsGagnesOuPerdusRobot2 = 4
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                     ElseIf derniereInteraction(robot1, robot2) = 1 Then
                         pointsGagnesOuPerdusRobot1 = 3
                         pointsGagnesOuPerdusRobot2 = 3
                         scoreBadBoy(robot1) = scoreBadBoy(robot1)
                         scoreBadBoy(robot2) = scoreBadBoy(robot2)
                     End If
                 End If
                 If reglesRobots(robot2) = REGLE_PARDON Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot2) <= 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot2) > 1 Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                 End If
                 If reglesRobots(robot2) = REGLE_ELEPHANT Then
                    For i = 0 To 9
                        If historiqueInteractions(robot1, robot2, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    If coopere > defection Then
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf defection > coopere Then
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                    End If
                 End If
                
                 If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                    
                    Randomize
                    randomIndex = Int((10 * Rnd()))
                
                    
                    If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                        ' If the selected interaction was cooperation, then cooperate
                        pointsGagnesOuPerdusRobot1 = 3
                        pointsGagnesOuPerdusRobot2 = 3
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                        ' If the selected interaction was defection, then defect
                        pointsGagnesOuPerdusRobot1 = -1
                        pointsGagnesOuPerdusRobot2 = 4
                        scoreBadBoy(robot1) = scoreBadBoy(robot1)
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                 End If
         '+++++++++++++++++++++++++++++++++++++++++ Pardon COOPERE PAS ++++++++++++++++++++++++++++++++
            ElseIf derniereInteraction(robot2, robot1) = 0 And compteurNonCooperation(robot2) > 1 Then
                If reglesRobots(robot2) = REGLE_GENEREUX Then
                    pointsGagnesOuPerdusRobot1 = 4
                    pointsGagnesOuPerdusRobot2 = -1
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                End If
                
                If reglesRobots(robot2) = REGLE_EGOISTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_FAMILIAL Then
                    If distance <= 2 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                    ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                    
                    decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                    
                    ' L'égoïste gagne des points seulement si le lunatique coopère.
                    If decisionLunatique = 3 Then
                            pointsGagnesOuPerdusRobot1 = 4
                            pointsGagnesOuPerdusRobot2 = -1
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                            pointsGagnesOuPerdusRobot1 = 0
                            pointsGagnesOuPerdusRobot2 = 0
                            scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1 'AJOUTER DANS LES AUTREEES
                    End If
                End If
                    
                If reglesRobots(robot2) = REGLE_SECTE Then
                    pointsGagnesOuPerdusRobot1 = 0
                    pointsGagnesOuPerdusRobot2 = 0
                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                End If
                
                If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                    If rancuneRobots(robot2) > x Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                    Else
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                    End If
                End If
                        
                
                If reglesRobots(robot2) = REGLE_REPUTATION Then
                    If scoreBadBoy(robot2) <= y Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    Else
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_PARDON Then
                    If derniereInteraction(robot1, robot2) = -1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf derniereInteraction(robot1, robot2) = 1 Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT Then
                    For i = 0 To 9
                        If historiqueInteractions(robot1, robot2, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    If coopere > defection Then
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf defection > coopere Then
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                    End If
                End If
                
                If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                    
                    Randomize
                    randomIndex = Int((10 * Rnd()))
                
                    
                    If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                        ' If the selected interaction was cooperation, then cooperate
                        pointsGagnesOuPerdusRobot1 = 4
                        pointsGagnesOuPerdusRobot2 = -1
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2)
                    ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                        ' If the selected interaction was defection, then defect
                        pointsGagnesOuPerdusRobot1 = 0
                        pointsGagnesOuPerdusRobot2 = 0
                        scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                        scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                    End If
                End If
            
            
           
            End If
            
            
        '=====================================================ELEPHANT===========================================
        Case REGLE_ELEPHANT
         
         ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
            
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2
            
            '+++++++++++++++++++++++++++++ ELEPHANT COOPERE PAS+++++++++++++++++++++++++++++++++
            If reglesRobots(robot1) = REGLE_ELEPHANT Then '///////////////////////////
                    For i = 0 To 9
                        If historiqueInteractions(robot2, robot1, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot2, robot1, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    
                    If defection > coopere Then
        
        
                         If reglesRobots(robot2) = REGLE_GENEREUX Then
                             pointsGagnesOuPerdusRobot1 = 4
                             pointsGagnesOuPerdusRobot2 = -1
                             scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                             scoreBadBoy(robot2) = scoreBadBoy(robot2)
                         End If
                         
                         If reglesRobots(robot2) = REGLE_EGOISTE Then
                             pointsGagnesOuPerdusRobot1 = 0
                             pointsGagnesOuPerdusRobot2 = 0
                             scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                             scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                         End If
                         If reglesRobots(robot2) = REGLE_FAMILIAL Then
                             If distance <= 2 Then pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             Else: pointsGagnesOuPerdusRobot1 = 0
                                 pointsGagnesOuPerdusRobot2 = 0
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                         End If
                          
                         If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                             ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                             
                             decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                             
                             ' L'égoïste gagne des points seulement si le lunatique coopère.
                             If decisionLunatique = 3 Then
                                     pointsGagnesOuPerdusRobot1 = 4
                                     pointsGagnesOuPerdusRobot2 = -1
                                     scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                     scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             Else
                                     pointsGagnesOuPerdusRobot1 = 0
                                     pointsGagnesOuPerdusRobot2 = 0
                                     scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                             End If
                         End If
                        
                         If reglesRobots(robot2) = REGLE_SECTE Then
                             pointsGagnesOuPerdusRobot1 = 0
                             pointsGagnesOuPerdusRobot2 = 0
                             scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                             scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                         End If
                         
                         If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                             If rancuneRobots(robot2) > x Then
                                 pointsGagnesOuPerdusRobot1 = 0
                                 pointsGagnesOuPerdusRobot2 = 0
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                 rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                             Else
                                 pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                                 rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                             End If
                         End If
                                 
                         
                         If reglesRobots(robot2) = REGLE_REPUTATION Then
                             If scoreBadBoy(robot2) <= y Then
                                 pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             Else
                                 pointsGagnesOuPerdusRobot1 = 0
                                 pointsGagnesOuPerdusRobot2 = 0
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                             End If
                         End If
                         
                         If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                             If derniereInteraction(robot1, robot2) = -1 Then
                                 pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             ElseIf derniereInteraction(robot1, robot2) = 0 Then
                                 pointsGagnesOuPerdusRobot1 = 0
                                 pointsGagnesOuPerdusRobot2 = 0
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                             ElseIf derniereInteraction(robot1, robot2) = 1 Then
                                 pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             End If
                         End If
                         
                        If reglesRobots(robot2) = REGLE_PARDON Then
                            If derniereInteraction(robot1, robot2) = -1 Then
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                                pointsGagnesOuPerdusRobot1 = 0
                                pointsGagnesOuPerdusRobot2 = 0
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            ElseIf derniereInteraction(robot1, robot2) = 1 Then
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            End If
                        End If
                        
                        If reglesRobots(robot2) = REGLE_ELEPHANT Then
                            For i = 0 To 9
                                If historiqueInteractions(robot1, robot2, i) = 1 Then
                                    coopere = coopere + 1
                                ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                    defection = defection + 1
                                End If
                            Next
                            If coopere > defection Then
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf defection > coopere Then
                                pointsGagnesOuPerdusRobot1 = 0
                                pointsGagnesOuPerdusRobot2 = 0
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                    
                            End If
                        End If
                        
                        If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                            
                            Randomize
                            randomIndex = Int((10 * Rnd()))
                        
                            
                            If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                                ' If the selected interaction was cooperation, then cooperate
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                                ' If the selected interaction was defection, then defect
                                pointsGagnesOuPerdusRobot1 = 0
                                pointsGagnesOuPerdusRobot2 = 0
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            End If
                        End If
                        
                    End If
            End If
            
            
            
            
                '++++++++++++++++++++++++++++++++++++ELEPHANT COOPERE+++++++++++++++++++++++++++++++++++++
            If reglesRobots(robot1) = REGLE_ELEPHANT Then '///////////////////////////
                    For i = 0 To 9
                        If historiqueInteractions(robot2, robot1, i) = 1 Then
                            coopere = coopere + 1
                        ElseIf historiqueInteractions(robot2, robot1, i) = 0 Then
                            defection = defection + 1
                        End If
                    Next
                    If coopere > defection Then
        
                        If reglesRobots(robot2) = REGLE_GENEREUX Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                        
                        If reglesRobots(robot2) = REGLE_EGOISTE Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                        
                        If reglesRobots(robot2) = REGLE_FAMILIAL Then
                            If distance <= 2 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            Else
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            End If
                        End If
                        
                        If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                            ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                            
                            decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                            
                            ' L'égoïste gagne des points seulement si le lunatique coopère.
                            If decisionLunatique = 3 Then
                                    pointsGagnesOuPerdusRobot1 = 3
                                    pointsGagnesOuPerdusRobot2 = 3
                                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            Else
                                    pointsGagnesOuPerdusRobot1 = -1
                                    pointsGagnesOuPerdusRobot2 = 4
                                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1 'AJOUTER DANS LES AUTREEES
                            End If
                        End If
                            
                        If reglesRobots(robot2) = REGLE_SECTE Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                        
                        If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                            If rancuneRobots(robot2) > x Then
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                
                            Else
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                
                            End If
                        End If
                                
                        
                        If reglesRobots(robot2) = REGLE_REPUTATION Then
                            If scoreBadBoy(robot2) <= y Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                
                            Else
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            End If
                        End If
                        
                        
                        If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                            If derniereInteraction(robot1, robot2) = -1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 Then
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            ElseIf derniereInteraction(robot1, robot2) = 1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            End If
                        End If
                        If reglesRobots(robot2) = REGLE_PARDON Then
                            If derniereInteraction(robot1, robot2) = -1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            ElseIf derniereInteraction(robot1, robot2) = 1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            End If
                        End If
                        If reglesRobots(robot2) = REGLE_ELEPHANT Then
                            For i = 0 To 9
                                If historiqueInteractions(robot1, robot2, i) = 1 Then
                                    coopere = coopere + 1
                                ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                    defection = defection + 1
                                End If
                            Next
                            If coopere > defection Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf defection > coopere Then
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                    
                            End If
                        End If
                        
                        If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                            
                            Randomize
                            randomIndex = Int((10 * Rnd()))
                        
                            
                            If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                                ' If the selected interaction was cooperation, then cooperate
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                                ' If the selected interaction was defection, then defect
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            End If
                        End If
                    End If
                    
            End If
            
            
        '=====================================================ELEPHANT LUNATIQUE===========================================
        Case REGLE_ELEPHANT_LUNATIQUE
         
         ''''!!!!!Cooperation our pas pour donnant donnant
              ' À la fin de chaque case, déterminez si l'interaction était coopérative
        ' et mettez à jour derniereInteraction en conséquence
            If pointsGagnesOuPerdusRobot1 = 3 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 1 ' Assurez-vous de réciproquer pour la cohérence
            ElseIf pointsGagnesOuPerdusRobot1 = -1 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 1 ' Coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0
                derniereInteraction(robot2, robot1) = 1
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
            ElseIf pointsGagnesOuPerdusRobot1 = 0 And reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                derniereInteraction(robot1, robot2) = 0 ' Non-coopération
                derniereInteraction(robot2, robot1) = 0
                compteurNonCooperation(robot1) = compteurNonCooperation(robot1) + 1
                compteurNonCooperation(robot2) = compteurNonCooperation(robot2) + 1
                
            End If
            
            If pointsGagnesOuPerdusRobot1 = 3 Or pointsGagnesOuPerdusRobot1 = -1 Then
                resultatInteractionRobot1 = 1
            ElseIf pointsGagnesOuPerdusRobot1 = 4 Or pointsGagnesOuPerdusRobot1 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            If pointsGagnesOuPerdusRobot2 = 3 Or pointsGagnesOuPerdusRobot2 = -1 Then
                resultatInteractionRobot2 = 1
            ElseIf pointsGagnesOuPerdusRobot2 = 4 Or pointsGagnesOuPerdusRobot2 = 0 Then
                resultatInteractionRobot1 = 0
            End If
            For i = 0 To 8 ' Décaler de 1 vers le début, perdant la plus ancienne interaction
                historiqueInteractions(robot1, robot2, i) = historiqueInteractions(robot1, robot2, i + 1)
                historiqueInteractions(robot2, robot1, i) = historiqueInteractions(robot2, robot1, i + 1)
            Next
            ' Ajouter la nouvelle interaction à la fin
            historiqueInteractions(robot1, robot2, 9) = resultatInteractionRobot1
            historiqueInteractions(robot2, robot1, 9) = resultatInteractionRobot2
            
            '+++++++++++++++++++++++++++++ ELEPHANT LUNATIQUE COOPERE PAS+++++++++++++++++++++++++++++++++
            If reglesRobots(robot1) = REGLE_ELEPHANT_LUNATIQUE Then '///////////////////////////
            

                    Randomize
                    randomIndex = Int((10 * Rnd()))
                    
                    If historiqueInteractions(robot2, robot1, randomIndex) = 0 Then
        
        
                         If reglesRobots(robot2) = REGLE_GENEREUX Then
                             pointsGagnesOuPerdusRobot1 = 4
                             pointsGagnesOuPerdusRobot2 = -1
                             scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                             scoreBadBoy(robot2) = scoreBadBoy(robot2)
                         End If
                         
                         If reglesRobots(robot2) = REGLE_EGOISTE Then
                             pointsGagnesOuPerdusRobot1 = 0
                             pointsGagnesOuPerdusRobot2 = 0
                             scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                             scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                         End If
                         If reglesRobots(robot2) = REGLE_FAMILIAL Then
                             If distance <= 2 Then pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             Else: pointsGagnesOuPerdusRobot1 = 0
                                 pointsGagnesOuPerdusRobot2 = 0
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            
                         End If
                          
                         If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                             ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                             
                             decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                             
                             ' L'égoïste gagne des points seulement si le lunatique coopère.
                             If decisionLunatique = 3 Then
                                     pointsGagnesOuPerdusRobot1 = 4
                                     pointsGagnesOuPerdusRobot2 = -1
                                     scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                     scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             Else
                                     pointsGagnesOuPerdusRobot1 = 0
                                     pointsGagnesOuPerdusRobot2 = 0
                                     scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                     scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                             End If
                         End If
                        
                         If reglesRobots(robot2) = REGLE_SECTE Then
                             pointsGagnesOuPerdusRobot1 = 0
                             pointsGagnesOuPerdusRobot2 = 0
                             scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                             scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                         End If
                         
                         If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                             If rancuneRobots(robot2) > x Then
                                 pointsGagnesOuPerdusRobot1 = 0
                                 pointsGagnesOuPerdusRobot2 = 0
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                 rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                             Else
                                 pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                                 rancuneRobots(robot2) = rancuneRobots(robot2) + 1
                             End If
                         End If
                                 
                         
                         If reglesRobots(robot2) = REGLE_REPUTATION Then
                             If scoreBadBoy(robot2) <= y Then
                                 pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             Else
                                 pointsGagnesOuPerdusRobot1 = 0
                                 pointsGagnesOuPerdusRobot2 = 0
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                             End If
                         End If
                         
                         If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                             If derniereInteraction(robot1, robot2) = -1 Then
                                 pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             ElseIf derniereInteraction(robot1, robot2) = 0 Then
                                 pointsGagnesOuPerdusRobot1 = 0
                                 pointsGagnesOuPerdusRobot2 = 0
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                             ElseIf derniereInteraction(robot1, robot2) = 1 Then
                                 pointsGagnesOuPerdusRobot1 = 4
                                 pointsGagnesOuPerdusRobot2 = -1
                                 scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                 scoreBadBoy(robot2) = scoreBadBoy(robot2)
                             End If
                         End If
                         
                        If reglesRobots(robot2) = REGLE_PARDON Then
                            If derniereInteraction(robot1, robot2) = -1 Then
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                                pointsGagnesOuPerdusRobot1 = 0
                                pointsGagnesOuPerdusRobot2 = 0
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            ElseIf derniereInteraction(robot1, robot2) = 1 Then
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            End If
                        End If
                        
                        If reglesRobots(robot2) = REGLE_ELEPHANT Then
                            For i = 0 To 9
                                If historiqueInteractions(robot1, robot2, i) = 1 Then
                                    coopere = coopere + 1
                                ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                    defection = defection + 1
                                End If
                            Next
                            If coopere > defection Then
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf defection > coopere Then
                                pointsGagnesOuPerdusRobot1 = 0
                                pointsGagnesOuPerdusRobot2 = 0
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                    
                            End If
                        End If
                        
                        If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                            
                            Randomize
                            randomIndex = Int((10 * Rnd()))
                        
                            
                            If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                                ' If the selected interaction was cooperation, then cooperate
                                pointsGagnesOuPerdusRobot1 = 4
                                pointsGagnesOuPerdusRobot2 = -1
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                                ' If the selected interaction was defection, then defect
                                pointsGagnesOuPerdusRobot1 = 0
                                pointsGagnesOuPerdusRobot2 = 0
                                scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            End If
                        End If
                        
                    End If
            End If
            
            
            
            
                '++++++++++++++++++++++++++++++++++++ELEPHANT COOPERE+++++++++++++++++++++++++++++++++++++
            If reglesRobots(robot1) = REGLE_ELEPHANT_LUNATIQUE Then '///////////////////////////
                    Randomize
                    randomIndex = Int((10 * Rnd()))
                    
                    If historiqueInteractions(robot2, robot1, randomIndex) = 1 Then
        
                        If reglesRobots(robot2) = REGLE_GENEREUX Then
                            pointsGagnesOuPerdusRobot1 = 3
                            pointsGagnesOuPerdusRobot2 = 3
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2)
                        End If
                        
                        If reglesRobots(robot2) = REGLE_EGOISTE Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                        
                        If reglesRobots(robot2) = REGLE_FAMILIAL Then
                            If distance <= 2 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            Else
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            End If
                        End If
                        
                        If reglesRobots(robot2) = REGLE_LUNATIQUE Then
                            ' Le robot lunatique choisit aléatoirement de coopérer ou trahir.
                            
                            decisionLunatique = IIf(Rnd() < 0.5, 3, 0) ' Coopère ou non, mais ne trahit pas activement.
                            
                            ' L'égoïste gagne des points seulement si le lunatique coopère.
                            If decisionLunatique = 3 Then
                                    pointsGagnesOuPerdusRobot1 = 3
                                    pointsGagnesOuPerdusRobot2 = 3
                                    scoreBadBoy(robot1) = scoreBadBoy(robot1) + 1
                                    scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            Else
                                    pointsGagnesOuPerdusRobot1 = -1
                                    pointsGagnesOuPerdusRobot2 = 4
                                    scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                    scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1 'AJOUTER DANS LES AUTREEES
                            End If
                        End If
                            
                        If reglesRobots(robot2) = REGLE_SECTE Then
                            pointsGagnesOuPerdusRobot1 = -1
                            pointsGagnesOuPerdusRobot2 = 4
                            scoreBadBoy(robot1) = scoreBadBoy(robot1)
                            scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                        End If
                        
                        If reglesRobots(robot2) = REGLE_PSYCHOTIQUE Then
                            If rancuneRobots(robot2) > x Then
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                
                            Else
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                
                            End If
                        End If
                                
                        
                        If reglesRobots(robot2) = REGLE_REPUTATION Then
                            If scoreBadBoy(robot2) <= y Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                
                            Else
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            End If
                        End If
                        
                        
                        If reglesRobots(robot2) = REGLE_DONNANT_DONNANT Then
                            If derniereInteraction(robot1, robot2) = -1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 Then
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            ElseIf derniereInteraction(robot1, robot2) = 1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            End If
                        End If
                        If reglesRobots(robot2) = REGLE_PARDON Then
                            If derniereInteraction(robot1, robot2) = -1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) <= 1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf derniereInteraction(robot1, robot2) = 0 And compteurNonCooperation(robot1) > 1 Then
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            ElseIf derniereInteraction(robot1, robot2) = 1 Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            End If
                        End If
                        If reglesRobots(robot2) = REGLE_ELEPHANT Then
                            For i = 0 To 9
                                If historiqueInteractions(robot1, robot2, i) = 1 Then
                                    coopere = coopere + 1
                                ElseIf historiqueInteractions(robot1, robot2, i) = 0 Then
                                    defection = defection + 1
                                End If
                            Next
                            If coopere > defection Then
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf defection > coopere Then
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                                    
                            End If
                        End If
                        
                        If reglesRobots(robot2) = REGLE_ELEPHANT_LUNATIQUE Then
                            
                            Randomize
                            randomIndex = Int((10 * Rnd()))
                        
                            
                            If historiqueInteractions(robot1, robot2, randomIndex) = 1 Then
                                ' If the selected interaction was cooperation, then cooperate
                                pointsGagnesOuPerdusRobot1 = 3
                                pointsGagnesOuPerdusRobot2 = 3
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2)
                            ElseIf historiqueInteractions(robot1, robot2, randomIndex) = 0 Then
                                ' If the selected interaction was defection, then defect
                                pointsGagnesOuPerdusRobot1 = -1
                                pointsGagnesOuPerdusRobot2 = 4
                                scoreBadBoy(robot1) = scoreBadBoy(robot1)
                                scoreBadBoy(robot2) = scoreBadBoy(robot2) + 1
                            End If
                        End If
                    End If
                    
            End If
            
            
        
    End Select
    

         


    
    ' Mise à jour des scores individuels des robots
    scoresRobots(robot1) = scoresRobots(robot1) + pointsGagnesOuPerdusRobot1
    scoresRobots(robot2) = scoresRobots(robot2) + pointsGagnesOuPerdusRobot2

    ' Mise à jour du score total pour la règle du robot
    scoresParRegle(reglesRobots(robot1)) = scoresParRegle(reglesRobots(robot1)) + pointsGagnesOuPerdusRobot1
    scoresParRegle(reglesRobots(robot2)) = scoresParRegle(reglesRobots(robot2)) + pointsGagnesOuPerdusRobot2
End Sub

   
Sub AnalyserResultats()
    Dim i As Integer
    Dim MaxScore As Long, MinScore As Long
    Dim meilleureRegle As Integer, pireRegle As Integer
    Dim Seuil As Long

    ' Initialiser MaxScore et MinScore pour s'assurer de capturer correctement les scores max et min
    MaxScore = scoresParRegle(0)
    MinScore = scoresParRegle(0)
   
    meilleureRegle = 0
    pireRegle = 0
    
    ' Trouver la règle avec le score maximum et minimum
    For i = 0 To UBound(scoresParRegle)
        If scoresParRegle(i) > MaxScore Then
            MaxScore = scoresParRegle(i)
            meilleureRegle = i
        End If
        If scoresParRegle(i) < MinScore Then
            MinScore = scoresParRegle(i)
            pireRegle = i
        End If
    Next i
        
        ' Initialiser MaxScore et MinScore pour s'assurer de capturer correctement les scores max et min
    MaxScoreRobots = scoresRobots(1)
    MinScoreRobots = scoresRobots(1)
       
    ' Trouver le score maximum et minimum obtenu par les robots
    For i = 1 To 400
        If scoresRobots(i) > MaxScoreRobots Then
            MaxScoreRobots = scoresRobots(i)
        End If
        If scoresRobots(i) < MinScoreRobots Then
            MinScoreRobots = scoresRobots(i)
        End If
    Next i

' Calcul du seuil pour le changement de règle
Seuil = MaxScoreRobots - ((MaxScoreRobots - MinScoreRobots) / 4)

MsgBox "Seuil: " & Seuil & vbCrLf & _
        "Score maximum des robots: " & MaxScore & vbCrLf & _
        "Score minimum des robots: " & MinScore


    ' Appliquer le changement de règle pour les robots dont le score est sous le seuil
    For i = 1 To 400
        If scoresRobots(i) < Seuil Then
            reglesRobots(i) = meilleureRegle
            ' Mise à jour des couleurs dans la grille selon la nouvelle règle
            Dim row As Integer, col As Integer
            row = (i - 1) \ 20 + 1
            col = (i - 1) Mod 20 + 1
            Select Case meilleureRegle
                Case 0
                    Cells(row, col).Interior.Color = vbRed
                Case 1
                    Cells(row, col).Interior.Color = vbYellow
                Case 2
                    Cells(row, col).Interior.Color = vbCyan
                Case 3
                    Cells(row, col).Interior.Color = vbGreen
                Case 4
                    Cells(row, col).Interior.Color = vbMagenta
                Case 5
                    Cells(row, col).Interior.Color = vbBlue
            End Select
        End If
    Next i

    ' Afficher les résultats dans une boîte de message
 MsgBox "La règle avec le meilleur score est: " & meilleureRegle & " avec un score de: " & MaxScore & vbCrLf & _
           "La règle avec le moins bon score est: " & pireRegle & " avec un score de: " & MinScore
    
    MsgBox "La règle avec le meilleur score est: " & meilleureRegle & " avec un score de: " & MaxScore & vbCrLf & _
           "La règle avec le moins bon score est: " & pireRegle & " avec un score de: " & MinScore
End Sub


