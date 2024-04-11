' Proprietary Notice:
'
' This software was developed by Witold Gesing (wgesing@gmail.com)
' This software may be copied and re-distributed freely.
'
' If you make any changes, please re-name the functions affected and clearly identify
' and document any changes, improvements, modifications or additions.

Option Compare Text 'Uppercase letters to be equivalent to lowercase letters.
Option Base 1       'Set the default lower bound for array subscripts to 1.

Function csScoreVer()
        csScoreVer = "Cox Sprague Scoring Calculator Version 1.1, July 25, 2022, Witold Gesing"
        '1.1 Cox Sprague Table (CoxSpragueTable) added to the csScore() call sequence to avoid #value! error on loading.
End Function
              
' Function csScore() returns Cox-Sprague percentage score given boat's race results,
' the number of starters in each of these races and the number of discards.
' This version uses either the original or the modified (default) Cox-Sprague table.

' Cox-Sprague percentage score for a race series is the total number of points
' the boat earned in the series divided by the total number of points it could
' have earned by placing first in every race it started.
'
' Inputs:

' resultsVector:  a vector of race results:
'                 numbers are scored as finishing places
'                 DSQ/DNF/OCS/BFD/RET/RAF are scored as number of starters + 1
'                 DNE or DGM are scored as number of starters + 1 and are not excludable
'                 TLE is scored as number of finishers + 2 but not worse than number of starters + 1
'                 DNS, DNC and blank or zero entries are ignored
'
' nStartersVector:  an integer vector containing number of starters in each race
' CoxSpragueTable:            100 by 20 array containing Cox-Sprague table
' nFinishersVector:  an integer vector containing number of finishers in each race (used for scoring TLE only)
' nDiscards:   number of discards (defaults to 0).
'

'Details:

' In each race the number of starters will determine the column to be used in the Cox-Sprague table,
' and each boat will be credited with the number of points indicated for her finishing place.
' A boat's series score shall be her "Percentage of Perfection" calculated by dividing her total points
' scored by the total points she would have had, had she won every race in which she started. A boat
' which does not finish or is disqualified in a race shall receive a score for the place one greater
' than the number of starters in that race.

' If the modified Cox-Sprague table is used recommended), there are two differences
' from the "original" Cox-Sprague Scoring system:

'1) The score for boats that finish worse than 21st is computed using a formula which results in scores
'   ranging from 57.4% for the 22nd place to 53.5% for the 200th place.
'   This is more in line with the rest of the CS table which assigns scores for the last place
'   finish ranging 70% to 59% in races with 20 or fewer boats than the scores assigned
'   by the "original" version in which the 21st place gets a score of 58% and subsequent scores decrease
'   by 1% per place. For large fleets this would result in negative scores for boats finishing 80'th or higher.
'   This change only affects races with more than 21 participants.

'2) The score for a boat that finishes second in a two boat race has been changed from 40% to 70%.
'   This is more consistent with the 67.7% score assigned to a boat that finishes third in a three boat
'   race and a 59% score assigned to the boat that finishes last in a 20 boat race.
'   This change only affects races with exactly 2 participants.
            
' Witold Gesing, July 23, 2022

Function csScore(resultsVector, _
               nStartersVector, _
               CoxSpragueTable, _
               Optional nFinishersVector, _
               Optional nDiscards As Integer = 0) As Double
               
    'The difference between this function and csScore is that this function takes the Cox-Sprague Table (CoxSpragueTable) as input
    ' as opposed to calling function csTable1 (below).
    ' Funtion csTable1() can be used to compute the desired Cox-Sprague table.

    Dim r As Integer ' row index to csTable
    Dim c As Integer ' column index to csTable
    Dim nx As Integer 'counter of number of races started by the boat being scored
    Dim nsv As Integer ' length of nStartersVector

    Dim xTot As Double 'running C-S score total
    Dim pTot As Double 'running C-S score "perfect" total
    Dim cs As Double 'cs percentage
    Dim z As Double  'temp variable
    Dim a As Double  'fractional portion of finishing place (usually a=0)
    
    Dim i As Integer, j As Integer 'loop indexes
    Dim im As Integer ' index of the race to be discarded
    Dim ns As Variant ' For Each loop index - number of starters in each race
    Dim lbl As String ' temp variable for non-numeric elements of resultsVector
    
    Dim xv() As Double 'vector C-S scores in races started (for computing discards)
    Dim pv() As Double 'vector of perfect C-S scores (for computing discards)
            
    nsv = nStartersVector.Count
    
    ReDim xv(nsv) As Double
    ReDim pv(nsv) As Double
           
    xTot = 0 ' Initialize running C-S total
    pTot = 0 ' Initialize running C-S "perfect" total
    nx = 0 ' Initialize counter of races started
    
    i = 0
    For Each ns In nStartersVector
      r = 0
      lbl = ""
      i = i + 1
          If IsNumeric(ns) Then
              If ns > 0 Then
                c = ns
                  If c > 20 Then c = 20
                      
                  If IsNumeric(resultsVector(i)) Then
                     r = Int(Abs(resultsVector(i)))
                     a = Abs(resultsVector(i)) - r
                  Else
                     lbl = Left(Trim(resultsVector(i)), 3)
                     
                     ' DSQ/DNF/DNE/DGM/OCS/BFD/UFD/RET/RAF scored as ns + 1
                     If Left(lbl, 1) = "D" Or Left(lbl, 1) = "R" Then r = ns + 1
                     If lbl = "OCS" Or lbl = "BFD" Or lbl = "UFD" Then r = ns + 1
                                         
                     'Time Limit Expired scored as number of finishers + 2
                     If lbl = "TLE" Then
                         If IsMissing(nFinishersVector) Then r = ns + 1
                         If nFinishersVector(i) > 0 Then r = nFinishersVector(i) + 2
                         If (r > ns + 1) Then r = ns + 1
                     End If 'bl = "TLE"
                     
                    'Do not count DNS/DNC
                     If lbl = "DNS" Or lbl = "DNC" Then r = 0
                  End If 'IsNumeric(resultsVector(i))
                  
                'If r > 0 And c > 0 Then z = CoxSpragueTable(r, c)
                
                If r > 0 And c > 0 Then
                    'If r < 22 Then
                    z = CoxSpragueTable(r, c)
                    'If r > 21 And modified = True Then z = 58 - 2 * Log(r + a - 20) / Log(10)
                    'If r > 21 And modified = False Then z = 79 - r - a
                                        
                    'If finishing place is not an integer, interpolate between CS scores
                    If a > 0 And r < c Then z = z + a * (CoxSpragueTable(r + 1, c) - CoxSpragueTable(r, c))
                    
                    nx = nx + 1 'counter of races started
                    xv(nx) = z  'vector of C-S scores (for computing discards)
                    xTot = xTot + z ' running C-S score total
                    
                    ' Disqualification not Excludable (setting xv(nx)=0 prevents result of race nx from being discarded)
                    If lbl = "DNE" Or lbl = "DGM" Then xv(nx) = 0
                    
                    ' Perfect score and vector of perfect CS scores for computing discards
                    pv(nx) = CoxSpragueTable(1, c)
                    pTot = pTot + pv(nx)
                End If 'r > 0 and c > 0
            End If 'ns > 0
        End If 'IsNumeric(ns)
       
    Next ns
    
    'Cox-Sprague percentage of perfection before discards
    If pTot > 0 Then cs = xTot / pTot Else cs = 0
    
'Discards:

'The race result whose removal results in the greatest improvement to the
'C-S score cs is discarded and the improved CS score is returned.
'This process is repeated if there is more than one discard.

    If nDiscards > nx - 1 Then nDiscards = nx - 1
    If nDiscards > 0 Then
   
        For j = 1 To nDiscards
            im = 0
            For i = 1 To nx
            
                ' Skip if already discarded or marked as not excludable
                 If xv(i) > 0 Then
                     'Find discard that improves the cs percentage the most
                     If (pTot - pv(i) > 0) Then csz = (xTot - xv(i)) / (pTot - pv(i)) Else csz = 0
                     If csz >= cs Then
                         cs = csz
                         im = i
                     End If ' csz >= cs
                 End If ' xv(i) > 0
                 
            Next i
            
            'Discard the worst race
            If im > 0 Then
                xTot = xTot - xv(im)
                pTot = pTot - pv(im)
            
                'Mark as discarded
                xv(im) = 0
            End If 'im  > 0
           
        Next j
                
    End If 'nDiscards > 0
        
    csScore = cs
       
End Function


Function csTable1(r As Integer, c As Integer, Optional modified As Boolean = True) As Double
' Returns csTable( r, c).
' If modified = TRUE csTable(2,2)=7; csTable(3,2)=5; Else csTable(2,2)=4; csTable(3,2)=0.
' 2022-07-25: Extended to return Cox-Sprague scores for races with more than 20 starters

    Dim a(21) As Integer
    Dim i As Integer
    
    csTable1 = 0
    
    If c > 20 Then c = 20
       
    If r = 1 Then
        csTable1 = Array(0, 10, 31, 43, 52, 60, 66, 72, 76, 80, 84, 87, 90, 92, 94, 96, 97, 98, 99, 100)(c)
    ElseIf modified And ((r > 21) And (c = 20)) Then
        csTable1 = 58 - 2 * Log(r - 20) / Log(10)
    ElseIf Not (modified) And ((r > 21) And (c = 20)) Then
        csTable1 = 79 - r
    ElseIf c > 1 And r < c + 2 Then
             a(1) = Array(0, 10, 31, 43, 52, 60, 66, 72, 76, 80, 84, 87, 90, 92, 94, 96, 97, 98, 99, 100)(c)
             a(2) = a(1) - 6
          
            If r > 2 Then
                For i = 3 To 4
                    a(i) = a(i - 1) - 4
                Next i
            End If
            
            If r > 4 Then
                For i = 5 To 6
                    a(i) = a(i - 1) - 3
                Next i
            End If
            
            If r > 6 Then
                For i = 7 To 8
                      a(i) = a(i - 1) - 2
                Next i
            End If
            
            If r > 8 Then
                For i = 9 To 13
                        a(i) = a(i - 1) - 2
                Next i
            End If
            
            If r > 13 Then
                For i = 14 To 21
                       a(i) = a(i - 1) - 1
                Next i
            End If
            
            If (modified And c = 2) Then
                a(2) = 7
                a(3) = 5
            End If
                                            
         csTable1 = a(r)
        End If
    
End Function

