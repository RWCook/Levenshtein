Attribute VB_Name = "LevenshteinUDF"
Option Explicit

'Name:	LevDistance
'Function: Returns the Levenshtein string distance between two string
'
Public Function LevDistance(ByVal strA As String, ByVal strB As String) As Integer
Dim arrMatrix() As Variant
ReDim arrMatrix(Len(strB) + 1, Len(strA) + 1)
Const constMaxCharacters As Integer = 999

'#Put String A into Matrix
Dim intLenA As Integer

For intLenA = 1 To Len(strA)
    arrMatrix(0, intLenA + 1) = UCase$(Mid$(strA, intLenA, 1))
Next intLenA

Dim intLenB As Integer

For intLenB = 1 To Len(strB)
    arrMatrix(intLenB + 1, 0) = UCase$(Mid$(strB, intLenB, 1))
Next intLenB

Dim intX As Integer
Dim intY As Integer
Dim intCost As Integer
Dim arrPossibleValues(3) As Variant

For intX = 1 To UBound(arrMatrix, 1)
    For intY = 1 To UBound(arrMatrix, 2)
       
            If arrMatrix(0, intY) = arrMatrix(intX, 0) Then
                arrMatrix(intX, intY) = 0
                intCost = 0
            Else
                arrMatrix(intX, intY) = 1
                intCost = 1
            End If
                
            If IsNumeric(arrMatrix(intX - 1, intY)) = True Then
                arrPossibleValues(0) = arrMatrix(intX - 1, intY) + 1
            Else
                arrPossibleValues(0) = constMaxCharacters
            End If
                
            If IsNumeric(arrMatrix(intX, intY - 1)) = True Then
                arrPossibleValues(1) = arrMatrix(intX, intY - 1) + 1
            Else
                arrPossibleValues(1) = constMaxCharacters
            End If
                
            If IsNumeric(arrMatrix(intX - 1, intY - 1)) = True Then
                arrPossibleValues(2) = arrMatrix(intX - 1, intY - 1) + intCost
            Else
                arrPossibleValues(2) = constMaxCharacters
            End If
               
            arrMatrix(intX, intY) = WorksheetFunction.Min(arrPossibleValues)
                
    Next intY
Next intX

LevDistance = arrMatrix(UBound(arrMatrix, 1), UBound(arrMatrix, 2))

End Function
