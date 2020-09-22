Attribute VB_Name = "modMst"
Option Explicit
Option Base 1

Public Function numberOfEdge(m() As Double) As Integer
    Dim a As Integer
    Dim b As Integer
    Dim number As Integer
    For a = 1 To UBound(m)
        For b = a To UBound(m)
            If m(a, b) > 0 Then number = number + 1
        Next
    Next
    numberOfEdge = number
End Function

Public Function countTotalWeight(mst() As Double) As Double
    Dim a As Integer
    Dim b As Integer
    Dim totalWeight As Double
    
    For a = 1 To UBound(mst)
        For b = a To UBound(mst)
            totalWeight = totalWeight + mst(a, b)
        Next
    Next
    countTotalWeight = totalWeight
End Function

Public Function findBiggestEdge(m() As Double) As Double
    Dim biggestEdge As Double
    Dim a As Integer
    Dim b As Integer
    biggestEdge = m(1, 1)
    For a = 1 To UBound(m)
        For b = 1 To UBound(m)
            m(b, a) = m(a, b)
            If biggestEdge < m(a, b) Then biggestEdge = m(a, b)
        Next
    Next
    findBiggestEdge = biggestEdge
End Function
Public Sub prim(matrix() As Double, ByVal startAtNode As Integer, mst() As Double)
    Dim numberOfNode As Integer
    Dim idx As Integer
    Dim prevNode As Integer
    Dim nextNode As Integer
    Dim visitedNode() As Integer
    Dim biggestEdge As Double
    Dim shortestEdge As Double
    Dim matBackup() As Double
    Dim i As Integer
    Dim j As Integer
    Dim a As Integer
    Dim b As Integer
    Dim col As Integer
    Dim row As Integer
    
    biggestEdge = findBiggestEdge(matrix)
    numberOfNode = UBound(matrix)
    ReDim visitedNode(numberOfNode) As Integer
    matBackup = matrix
    idx = 1
    prevNode = 1
    nextNode = 1
    visitedNode(1) = startAtNode
    
    For i = 1 To numberOfNode - 1
        shortestEdge = biggestEdge
        For col = 1 To idx
            startAtNode = visitedNode(col)
            For row = 1 To numberOfNode ' - 1
                If (matrix(row, startAtNode) < shortestEdge) And (matrix(row, startAtNode) > 0) Then
                    shortestEdge = matrix(row, startAtNode)
                    prevNode = startAtNode
                    nextNode = row
                End If
            Next
        Next

        idx = idx + 1
        startAtNode = nextNode
        visitedNode(idx) = startAtNode
        mst(nextNode, prevNode) = matBackup(prevNode, nextNode)
        mst(prevNode, nextNode) = matBackup(prevNode, nextNode)
        
        For a = 1 To idx
            For b = 1 To idx
                matrix(visitedNode(b), visitedNode(a)) = 0
                matrix(visitedNode(a), visitedNode(b)) = 0
            Next
        Next
    Next
End Sub

Public Function isCyclic(matrix() As Double, ByVal startNode As Integer, ByVal endNode As Integer) As Boolean
    Dim numberOfNode As Integer
    Dim pEnd As Integer
    Dim pStart As Integer
    Dim stackPointer As Integer
    Dim stackStart() As Integer
    Dim stackEnd() As Integer
    
    numberOfNode = UBound(matrix)
    ReDim stackStart(numberOfNode * numberOfNode) As Integer
    ReDim stackEnd(numberOfNode * numberOfNode) As Integer
    
    pEnd = endNode
    stackPointer = 0
    
    Do While startNode <> endNode
        For pStart = 1 To numberOfNode
            If matrix(startNode, pStart) < 0 And pStart <> pEnd Then
                stackPointer = stackPointer + 1
                stackStart(stackPointer) = startNode
                stackEnd(stackPointer) = pStart
            End If
        Next
        
        If stackPointer > 0 Then
            startNode = stackEnd(stackPointer)
            pEnd = stackStart(stackPointer)
            stackPointer = stackPointer - 1
        Else
            isCyclic = False
            Exit Function
        End If
    Loop
    isCyclic = True
End Function
  
Public Sub kruskal(matrix() As Double, mst() As Double)
    Dim numberOfNode As Integer
    Dim biggestEdge As Double
    Dim shortestEdge As Double
    Dim row As Integer
    Dim col As Integer
    Dim startNode As Integer
    Dim endNode As Integer
    
    numberOfNode = UBound(matrix)
    biggestEdge = findBiggestEdge(matrix)
    
    Do While True
        shortestEdge = biggestEdge
        For row = 1 To numberOfNode
            For col = 1 To numberOfNode
                If matrix(row, col) < shortestEdge And matrix(row, col) > 0 Then
                    shortestEdge = matrix(row, col)
                    startNode = row
                    endNode = col
                End If
            Next
        Next
        
        If shortestEdge = biggestEdge Then Exit Sub
        
        If Not isCyclic(matrix, startNode, endNode) Then
            mst(startNode, endNode) = matrix(startNode, endNode)
            mst(endNode, startNode) = matrix(startNode, endNode)
            matrix(startNode, endNode) = -1
            matrix(endNode, startNode) = -1
        Else
            matrix(startNode, endNode) = biggestEdge
            matrix(endNode, startNode) = biggestEdge
        End If
    Loop
End Sub

