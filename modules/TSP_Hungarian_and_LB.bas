' Hungarian assignment solver + LB selection for TSP B&B
' Paste into an Excel standard module (with the other TSP B&B code).
Option Explicit

Private Const BIGM As Double = 1E+8

' Hungarian algorithm to compute min assignment cost
Public Function HungarianMinCost(costMat() As Double) As Double
    Dim n As Long
    n = UBound(costMat, 1)
    Dim u() As Double, v() As Double
    Dim p() As Long, way() As Long
    ReDim u(0 To n), v(0 To n)
    ReDim p(0 To n), way(0 To n)
    Dim i As Long, j As Long, j0 As Long, i0 As Long, cur As Long
    Dim minv() As Double
    ReDim minv(0 To n)
    Dim used() As Boolean
    ReDim used(0 To n)
    p(0) = 0
    For i = 1 To n
        p(0) = i
        j0 = 0
        For j = 0 To n
            minv(j) = 1E+99
            used(j) = False
            way(j) = 0
        Next j
        Do
            used(j0) = True
            i0 = p(j0)
            cur = -1
            Dim delta As Double: delta = 1E+99
            For j = 1 To n
                If Not used(j) Then
                    Dim curCost As Double
                    curCost = costMat(i0, j) - u(i0) - v(j)
                    If curCost < minv(j) Then
                        minv(j) = curCost
                        way(j) = j0
                    End If
                    If minv(j) < delta Then
                        delta = minv(j)
                        cur = j
                    End If
                End If
            Next j
            For j = 0 To n
                If used(j) Then
                    u(p(j)) = u(p(j)) + delta
                    v(j) = v(j) - delta
                Else
                    minv(j) = minv(j) - delta
                End If
            Next j
            j0 = cur
        Loop While p(j0) <> 0
        Do
            Dim j1 As Long
            j1 = way(j0)
            p(j0) = p(j1)
            j0 = j1
        Loop While j0 <> 0
    Next i
    Dim assignmentCost As Double: assignmentCost = 0
    For j = 1 To n
        If p(j) > 0 Then
            assignmentCost = assignmentCost + costMat(p(j), j)
        End If
    Next j
    HungarianMinCost = assignmentCost
End Function

' Build a modified cost matrix for partial node, forcing fixed edges and forbidding excluded ones.
Private Sub BuildCostMatrixForNode(nd As TNode, ByRef outMat() As Double)
    Dim n As Long: n = gN
    ReDim outMat(1 To n, 1 To n)
    Dim i As Long, j As Long
    ' copy base distances
    For i = 1 To n
        For j = 1 To n
            outMat(i, j) = gDist(i, j)
        Next j
    Next i
    ' forbid self-loops
    For i = 1 To n: outMat(i, i) = BIGM: Next i
    ' Force fixed edges from nd.tour sequence: for each consecutive pair (a->b) in nd.tour:
    Dim k As Long, a As Long, b As Long
    For k = LBound(nd.tour) To UBound(nd.tour) - 1
        a = nd.tour(k)
        b = nd.tour(k + 1)
        ' set entire row a and col b to BIGM except (a,b)
        For j = 1 To n: outMat(a, j) = BIGM: Next j
        For i = 1 To n: outMat(i, b) = BIGM: Next i
        outMat(a, b) = gDist(a, b)
    Next k
End Sub

' Combined LowerBoundForNode that chooses MST or Hungarian based on sheet setting
Private Function LowerBoundForNode(nd As TNode) As Double
    Dim boundMethod As String
    On Error Resume Next
    boundMethod = CStr(Range("BoundMethod").Value)
    If boundMethod = "" Then boundMethod = "MST"
    On Error GoTo 0
    If UCase(boundMethod) = "HUNGARIAN" Or UCase(boundMethod) = "ASSIGNMENT" Then
        Dim mat() As Double
        Call BuildCostMatrixForNode(nd, mat)
        Dim assignCost As Double
        assignCost = HungarianMinCost(mat)
        LowerBoundForNode = assignCost
    Else
        Dim remaining() As Boolean
        ReDim remaining(1 To gN)
        Dim i As Long
        For i = 1 To gN
            remaining(i) = Not nd.visited(i)
        Next i
        Dim mstw As Double: mstw = PrimMST(remaining)
        Dim last As Long: last = nd.tour(UBound(nd.tour))
        Dim minOut As Double: minOut = 1E+99
        Dim minIn As Double: minIn = 1E+99
        Dim v As Long
        For v = 1 To gN
            If remaining(v) Then
                If gDist(last, v) < minOut Then minOut = gDist(last, v)
                If gDist(v, nd.tour(1)) < minIn Then minIn = gDist(v, nd.tour(1))
            End If
        Next v
        If minOut > 1E+90 Then minOut = 0
        If minIn > 1E+90 Then minIn = 0
        LowerBoundForNode = nd.costSoFar + mstw + minOut + minIn
    End If
End Function
