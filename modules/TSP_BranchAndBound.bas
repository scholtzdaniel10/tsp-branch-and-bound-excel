' TSP Branch & Bound prototype (VBA)
' Paste into an Excel standard module.
Option Explicit

'------------- Config (adjust cell addresses if you change layout) -------------
Const rowIterStart As Long = 30 ' first row of iteration log
Const colNodeID As Long = 1     ' column A
Const colParent As Long = 2     ' column B
Const colFixedTour As Long = 3  ' column C
Const colExcluded As Long = 4   ' column D
Const colLB As Long = 5         ' column E
Const colBranch As Long = 6     ' column F
Const colIncumbentTour As Long = 7 ' column G
Const colIncumbentVal As Long = 8  ' column H
Const colPrune As Long = 9      ' column I
Const colNotes As Long = 10     ' column J
'-------------------------------------------------------------------------------

' Node type to track partial tours
Private Type TNode
    id As String
    parentId As String
    tour() As Long        ' sequence of visited cities in order
    visited() As Boolean  ' visited flags
    costSoFar As Double
    lb As Double
    depth As Long
End Type

' Globals for Dist matrix and n
Dim gDist() As Double
Dim gN As Long
Dim gStart As Long

' Incumbent best tour and value
Dim incumbentTour() As Long
Dim incumbentVal As Double
Dim precFormat As String

' Utilities -------------------------------------------------------------------
Private Sub ReadInputs()
    Dim rng As Range
    gN = CLng(Range("nCities").Value)
    gStart = CLng(Range("StartCity").Value)
    ' read Dist into array (assume Dist is n x n)
    Set rng = Range("Dist")
    ReDim gDist(1 To gN, 1 To gN)
    Dim i As Long, j As Long
    For i = 1 To gN
        For j = 1 To gN
            gDist(i, j) = rng.Cells(i, j).Value
        Next j
    Next i
    precFormat = "0." & String(Clamp(CLng(Range("prec").Value), 0, 8), "0")
End Sub

Private Function Clamp(val As Long, lo As Long, hi As Long) As Long
    If val < lo Then Clamp = lo ElseIf val > hi Then Clamp = hi Else Clamp = val
End Function

' Nearest neighbour heuristic (returns cost and fills tour array)
Private Function NearestNeighbourUB() As Double
    Dim visited() As Boolean
    ReDim visited(1 To gN)
    Dim tour() As Long
    ReDim tour(1 To gN + 1)
    Dim cur As Long, best As Long
    Dim i As Long, k As Long
    cur = gStart
    tour(1) = cur
    visited(cur) = True
    For k = 2 To gN
        best = -1
        For i = 1 To gN
            If Not visited(i) Then
                If best = -1 Or gDist(cur, i) < gDist(cur, best) Then best = i
            End If
        Next i
        tour(k) = best
        visited(best) = True
        cur = best
    Next k
    tour(gN + 1) = gStart ' return to start
    ' compute cost
    Dim cost As Double: cost = 0
    For k = 1 To gN
        cost = cost + gDist(tour(k), tour(k + 1))
    Next k
    ' save incumbent
    ReDim incumbentTour(1 To gN + 1)
    For k = 1 To gN + 1: incumbentTour(k) = tour(k): Next k
    incumbentVal = cost
    NearestNeighbourUB = cost
End Function

' Prim's MST over remaining nodes (nodeset is Boolean array 1..gN; returns MST weight)
Private Function PrimMST(nodeset() As Boolean) As Double
    Dim INF As Double: INF = 1E+99
    Dim inMST() As Boolean, minEdge() As Double
    Dim i As Long, j As Long, u As Long, v As Long, cnt As Long
    ReDim inMST(1 To gN)
    ReDim minEdge(1 To gN)
    For i = 1 To gN
        minEdge(i) = INF
        inMST(i) = False
    Next i
    ' find any node in nodeset to start
    Dim startNode As Long: startNode = -1
    For i = 1 To gN
        If nodeset(i) Then
            startNode = i: Exit For
        End If
    Next i
    If startNode = -1 Then
        PrimMST = 0
        Exit Function
    End If
    minEdge(startNode) = 0
    Dim total As Double: total = 0
    cnt = 0
    Do While True
        u = -1
        For i = 1 To gN
            If nodeset(i) And Not inMST(i) Then
                If u = -1 Or minEdge(i) < minEdge(u) Then u = i
            End If
        Next i
        If u = -1 Then Exit Do
        inMST(u) = True
        total = total + minEdge(u)
        cnt = cnt + 1
        ' relax edges from u
        For v = 1 To gN
            If nodeset(v) And Not inMST(v) And gDist(u, v) < minEdge(v) Then
                minEdge(v) = gDist(u, v)
            End If
        Next v
    Loop
    PrimMST = total
End Function

' Lower bound wrapper (calls Hungarian or MST version implemented in the other module)
Private Function LowerBoundForNode(nd As TNode) As Double
    ' This function is overridden by the Hungarian-enabled module (if included)
    ' Fallback to MST-based bound here if that module isn't present
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
End Function

' Logging helper: writes node info to the next available iteration log row
Private Sub LogNode(nd As TNode, parentId As String, nodeRow As Long, Optional note As String = "")
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long: r = nodeRow
    ws.Cells(r, colNodeID).Value = nd.id
    ws.Cells(r, colParent).Value = parentId
    ws.Cells(r, colFixedTour).Value = JoinTour(nd.tour)
    ws.Cells(r, colExcluded).Value = "" ' kept blank for this branching style
    ws.Cells(r, colLB).Value = Round(nd.lb, Range("prec").Value)
    ws.Cells(r, colBranch).Value = "" ' set when expanded
    ws.Cells(r, colIncumbentTour).Value = IIf(incumbentVal < 1E+90, JoinTour(incumbentTour), "")
    ws.Cells(r, colIncumbentVal).Value = IIf(incumbentVal < 1E+90, Round(incumbentVal, Range("prec").Value), "")
    ws.Cells(r, colPrune).Value = "" ' set later if pruned
    ws.Cells(r, colNotes).Value = note
End Sub

Private Function JoinTour(arr() As Long) As String
    Dim s As String, i As Long
    s = ""
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> 0 Then
            If s = "" Then s = CStr(arr(i)) Else s = s & "->" & CStr(arr(i))
        End If
    Next i
    JoinTour = s
End Function

' Main B&B routine (simple best-first)
Public Sub RunTSPBranchAndBound()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.Range(ws.Rows(rowIterStart & ":" & ws.Rows.Count)).ClearContents
    Call ReadInputs
    ' initial UB from nearest neighbour
    incumbentVal = 1E+90
    Call NearestNeighbourUB
    Dim nextRow As Long: nextRow = rowIterStart
    ' create root node
    Dim root As TNode
    ReDim root.tour(1 To 1)
    ReDim root.visited(1 To gN)
    root.tour(1) = gStart
    Dim i As Long
    For i = 1 To gN: root.visited(i) = False: Next i
    root.visited(gStart) = True
    root.costSoFar = 0
    root.depth = 1
    root.id = "1"
    root.parentId = ""
    root.lb = LowerBoundForNode(root)
    ' priority queue as array
    Dim queue() As TNode
    ReDim queue(0 To 0)
    queue(0) = root
    ' log root
    Call LogNode(root, "", nextRow, "root")
    nextRow = nextRow + 1
    Dim iter As Long: iter = 0
    Do While UBound(queue) >= 0
        ' pop best (lowest lb)
        Dim bestIdx As Long: bestIdx = 0
        Dim k As Long
        For k = 0 To UBound(queue)
            If queue(k).lb < queue(bestIdx).lb Then bestIdx = k
        Next k
        Dim cur As TNode: cur = queue(bestIdx)
        ' remove bestIdx from queue
        Dim newQ() As TNode
        If UBound(queue) > 0 Then
            ReDim newQ(0 To UBound(queue) - 1)
            Dim p As Long: p = 0
            For k = 0 To UBound(queue)
                If k <> bestIdx Then
                    newQ(p) = queue(k): p = p + 1
                End If
            Next k
            queue = newQ
        Else
            ReDim queue(-1 To -1) ' empty
        End If
        ' log expanded
        Call LogNode(cur, cur.parentId, nextRow, "expanded")
        Dim expandedRow As Long: expandedRow = nextRow
        nextRow = nextRow + 1
        iter = iter + 1
        ' check if complete tour
        If cur.depth = gN Then
            Dim total As Double: total = cur.costSoFar + gDist(cur.tour(UBound(cur.tour)), gStart)
            Dim fullt() As Long: ReDim fullt(1 To gN + 1)
            Dim idx As Long
            For idx = 1 To gN: fullt(idx) = cur.tour(idx): Next idx
            fullt(gN + 1) = gStart
            If total < incumbentVal Then
                incumbentVal = total
                incumbentTour = fullt
                ActiveSheet.Cells(expandedRow, colIncumbentTour).Value = JoinTour(incumbentTour)
                ActiveSheet.Cells(expandedRow, colIncumbentVal).Value = Round(incumbentVal, Range("prec").Value)
                ActiveSheet.Cells(expandedRow, colNotes).Value = "leaf (complete), improved UB"
            Else
                ActiveSheet.Cells(expandedRow, colNotes).Value = "leaf (complete)"
            End If
            GoTo ContinueLoop
        End If
        ' expand node: branch on every remaining city
        Dim childCities As Collection: Set childCities = New Collection
        For i = 1 To gN
            If Not cur.visited(i) Then childCities.Add i
        Next i
        Dim city As Variant
        For Each city In childCities
            Dim child As TNode
            child = cur
            Dim d As Long: d = child.depth + 1
            ReDim Preserve child.tour(1 To d)
            child.tour(d) = city
            child.visited(city) = True
            child.depth = d
            child.costSoFar = child.costSoFar + gDist(child.tour(d - 1), city)
            child.parentId = cur.id
            child.id = cur.id & "." & CStr(city)
            child.lb = LowerBoundForNode(child)
            If child.lb >= incumbentVal Then
                ActiveSheet.Cells(nextRow, colNodeID).Value = child.id
                ActiveSheet.Cells(nextRow, colParent).Value = child.parentId
 due to length limit...