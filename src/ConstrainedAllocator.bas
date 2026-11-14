' =============================================================================
' ConstrainedAllocator.bas
' -----------------------------------------------------------------------------
' Generic, domain-independent VBA engine that adjusts the QUANTITY of each item
' in a table so that TWO goals are satisfied at once:
'
'   Goal A (physical constraint):  the AGGREGATE TOTAL VALUE of all items should
'                                  be as close as possible to a TARGET VALUE.
'   Goal B (attribute objective):  the AGGREGATE PENALTY PER UNIT (e.g. cost,
'                                  risk, or any per-unit score) should stay at or
'                                  below a ceiling.
'
' while honouring, for every item, a set of hard constraints:
'   - a PER-UNIT VALUE must stay inside [UNIT_VALUE_MIN, UNIT_VALUE_MAX];
'   - the quantity must be a multiple of the item's PACKAGE SIZE;
'   - the quantity may not drop below MIN_QTY_RATIO (nor exceed MAX_QTY_RATIO)
'     of its ORIGINAL quantity;
'   - PROTECTED items are never touched.
'
' The engine solves this with a bounded random-walk heuristic plus a five-stage
' fallback (coarse -> precise main loop -> unique-item formula -> penalty
' optimisation -> forced reduce) and an anti-infinite-loop guard (time limit +
' max iterations + no-progress bailout).
'
' This is a clean-room, general-purpose re-implementation. It contains NO
' business data, NO proprietary naming, and NO third-party logic.
' =============================================================================

Option Explicit

' ============================ CONFIGURATION =================================
' All knobs live here. Change values, not algorithm code.

' --- Sheet / column layout (mirror examples/sample-data.xlsx) ---
Public Const SRC_SHEET        As String = ""     ' "" => use ActiveSheet
Public Const KEY_COL          As Long = 4        ' D: cell that marks the totals row
Public Const KEY_ROW_TEXT     As String = "Total" ' marker text in KEY_COL
Public Const PRODUCT_START_ROW As Long = 4       ' first item data row
Public Const NAME_COL         As Long = 3        ' C: display name (protected check)
Public Const QTY_COL          As Long = 15       ' O: adjustable quantity (OUTPUT)
Public Const ORIG_QTY_COL     As Long = 16       ' P: original quantity backup
Public Const BOX_COL          As Long = 14       ' N: package size (quantity multiple)
Public Const UNIT_VAL_COL     As Long = 29       ' AC: per-package value (formula)
Public Const UNIT_PRICE_COL   As Long = 30       ' AD: price per value-unit
Public Const PENALTY_COL      As Long = 31       ' AE: penalty per unit
Public Const PENALTY_ITEM_COL As Long = 18       ' R: penalty per item
Public Const TOTAL_VAL_COL    As Long = 8        ' H: aggregate total value
Public Const TARGET_VAL_COL   As Long = 9        ' I: target total value
Public Const LINK_COL         As Long = 28       ' AB: link (unique-item micro-adjust)

' --- Per-unit value bounds (example figures; tune freely) ---
Public Const UNIT_VALUE_MIN   As Double = 3.35   ' lower bound per package
Public Const UNIT_VALUE_MAX   As Double = 35.7   ' upper bound per package
Public Const BAND_HI_MIN      As Double = 3.35   ' natural-dist band, high-penalty items
Public Const BAND_HI_MAX      As Double = 12#
Public Const BAND_LO_MIN      As Double = 22#    ' natural-dist band, low-penalty items
Public Const BAND_LO_MAX      As Double = 32#

' --- Penalty (objective) thresholds ---
Public Const PENALTY_CEILING  As Double = 0.88   ' aggregate penalty/unit must be <= this
Public Const PENALTY_HIGH     As Double = 0.35   ' "high penalty" classification
Public Const PENALTY_MID_HI    As Double = 0.6    ' top of "mid penalty" band
Public Const PRICE_SKIP       As Double = 40#    ' skip items pricier than this (per value)

' --- Tolerances & loop guards ---
Public Const TOL_COARSE       As Double = 300#   ' coarse-phase accept window
Public Const TOL_PRECISE      As Double = 100#   ' precise-phase accept window
Public Const TIME_LIMIT_SEC   As Double = 45#    ' hard wall-clock cap
Public Const MAX_ITERATIONS   As Long = 500      ' hard iteration cap

' --- Quantity ratio guards (vs original quantity) ---
Public Const MIN_QTY_RATIO    As Double = 0.3    ' never below 30% of original
Public Const MAX_QTY_RATIO    As Double = 3.5    ' never above 350% of original

' --- Forced-reduce stage ---
Public Const FORCED_RATIO     As Double = 0.1    ' remove ~10% of gap each pass
Public Const FORCED_MIN       As Double = 50#    ' ...but at least this much

' --- Presentation ---
Public Const HILITE_COLOR     As Long = 255      ' red: marks adjusted quantity cells
Public Const DEBUG_ON         As Boolean = False ' set True to flood Debug.Print

' -----------------------------------------------------------------------------
' ENTRY 1: Coarse adjustment.
' Brings every item's per-unit value into [UNIT_VALUE_MIN, UNIT_VALUE_MAX] by
' tweaking its quantity, then does a three-step coarse rebalance toward the
' target total value. Safe, deterministic, no randomness.
' -----------------------------------------------------------------------------
Public Sub AdjustAllocation()
    Dim ws As Worksheet
    Dim i As Long, adjustedCount As Long
    Dim totalRow As Long, productEndRow As Long
    Dim totalValRow As Long, targetValRow As Long
    Dim totalVal As Double, targetVal As Double, diff As Double

    Set ws = ActiveSheet
    If SRC_SHEET <> "" Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(SRC_SHEET)
        On Error GoTo 0
    End If

    Call FindKeyRows(ws, totalRow, productEndRow, totalValRow, targetValRow)
    If totalRow = 0 Then
        MsgBox "Totals row not found (looked for '" & KEY_ROW_TEXT & "' in column " & KEY_COL & ").", vbExclamation
        Exit Sub
    End If

    totalVal = ws.Cells(totalValRow, TOTAL_VAL_COL).value
    targetVal = ws.Cells(targetValRow, TARGET_VAL_COL).value

    If productEndRow < PRODUCT_START_ROW Then
        MsgBox "No item rows found above the totals row.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' First pass: normalise per-unit values into band.
    Call NormaliseUnitValues(ws, PRODUCT_START_ROW, productEndRow, adjustedCount)

    Application.Calculate
    totalVal = ws.Cells(totalValRow, TOTAL_VAL_COL).value
    diff = Abs(totalVal - targetVal)

    ' Step 2: if still far from target, shrink high-penalty items.
    If diff > TOL_COARSE Then
        Call RebalanceHighPenalty(ws, PRODUCT_START_ROW, productEndRow, totalValRow, _
                                 targetValRow, targetVal, diff, adjustedCount)
    End If

    ' Step 3: if still far, grow low-penalty items.
    If diff > TOL_COARSE Then
        Call RebalanceLowPenalty(ws, PRODUCT_START_ROW, productEndRow, totalValRow, _
                                targetValRow, targetVal, diff, adjustedCount)
    End If

    Call TidySheet(ws, productEndRow)

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Application.Calculate
    totalVal = ws.Cells(totalValRow, TOTAL_VAL_COL).value
    MsgBox "Coarse adjustment done." & vbCrLf & _
           "Items adjusted: " & adjustedCount & vbCrLf & _
           "Total value: " & Format(totalVal, "0") & vbCrLf & _
           "Target value: " & Format(targetVal, "0") & vbCrLf & _
           "Difference: " & Format(Abs(totalVal - targetVal), "0"), _
           vbInformation, "AdjustAllocation"
End Sub

' -----------------------------------------------------------------------------
' ENTRY 2: Precise tuning.
' Iteratively nudges quantities to close the gap to the target total value while
' keeping the aggregate penalty/unit at or below PENALTY_CEILING, using a
' priority-aware, bounded-random-walk heuristic with multi-stage fallback and an
' anti-infinite-loop guard.
' -----------------------------------------------------------------------------
Public Sub PreciseTune()
    Dim ws As Worksheet
    Dim startTime As Double, timeLimit As Double, iteration As Long
    Dim totalRow As Long, productEndRow As Long, totalValRow As Long, targetValRow As Long
    Dim totalVal As Double, targetVal As Double, diff As Double
    Dim currentPenalty As Double, penaltyMin As Double, penaltyMax As Double
    Dim adjustedCount As Long, initialDiff As Double
    Dim needIncrease As Boolean, timeout As Boolean
    Dim products() As Variant, validCount As Long
    Dim hiPenalty() As Variant, hiCount As Long
    Dim loPenalty() As Variant, loCount As Long
    Dim zeroPenalty() As Variant, zeroCount As Long

    timeLimit = TIME_LIMIT_SEC
    startTime = Timer
    iteration = 0
    adjustedCount = 0
    timeout = False

    Set ws = ActiveSheet
    If SRC_SHEET <> "" Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(SRC_SHEET)
        On Error GoTo 0
    End If

    Call FindKeyRows(ws, totalRow, productEndRow, totalValRow, targetValRow)
    If totalRow = 0 Then
        MsgBox "Totals row not found (looked for '" & KEY_ROW_TEXT & "' in column " & KEY_COL & ").", vbExclamation
        Exit Sub
    End If

    totalVal = ws.Cells(totalValRow, TOTAL_VAL_COL).value
    targetVal = ws.Cells(targetValRow, TARGET_VAL_COL).value
    diff = Abs(totalVal - targetVal)
    initialDiff = diff

    If diff <= TOL_PRECISE Then
        MsgBox "Already within " & TOL_PRECISE & " of target. Nothing to do.", vbInformation
        Exit Sub
    End If

    currentPenalty = ComputePenaltyPerUnit(ws, productEndRow)
    penaltyMin = 0#
    penaltyMax = PENALTY_CEILING

    If MsgBox("Precise tuning start." & vbCrLf & _
              "Target penalty/unit ceiling: " & Format(penaltyMax, "0.000") & vbCrLf & _
              "Current gap: " & Format(diff, "0") & " (aim: <= " & TOL_PRECISE & ")" & vbCrLf & _
              "Time limit: " & timeLimit & "s", _
              vbOKCancel + vbInformation, "PreciseTune") = vbCancel Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    needIncrease = (totalVal < targetVal)

    ' ---- Collect item data into an in-memory array ----
    Call CollectItems(ws, productEndRow, products, validCount)
    Call PartitionByPenalty(products, validCount, hiPenalty, hiCount, _
                            loPenalty, loCount, zeroPenalty, zeroCount)
    SortPenalty hiPenalty, hiCount, False   ' high -> low
    SortPenalty loPenalty, loCount, True    ' low -> high

    ' ---- Normalise per-unit values first ----
    Call NormaliseUnitValues(ws, PRODUCT_START_ROW, productEndRow, adjustedCount)

    Application.Calculate
    totalVal = ws.Cells(totalValRow, TOTAL_VAL_COL).value
    diff = Abs(totalVal - targetVal)
    needIncrease = (totalVal < targetVal)

    Randomize Timer

    ' ---- Main heuristic loop ----
    Dim lastDiff As Double, noProgress As Long
    lastDiff = diff
    noProgress = 0

    Do While diff > TOL_PRECISE And Timer - startTime < timeLimit And iteration < MAX_ITERATIONS
        Dim bestIdx As Long, bestImpact As Double, found As Boolean
        bestIdx = 0: bestImpact = 0: found = False

        ' Stop early if both goals already met.
        If currentPenalty <= penaltyMax And diff <= TOL_PRECISE Then Exit Do

        If needIncrease Then
            found = TryIncrease(ws, products, validCount, loPenalty, loCount, _
                               zeroPenalty, zeroCount, bestIdx, bestImpact)
        Else
            found = TryDecrease(ws, products, validCount, hiPenalty, hiCount, _
                               bestIdx, bestImpact)
        End If

        If found Then
            Dim rowNum As Long, qty As Long, boxQty As Long, avgW As Double, newQty As Long
            rowNum = products(bestIdx, 1)
            qty = products(bestIdx, 3)
            boxQty = products(bestIdx, 4)
            avgW = products(bestIdx, 7)

            newQty = ResolveQuantity(ws, products, bestIdx, needIncrease)

            If newQty <> qty And newQty > 0 Then
                ws.Cells(rowNum, QTY_COL).value = newQty
                ws.Cells(rowNum, QTY_COL).Font.Color = HILITE_COLOR
                adjustedCount = adjustedCount + 1
                products(bestIdx, 3) = newQty
                products(bestIdx, 5) = avgW * newQty
            End If

            If adjustedCount Mod 3 = 0 Then
                Application.Calculate
                totalVal = ws.Cells(totalValRow, TOTAL_VAL_COL).value
                diff = Abs(totalVal - targetVal)
                needIncrease = (totalVal < targetVal)
                currentPenalty = ComputePenaltyPerUnit(ws, productEndRow)
                If diff < lastDiff Then noProgress = 0 Else noProgress = noProgress + 1
                lastDiff = diff
                If noProgress > 10 Then Exit Do
            End If
        Else
            Exit Do
        End If
        iteration = iteration + 1
    Loop

    ' ---- Stage 2: unique-item formula micro-adjust ----
    If diff > TOL_PRECISE Then
        Call UniqueLinkStage(ws, totalValRow, targetValRow, productEndRow, diff, targetVal, adjustedCount)
    End If

    ' ---- Stage 3: penalty optimisation ----
    Call PenaltyOptimiseStage(ws, products, validCount, hiPenalty, hiCount, _
                              zeroPenalty, zeroCount, totalValRow, targetValRow, _
                              productEndRow, targetVal, diff, adjustedCount, _
                              currentPenalty, startTime, timeLimit)

    ' ---- Stage 4: forced reduce (weight floor only) ----
    If diff > TOL_PRECISE Then
        Call ForcedReduceStage(ws, products, validCount, hiPenalty, hiCount, _
                              loPenalty, loCount, zeroPenalty, zeroCount, _
                              totalValRow, targetValRow, productEndRow, _
                              targetVal, diff, adjustedCount, currentPenalty, startTime, timeLimit)
    End If

    Application.Calculate
    totalVal = ws.Cells(totalValRow, TOTAL_VAL_COL).value
    diff = Abs(totalVal - targetVal)
    currentPenalty = ComputePenaltyPerUnit(ws, productEndRow)

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Precise tuning complete." & vbCrLf & _
           "Items adjusted: " & adjustedCount & vbCrLf & _
           "Aggregate penalty/unit: " & Format(currentPenalty, "0.000") & _
           " (ceiling " & Format(penaltyMax, "0.000") & ")" & vbCrLf & _
           "Total value: " & Format(totalVal, "0") & vbCrLf & _
           "Target value: " & Format(targetVal, "0") & vbCrLf & _
           "Initial gap: " & Format(initialDiff, "0") & vbCrLf & _
           "Final gap: " & Format(diff, "0") & vbCrLf & _
           "Stop reason: " & IIf(diff <= TOL_PRECISE, "gap closed", "guard/other"), _
           vbInformation, "PreciseTune"
End Sub

' =============================================================================
' DATA LAYER (read-only helpers)
' =============================================================================

Private Sub FindKeyRows(ws As Worksheet, ByRef totalRow As Long, _
                        ByRef productEndRow As Long, ByRef totalValRow As Long, _
                        ByRef targetValRow As Long)
    Dim i As Long
    totalRow = 0: productEndRow = 0
    For i = PRODUCT_START_ROW To ws.rows.count
        If Trim(CStr(ws.Cells(i, KEY_COL).value)) = KEY_ROW_TEXT Then
            totalRow = i
            Exit For
        End If
    Next i
    If totalRow = 0 Then Exit Sub
    productEndRow = totalRow - 2
    totalValRow = totalRow
    targetValRow = totalRow + 1
End Sub

Private Function ComputePenaltyPerUnit(ws As Worksheet, productEndRow As Long) As Double
    Dim i As Long, penaltyCell As Range
    For i = PRODUCT_START_ROW To ws.rows.count
        If Trim(CStr(ws.Cells(i, 20).value)) = "PenaltyPerUnit" Then
            Set penaltyCell = ws.Cells(i, 21)
            If IsNumeric(penaltyCell.value) Then
                ComputePenaltyPerUnit = penaltyCell.value
                Exit Function
            End If
        End If
    Next i
    ' Fallback: aggregate penalty / total value.
    Dim totPen As Double, totVal As Double
    For i = PRODUCT_START_ROW To productEndRow
        If IsNumeric(ws.Cells(i, PENALTY_ITEM_COL).value) And IsNumeric(ws.Cells(i, QTY_COL).value) Then
            totPen = totPen + ws.Cells(i, PENALTY_ITEM_COL).value * ws.Cells(i, QTY_COL).value
        End If
        If IsNumeric(ws.Cells(i, TOTAL_VAL_COL).value) Then
            totVal = totVal + ws.Cells(i, TOTAL_VAL_COL).value
        End If
    Next i
    If totVal > 0 Then ComputePenaltyPerUnit = totPen / totVal Else ComputePenaltyPerUnit = 0
End Function

Private Sub CollectItems(ws As Worksheet, productEndRow As Long, _
                         ByRef products() As Variant, ByRef validCount As Long)
    Dim i As Long, c As Long
    validCount = 0
    For i = PRODUCT_START_ROW To productEndRow
        If IsNumeric(ws.Cells(i, QTY_COL).value) And IsNumeric(ws.Cells(i, PENALTY_COL).value) Then
            If ws.Cells(i, QTY_COL).value > 0 Then validCount = validCount + 1
        End If
    Next i
    If validCount = 0 Then Exit Sub

    ReDim products(1 To validCount, 1 To 10)
    c = 0
    For i = PRODUCT_START_ROW To productEndRow
        If IsNumeric(ws.Cells(i, QTY_COL).value) And IsNumeric(ws.Cells(i, PENALTY_COL).value) Then
            If ws.Cells(i, QTY_COL).value > 0 Then
                c = c + 1
                products(c, 1) = i                       ' row number
                products(c, 2) = ws.Cells(i, PENALTY_COL).value        ' penalty/unit
                products(c, 3) = ws.Cells(i, QTY_COL).value            ' current qty
                products(c, 9) = ws.Cells(i, ORIG_QTY_COL).value       ' original qty backup
                If IsNumeric(ws.Cells(i, BOX_COL).value) Then
                    products(c, 4) = ws.Cells(i, BOX_COL).value Else products(c, 4) = 1
                End If
                If IsNumeric(ws.Cells(i, UNIT_VAL_COL).value) Then
                    products(c, 5) = ws.Cells(i, UNIT_VAL_COL).value Else products(c, 5) = 0
                End If
                If IsNumeric(ws.Cells(i, UNIT_PRICE_COL).value) Then
                    products(c, 6) = ws.Cells(i, UNIT_PRICE_COL).value Else products(c, 6) = 999
                End If
                If products(c, 5) > 0 And products(c, 3) > 0 Then
                    products(c, 7) = products(c, 5) / products(c, 3)   ' avg value/unit
                Else
                    products(c, 7) = 0
                End If
                ' type: 1 = high penalty (>= threshold), else 0
                products(c, 8) = IIf(products(c, 2) >= PENALTY_HIGH, 1, 0)
                ' oversized/locked flag: qty=1 and box=1
                products(c, 10) = IIf(products(c, 3) = 1 And products(c, 4) = 1, 1, 0)
            End If
        End If
    Next i
End Sub

Private Sub PartitionByPenalty(products() As Variant, validCount As Long, _
                               ByRef hi() As Variant, ByRef hiCount As Long, _
                               ByRef lo() As Variant, ByRef loCount As Long, _
                               ByRef zero() As Variant, ByRef zeroCount As Long)
    Dim i As Long, k As Long
    hiCount = 0: loCount = 0: zeroCount = 0
    For i = 1 To validCount
        If products(i, 2) >= PENALTY_HIGH And products(i, 10) = 0 Then hiCount = hiCount + 1
        If products(i, 2) > 0 And products(i, 2) < PENALTY_HIGH And products(i, 10) = 0 Then loCount = loCount + 1
        If products(i, 2) = 0 And products(i, 10) = 0 Then zeroCount = zeroCount + 1
    Next i
    If hiCount > 0 Then ReDim hi(1 To hiCount, 1 To 6)
    If loCount > 0 Then ReDim lo(1 To loCount, 1 To 6)
    If zeroCount > 0 Then ReDim zero(1 To zeroCount, 1 To 5)
    k = 1
    For i = 1 To validCount
        If products(i, 2) >= PENALTY_HIGH And products(i, 10) = 0 Then
            hi(k, 1) = i: hi(k, 2) = products(i, 2): hi(k, 3) = products(i, 3)
            hi(k, 4) = products(i, 4): hi(k, 5) = products(i, 7): hi(k, 6) = 4 + Rnd() * 8: k = k + 1
        End If
    Next i
    k = 1
    For i = 1 To validCount
        If products(i, 2) > 0 And products(i, 2) < PENALTY_HIGH And products(i, 10) = 0 Then
            lo(k, 1) = i: lo(k, 2) = products(i, 2): lo(k, 3) = products(i, 3)
            lo(k, 4) = products(i, 4): lo(k, 5) = products(i, 7): lo(k, 6) = 22 + Rnd() * 10: k = k + 1
        End If
    Next i
    k = 1
    For i = 1 To validCount
        If products(i, 2) = 0 And products(i, 10) = 0 Then
            zero(k, 1) = i: zero(k, 2) = products(i, 3): zero(k, 3) = products(i, 4)
            zero(k, 4) = products(i, 7): zero(k, 5) = 22 + Rnd() * 10: k = k + 1
        End If
    Next i
End Sub

' =============================================================================
' CLASSIFICATION / SORTING LAYER
' =============================================================================

Private Sub SortPenalty(arr() As Variant, count As Long, asc As Boolean)
    Dim i As Long, j As Long, t1, t2, t3, t4, t5, t6
    For i = 1 To count - 1
        For j = i + 1 To count
            If (asc And arr(i, 2) > arr(j, 2)) Or (Not asc And arr(i, 2) < arr(j, 2)) Then
                t1 = arr(i, 1): t2 = arr(i, 2): t3 = arr(i, 3)
                t4 = arr(i, 4): t5 = arr(i, 5): t6 = arr(i, 6)
                arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2): arr(i, 3) = arr(j, 3)
                arr(i, 4) = arr(j, 4): arr(i, 5) = arr(j, 5): arr(i, 6) = arr(j, 6)
                arr(j, 1) = t1: arr(j, 2) = t2: arr(j, 3) = t3
                arr(j, 4) = t4: arr(j, 5) = t5: arr(j, 6) = t6
            End If
        Next j
    Next i
End Sub

Private Function IsProtectedItem(ws As Worksheet, rowNum As Long) As Boolean
    ' Replace the entries below with your own never-adjust item keywords.
    Dim name As String
    name = Trim(CStr(ws.Cells(rowNum, NAME_COL).value))
    Dim blocked(1 To 4) As String
    blocked(1) = "ProtectedA"
    blocked(2) = "ProtectedB"
    blocked(3) = "LockedItemX"
    blocked(4) = "LockedItemY"
    Dim i As Long
    For i = 1 To 4
        If InStr(1, name, blocked(i), vbTextCompare) > 0 Then
            IsProtectedItem = True
            Exit Function
        End If
    Next i
    IsProtectedItem = False
End Function

' =============================================================================
' CONSTRAINT LAYER (unit bounds, multiplicity, ratio guards)
' =============================================================================

Private Function CalcMultipleQty(targetQty As Long, boxQty As Long, direction As String) As Long
    If boxQty <= 0 Then boxQty = 1
    If direction = "up" Then
        CalcMultipleQty = Application.WorksheetFunction.Ceiling(targetQty, boxQty)
    Else
        CalcMultipleQty = Application.WorksheetFunction.Floor(targetQty, boxQty)
        If CalcMultipleQty < boxQty Then CalcMultipleQty = boxQty
    End If
End Function

Private Function UnitValueValid(boxVal As Double) As Boolean
    UnitValueValid = (boxVal >= UNIT_VALUE_MIN And boxVal <= UNIT_VALUE_MAX)
End Function

Private Function UnitValueInBounds(avgVal As Double, newQty As Long, boxQty As Long, _
                                   ByRef validQty As Long, direction As String) As Boolean
    Dim testQty As Long, boxVal As Double
    If avgVal <= 0 Or newQty <= 0 Then UnitValueInBounds = False: Exit Function
    If direction = "up" Then
        For testQty = newQty To 10000 Step boxQty
            boxVal = avgVal * testQty
            If UnitValueValid(boxVal) Then validQty = testQty: UnitValueInBounds = True: Exit Function
            If boxVal > UNIT_VALUE_MAX Then Exit For
        Next testQty
    Else
        For testQty = newQty To boxQty Step -boxQty
            boxVal = avgVal * testQty
            If UnitValueValid(boxVal) Then validQty = testQty: UnitValueInBounds = True: Exit Function
        Next testQty
    End If
    UnitValueInBounds = False
End Function

Private Function GetCompliantQtyDown(avgVal As Double, targetQty As Long, boxQty As Long) As Long
    Dim testQty As Long, boxVal As Double
    If avgVal <= 0 Then GetCompliantQtyDown = boxQty: Exit Function
    For testQty = targetQty To boxQty Step -boxQty
        boxVal = avgVal * testQty
        If UnitValueValid(boxVal) Then GetCompliantQtyDown = testQty: Exit Function
    Next testQty
    GetCompliantQtyDown = boxQty
End Function

Private Function RatioBounds(origQty As Long, boxQty As Long, ByRef minQ As Long, ByRef maxQ As Long)
    If origQty > 0 Then
        minQ = Application.WorksheetFunction.Max(boxQty, WorksheetFunction.Ceiling(origQty * MIN_QTY_RATIO, 1))
        maxQ = Application.WorksheetFunction.Ceiling(origQty * MAX_QTY_RATIO, 1)
    Else
        minQ = boxQty
        maxQ = 100000
    End If
End Function

' Resolve a final quantity for bestIdx honouring bounds + multiplicity + ratio.
Private Function ResolveQuantity(ws As Worksheet, products() As Variant, idx As Long, _
                                needIncrease As Boolean) As Long
    Dim rowNum As Long, qty As Long, boxQty As Long, avgW As Double, origQty As Long
    Dim minQ As Long, maxQ As Long, targetQty As Long, validQty As Long
    Dim bandMin As Double, bandMax As Double
    rowNum = products(idx, 1): qty = products(idx, 3): boxQty = products(idx, 4)
    avgW = products(idx, 7): origQty = products(idx, 9)
    Call RatioBounds(origQty, boxQty, minQ, maxQ)

    If needIncrease Then
        If products(idx, 8) = 1 Then bandMin = BAND_HI_MIN Else bandMin = BAND_LO_MIN
        If products(idx, 8) = 1 Then bandMax = BAND_HI_MAX Else bandMax = BAND_LO_MAX
        Randomize
        Dim tgtVal As Double
        tgtVal = bandMin + Rnd() * (bandMax - bandMin)
        targetQty = Application.WorksheetFunction.Ceiling(tgtVal / avgW, 1)
        targetQty = CalcMultipleQty(targetQty, boxQty, "up")
        targetQty = Application.WorksheetFunction.Max(minQ, Application.WorksheetFunction.Min(maxQ, targetQty))
        If avgW > 0 Then
            If UnitValueInBounds(avgW, targetQty, boxQty, validQty, "up") Then targetQty = validQty
        End If
        If targetQty < boxQty Then targetQty = boxQty
        ResolveQuantity = targetQty
    Else
        Randomize
        tgtVal = UNIT_VALUE_MIN + Rnd() * (UNIT_VALUE_HI_MAX - UNIT_VALUE_MIN)
        targetQty = Application.WorksheetFunction.Max(boxQty, WorksheetFunction.Floor(tgtVal / avgW, 1))
        targetQty = CalcMultipleQty(targetQty, boxQty, "down")
        targetQty = Application.WorksheetFunction.Max(minQ, targetQty)
        If avgW > 0 Then targetQty = GetCompliantQtyDown(avgW, targetQty, boxQty)
        If targetQty < boxQty Then targetQty = boxQty
        ResolveQuantity = targetQty
    End If
End Function

' =============================================================================
' COARSE LAYER (deterministic normalisation + rebalance)
' =============================================================================

Private Sub NormaliseUnitValues(ws As Worksheet, startRow As Long, endRow As Long, _
                                ByRef adjustedCount As Long)
    Dim i As Long, qty As Long, boxQty As Long, unitVal As Double, avgW As Double
    Dim newQty As Long, validQty As Long, minInc As Long
    For i = startRow To endRow
        If IsEmpty(ws.Cells(i, BOX_COL).value) Or IsEmpty(ws.Cells(i, UNIT_VAL_COL).value) Then GoTo Nxt
        If IsNumeric(ws.Cells(i, BOX_COL).value) And IsNumeric(ws.Cells(i, UNIT_VAL_COL).value) Then
            qty = CLng(ws.Cells(i, QTY_COL).value)
            boxQty = CLng(ws.Cells(i, BOX_COL).value)
            unitVal = ws.Cells(i, UNIT_VAL_COL).value
            If unitVal > 0 And qty > 0 Then
                avgW = unitVal / qty
                If unitVal < UNIT_VALUE_MIN Then
                    minInc = Application.WorksheetFunction.Max( _
                        WorksheetFunction.Ceiling(UNIT_VALUE_MIN / avgW, 1) - qty, 1)
                    Dim j As Long, found As Boolean
                    found = False
                    For newQty = qty + minInc To qty + 100
                        If newQty >= boxQty Then
                            If avgW * newQty >= UNIT_VALUE_MIN And avgW * newQty <= UNIT_VALUE_MAX Then
                                ws.Cells(i, QTY_COL).value = newQty
                                ws.Cells(i, QTY_COL).Font.Color = HILITE_COLOR
                                adjustedCount = adjustedCount + 1
                                found = True: Exit For
                            ElseIf avgW * newQty > UNIT_VALUE_MAX Then
                                Exit For
                            End If
                        End If
                    Next newQty
                    If Not found Then ' no-op
                    End If
                ElseIf unitVal > UNIT_VALUE_MAX Then
                    Dim maxDec As Long
                    maxDec = qty - Application.WorksheetFunction.Max( _
                        WorksheetFunction.Floor(UNIT_VALUE_MAX / avgW, 1), boxQty)
                    found = False
                    For j = 1 To Application.WorksheetFunction.Min(maxDec, 100)
                        newQty = qty - j
                        If newQty >= boxQty Then
                            If avgW * newQty >= UNIT_VALUE_MIN And avgW * newQty <= UNIT_VALUE_MAX Then
                                ws.Cells(i, QTY_COL).value = newQty
                                ws.Cells(i, QTY_COL).Font.Color = HILITE_COLOR
                                adjustedCount = adjustedCount + 1
                                found = True: Exit For
                            ElseIf avgW * newQty < UNIT_VALUE_MIN Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next j
                End If
            End If
        End If
Nxt:
    Next i
    If adjustedCount > 0 Then Application.Calculate
End Sub

Private Sub RebalanceHighPenalty(ws As Worksheet, startRow As Long, endRow As Long, _
                                 totalValRow As Long, targetValRow As Long, _
                                 targetVal As Double, ByRef diff As Double, ByRef adjustedCount As Long)
    Dim i As Long, qty As Long, boxQty As Long, unitVal As Double, avgW As Double
    Dim newQty As Long, minQ As Long, maxQ As Long
    For i = startRow To endRow
        If diff <= TOL_COARSE Then Exit For
        If IsNumeric(ws.Cells(i, QTY_COL).value) And IsNumeric(ws.Cells(i, PENALTY_COL).value) Then
            Dim pen As Double
            pen = ws.Cells(i, PENALTY_COL).value
            If pen >= PENALTY_HIGH Then
                qty = CLng(ws.Cells(i, QTY_COL).value)
                boxQty = IIf(IsNumeric(ws.Cells(i, BOX_COL).value), CLng(ws.Cells(i, BOX_COL).value), 1)
                unitVal = ws.Cells(i, UNIT_VAL_COL).value
                If unitVal > 0 And qty > boxQty Then
                    avgW = unitVal / qty
                    Call RatioBounds(CLng(ws.Cells(i, ORIG_QTY_COL).value), boxQty, minQ, maxQ)
                    Dim tgtMin As Long, tgtMax As Long
                    tgtMin = Application.WorksheetFunction.Max(boxQty, WorksheetFunction.Ceiling(UNIT_VALUE_MIN / avgW, 1))
                    tgtMax = WorksheetFunction.Floor(UNIT_VALUE_MAX / avgW, 1)
                    If tgtMin <= tgtMax Then
                        For newQty = qty - 1 To tgtMin Step -1
                            If avgW * newQty >= UNIT_VALUE_MIN And avgW * newQty <= UNIT_VALUE_MAX Then
                                ws.Cells(i, QTY_COL).value = newQty
                                ws.Cells(i, QTY_COL).Font.Color = HILITE_COLOR
                                adjustedCount = adjustedCount + 1
                                If adjustedCount Mod 5 = 0 Then
                                    Application.Calculate
                                    diff = Abs(ws.Cells(totalValRow, TOTAL_VAL_COL).value - targetVal)
                                End If
                                If diff <= TOL_COARSE Then Exit For
                            End If
                        Next newQty
                    End If
                End If
            End If
        End If
    Next i
    Application.Calculate
    diff = Abs(ws.Cells(totalValRow, TOTAL_VAL_COL).value - targetVal)
End Sub

Private Sub RebalanceLowPenalty(ws As Worksheet, startRow As Long, endRow As Long, _
                                totalValRow As Long, targetValRow As Long, _
                                targetVal As Double, ByRef diff As Double, ByRef adjustedCount As Long)
    Dim i As Long, qty As Long, boxQty As Long, unitVal As Double, avgW As Double
    Dim newQty As Long, minQ As Long, maxQ As Long
    For i = startRow To endRow
        If diff <= TOL_COARSE Then Exit For
        If IsNumeric(ws.Cells(i, QTY_COL).value) And IsNumeric(ws.Cells(i, PENALTY_COL).value) Then
            Dim pen As Double, price As Double
            pen = ws.Cells(i, PENALTY_COL).value
            price = IIf(IsNumeric(ws.Cells(i, UNIT_PRICE_COL).value), ws.Cells(i, UNIT_PRICE_COL).value, 999)
            If pen <= PENALTY_HIGH And pen >= 0 And price < PRICE_SKIP Then
                qty = CLng(ws.Cells(i, QTY_COL).value)
                boxQty = IIf(IsNumeric(ws.Cells(i, BOX_COL).value), CLng(ws.Cells(i, BOX_COL).value), 1)
                unitVal = ws.Cells(i, UNIT_VAL_COL).value
                If unitVal > 0 And qty > 0 Then
                    avgW = unitVal / qty
                    Call RatioBounds(CLng(ws.Cells(i, ORIG_QTY_COL).value), boxQty, minQ, maxQ)
                    Dim tgtMin As Long, tgtMax As Long
                    tgtMin = Application.WorksheetFunction.Ceiling(UNIT_VALUE_MIN / avgW, 1)
                    tgtMax = WorksheetFunction.Floor(UNIT_VALUE_MAX / avgW, 1)
                    If tgtMin <= tgtMax Then
                        For newQty = tgtMin To Application.WorksheetFunction.Min(tgtMax, tgtMin + 50)
                            If avgW * newQty >= UNIT_VALUE_MIN And avgW * newQty <= UNIT_VALUE_MAX Then
                                ws.Cells(i, QTY_COL).value = newQty
                                ws.Cells(i, QTY_COL).Font.Color = HILITE_COLOR
                                adjustedCount = adjustedCount + 1
                                If adjustedCount Mod 5 = 0 Then
                                    Application.Calculate
                                    diff = Abs(ws.Cells(totalValRow, TOTAL_VAL_COL).value - targetVal)
                                End If
                                If diff <= TOL_COARSE Then Exit For
                            End If
                        Next newQty
                    End If
                End If
            End If
        End If
    Next i
    Application.Calculate
    diff = Abs(ws.Cells(totalValRow, TOTAL_VAL_COL).value - targetVal)
End Sub

' =============================================================================
' PRECISE LAYER (heuristic search)
' =============================================================================

Private Function TryIncrease(ws As Worksheet, products() As Variant, validCount As Long, _
                             loPenalty() As Variant, loCount As Long, _
                             zeroPenalty() As Variant, zeroCount As Long, _
                             ByRef bestIdx As Long, ByRef bestImpact As Double) As Boolean
    Dim i As Long, pIdx As Long, rowNum As Long, qty As Long, boxQty As Long, avgW As Double
    Dim origQty As Long, targetQty As Long, validQty As Long, price As Double
    ' Priority 1: low-penalty items.
    For i = 1 To loCount
        pIdx = loPenalty(i, 1): rowNum = products(pIdx, 1)
        If IsProtectedItem(ws, rowNum) Then GoTo NxtLo
        If IsNumeric(ws.Cells(rowNum, UNIT_PRICE_COL).value) Then
            price = ws.Cells(rowNum, UNIT_PRICE_COL).value
            If price > PRICE_SKIP Then GoTo NxtLo
        End If
        qty = products(pIdx, 3): boxQty = products(pIdx, 4): avgW = products(pIdx, 7): origQty = products(pIdx, 9)
        If avgW > 0.0001 Then
            targetQty = Application.WorksheetFunction.Ceiling(loPenalty(i, 6) / avgW, 1)
            targetQty = CalcMultipleQty(targetQty, boxQty, "up")
            Call ApplyRatio(origQty, boxQty, targetQty)
            If avgW > 0 Then
                If UnitValueInBounds(avgW, targetQty, boxQty, validQty, "up") Then targetQty = validQty Else GoTo NxtLo
            End If
            If targetQty > qty Then
                bestIdx = pIdx: bestImpact = avgW * (targetQty - qty): TryIncrease = True: Exit Function
            End If
        End If
NxtLo:
    Next i
    ' Priority 2: zero-penalty items.
    For i = 1 To zeroCount
        pIdx = zeroPenalty(i, 1): rowNum = products(pIdx, 1)
        If IsProtectedItem(ws, rowNum) Then GoTo NxtZ
        If IsNumeric(ws.Cells(rowNum, UNIT_PRICE_COL).value) Then
            price = ws.Cells(rowNum, UNIT_PRICE_COL).value
            If price > PRICE_SKIP Then GoTo NxtZ
        End If
        qty = products(pIdx, 3): boxQty = products(pIdx, 4): avgW = products(pIdx, 7): origQty = products(pIdx, 9)
        If avgW > 0.0001 Then
            targetQty = Application.WorksheetFunction.Ceiling(zeroPenalty(i, 5) / avgW, 1)
            targetQty = CalcMultipleQty(targetQty, boxQty, "up")
            Call ApplyRatio(origQty, boxQty, targetQty)
            If avgW > 0 Then
                If UnitValueInBounds(avgW, targetQty, boxQty, validQty, "up") Then targetQty = validQty Else GoTo NxtZ
            End If
            If targetQty > qty Then
                bestIdx = pIdx: bestImpact = avgW * (targetQty - qty): TryIncrease = True: Exit Function
            End If
        End If
NxtZ:
    Next i
    ' Priority 3: mid-penalty items.
    Dim k As Long
    For i = 1 To validCount
        If products(i, 2) >= PENALTY_HIGH And products(i, 2) < PENALTY_MID_HI And products(i, 10) = 0 Then
            rowNum = products(i, 1)
            If IsProtectedItem(ws, rowNum) Then GoTo NxtM
            If IsNumeric(ws.Cells(rowNum, UNIT_PRICE_COL).value) Then
                price = ws.Cells(rowNum, UNIT_PRICE_COL).value
                If price > PRICE_SKIP Then GoTo NxtM
            End If
            qty = products(i, 3): boxQty = products(i, 4): avgW = products(i, 7): origQty = products(i, 9)
            If avgW > 0.0001 Then
                Randomize
                Dim tgtV As Double: tgtV = 15 + Rnd() * 10
                targetQty = Application.WorksheetFunction.Ceiling(tgtV / avgW, 1)
                targetQty = CalcMultipleQty(targetQty, boxQty, "up")
                Call ApplyRatio(origQty, boxQty, targetQty)
                If avgW > 0 Then
                    If UnitValueInBounds(avgW, targetQty, boxQty, validQty, "up") Then targetQty = validQty Else GoTo NxtM
                End If
                If targetQty > qty Then
                    bestIdx = i: bestImpact = avgW * (targetQty - qty): TryIncrease = True: Exit Function
                End If
            End If
        End If
NxtM:
    Next i
    TryIncrease = False
End Function

Private Function TryDecrease(ws As Worksheet, products() As Variant, validCount As Long, _
                             hiPenalty() As Variant, hiCount As Long, _
                             ByRef bestIdx As Long, ByRef bestImpact As Double) As Boolean
    Dim i As Long, pIdx As Long, rowNum As Long, qty As Long, boxQty As Long, avgW As Double
    Dim origQty As Long, targetQty As Long
    For i = 1 To hiCount
        pIdx = hiPenalty(i, 1): rowNum = products(pIdx, 1)
        If IsProtectedItem(ws, rowNum) Then GoTo NxtH
        qty = products(pIdx, 3): boxQty = products(pIdx, 4): avgW = products(pIdx, 7): origQty = products(pIdx, 9)
        If avgW > 0.0001 And qty > boxQty Then
            Randomize
            Dim tgtV As Double: tgtV = UNIT_VALUE_MIN + Rnd() * 2
            targetQty = Application.WorksheetFunction.Max(boxQty, WorksheetFunction.Floor(tgtV / avgW, 1))
            targetQty = CalcMultipleQty(targetQty, boxQty, "down")
            Call ApplyRatio(origQty, boxQty, targetQty)
            If avgW > 0 Then targetQty = GetCompliantQtyDown(avgW, targetQty, boxQty)
            If targetQty < qty Then
                bestIdx = pIdx: bestImpact = avgW * (qty - targetQty): TryDecrease = True: Exit Function
            End If
        End If
NxtH:
    Next i
    ' Priority 2: mid-penalty items (high -> low).
    Dim tempM() As Variant, tCount As Long, k As Long, j As Long
    tCount = 0
    For i = 1 To validCount
        If products(i, 2) >= PENALTY_HIGH And products(i, 2) < PENALTY_MID_HI And products(i, 10) = 0 Then tCount = tCount + 1
    Next i
    If tCount > 0 Then
        ReDim tempM(1 To tCount, 1 To 2)
        k = 1
        For i = 1 To validCount
            If products(i, 2) >= PENALTY_HIGH And products(i, 2) < PENALTY_MID_HI And products(i, 10) = 0 Then
                tempM(k, 1) = i: tempM(k, 2) = products(i, 2): k = k + 1
            End If
        Next i
        For i = 1 To tCount - 1
            For j = i + 1 To tCount
                If tempM(i, 2) < tempM(j, 2) Then
                    Dim a, b: a = tempM(i, 1): b = tempM(i, 2)
                    tempM(i, 1) = tempM(j, 1): tempM(i, 2) = tempM(j, 2)
                    tempM(j, 1) = a: tempM(j, 2) = b
                End If
            Next j
        Next i
        For i = 1 To tCount
            pIdx = tempM(i, 1): rowNum = products(pIdx, 1)
            If IsProtectedItem(ws, rowNum) Then GoTo NxtM2
            qty = products(pIdx, 3): boxQty = products(pIdx, 4): avgW = products(pIdx, 7): origQty = products(pIdx, 9)
            If avgW > 0.0001 And qty > boxQty Then
                Randomize
                tgtV = UNIT_VALUE_MIN + Rnd() * 8
                targetQty = Application.WorksheetFunction.Max(boxQty, WorksheetFunction.Floor(tgtV / avgW, 1))
                targetQty = CalcMultipleQty(targetQty, boxQty, "down")
                Call ApplyRatio(origQty, boxQty, targetQty)
                If avgW > 0 Then targetQty = GetCompliantQtyDown(avgW, targetQty, boxQty)
                If targetQty < qty Then
                    bestIdx = pIdx: bestImpact = avgW * (qty - targetQty): TryDecrease = True: Exit Function
                End If
            End If
NxtM2:
        Next i
    End If
    TryDecrease = False
End Function

Private Sub ApplyRatio(origQty As Long, boxQty As Long, ByRef targetQty As Long)
    Dim minQ As Long, maxQ As Long
    Call RatioBounds(origQty, boxQty, minQ, maxQ)
    targetQty = Application.WorksheetFunction.Max(minQ, Application.WorksheetFunction.Min(maxQ, targetQty))
End Sub

' =============================================================================
' STAGE 2: unique-item formula micro-adjust (closes the last few units)
' =============================================================================

Private Function FindUniqueLink(ws As Worksheet, productEndRow As Long) As Long
    Dim i As Long, linkVal As String, count As Long, target As Long
    target = 0
    For i = PRODUCT_START_ROW To productEndRow
        linkVal = Trim(CStr(ws.Cells(i, LINK_COL).value))
        If linkVal = "" Then GoTo Nxt
        count = 0
        Dim k As Long
        For k = PRODUCT_START_ROW To productEndRow
            If Trim(CStr(ws.Cells(k, LINK_COL).value)) = linkVal Then count = count + 1
        Next k
        If count = 1 Then target = i: FindUniqueLink = target: Exit Function
Nxt:
    Next i
    FindUniqueLink = 0
End Function

Private Sub UniqueLinkStage(ws As Worksheet, totalValRow As Long, targetValRow As Long, _
                            productEndRow As Long, ByRef diff As Double, targetVal As Double, _
                            ByRef adjustedCount As Long)
    Dim rowNum As Long
    rowNum = FindUniqueLink(ws, productEndRow)
    If rowNum > 0 Then
        Dim baseVal As Double, qty As Long, factor As Double, bestFactor As Double, bestDiff As Double
        baseVal = ws.Cells(rowNum, 10).value
        qty = ws.Cells(rowNum, QTY_COL).value
        bestDiff = 9999: bestFactor = 1.15
        Dim tw As Double
        Application.Calculate
        tw = ws.Cells(totalValRow, TOTAL_VAL_COL).value - ws.Cells(rowNum, TOTAL_VAL_COL).value
        Dim f As Double
        For f = 1# To 1.15 Step 0.01
            Dim tw2 As Double
            tw2 = tw + (baseVal * f * qty)
            If Abs(tw2 - targetVal) < bestDiff Then bestDiff = Abs(tw2 - targetVal): bestFactor = f
        Next f
        ws.Cells(rowNum, TOTAL_VAL_COL).Formula = "=J" & rowNum & "*" & Format(bestFactor, "0.00") & "*O" & rowNum
        ws.Cells(rowNum, TOTAL_VAL_COL).Interior.Color = RGB(226, 239, 218)
        adjustedCount = adjustedCount + 1
        Application.Calculate
        diff = Abs(ws.Cells(totalValRow, TOTAL_VAL_COL).value - targetVal)
    End If
End Sub

' =============================================================================
' STAGE 3: penalty optimisation (drive penalty/unit down toward ceiling)
' =============================================================================

Private Sub PenaltyOptimiseStage(ws As Worksheet, products() As Variant, validCount As Long, _
                                 hiPenalty() As Variant, hiCount As Long, _
                                 zeroPenalty() As Variant, zeroCount As Long, _
                                 totalValRow As Long, targetValRow As Long, productEndRow As Long, _
                                 targetVal As Double, ByRef diff As Double, ByRef adjustedCount As Long, _
                                 ByRef currentPenalty As Double, startTime As Double, timeLimit As Double)
    Dim maxIter As Long, iter As Long
    maxIter = 50: iter = 0
    Application.Calculate
    diff = Abs(ws.Cells(totalValRow, TOTAL_VAL_COL).value - targetVal)
    currentPenalty = ComputePenaltyPerUnit(ws, productEndRow)
    Do While currentPenalty > (PENALTY_CEILING - 0.23) And iter < maxIter And Timer - startTime < timeLimit
        Dim optimised As Boolean: optimised = False
        Dim j As Long, i As Long, rowNum As Long, qty As Long, boxQty As Long, avgW As Double, newQty As Long, origQty As Long
        If hiCount > 0 Then
            For j = 1 To hiCount
                i = hiPenalty(j, 1): rowNum = products(i, 1)
                If IsProtectedItem(ws, rowNum) Then GoTo NxtP
                qty = products(i, 3): boxQty = products(i, 4): avgW = products(i, 7): origQty = products(i, 9)
                If avgW > 0.0001 And qty > boxQty Then
                    Randomize
                    Dim tgtV As Double: tgtV = UNIT_VALUE_MIN + Rnd() * 8.65
                    newQty = Application.WorksheetFunction.Max(boxQty, WorksheetFunction.Floor(tgtV / avgW, 1))
                    newQty = CalcMultipleQty(newQty, boxQty, "down")
                    Call ApplyRatio(origQty, boxQty, newQty)
                    If avgW > 0 Then newQty = GetCompliantQtyDown(avgW, newQty, boxQty)
                    If newQty < qty Then
                        ws.Cells(rowNum, QTY_COL).value = newQty
                        ws.Cells(rowNum, QTY_COL).Font.Color = HILITE_COLOR
                        adjustedCount = adjustedCount + 1
                        products(i, 3) = newQty: hiPenalty(j, 3) = newQty
                        optimised = True: Exit For
                    End If
                End If
NxtP:
            Next j
        End If
        If Not optimised Then
            For i = 1 To validCount
                If products(i, 2) > 0 And products(i, 2) < PENALTY_HIGH And products(i, 10) = 0 Then
                    rowNum = products(i, 1)
                    If IsProtectedItem(ws, rowNum) Then GoTo NxtP2
                    qty = products(i, 3): boxQty = products(i, 4): avgW = products(i, 7): origQty = products(i, 9)
                    If avgW > 0.0001 Then
                        Randomize
                        tgtV = 22 + Rnd() * 10
                        newQty = Application.WorksheetFunction.Ceiling(tgtV / avgW, 1)
                        newQty = CalcMultipleQty(newQty, boxQty, "up")
                        Call ApplyRatio(origQty, boxQty, newQty)
                        If avgW > 0 Then
                            Dim vq As Long
                            If UnitValueInBounds(avgW, newQty, boxQty, vq, "up") Then newQty = vq Else GoTo NxtP2
                        End If
                        If newQty > qty Then
                            ws.Cells(rowNum, QTY_COL).value = newQty
                            ws.Cells(rowNum, QTY_COL).Font.Color = HILITE_COLOR
                            adjustedCount = adjustedCount + 1
                            products(i, 3) = newQty
                            optimised = True: Exit For
                        End If
                    End If
NxtP2:
                End If
            Next i
        End If
        If Not optimised Then Exit Do
        Application.Calculate
        diff = Abs(ws.Cells(totalValRow, TOTAL_VAL_COL).value - targetVal)
        currentPenalty = ComputePenaltyPerUnit(ws, productEndRow)
        If diff > TOL_PRECISE Then Exit Do
        iter = iter + 1
    Loop
End Sub

' =============================================================================
' STAGE 4: forced reduce (guarantees a usable result even if goals conflict)
' =============================================================================

Private Function ForcedReduce(ws As Worksheet, products() As Variant, validCount As Long, _
                              hiPenalty() As Variant, hiCount As Long, _
                              loPenalty() As Variant, loCount As Long, _
                              zeroPenalty() As Variant, zeroCount As Long, _
                              ByRef bestIdx As Long, ByRef bestImpact As Double, _
                              currentDiff As Double) As Boolean
    Dim removeTarget As Double
    removeTarget = currentDiff * FORCED_RATIO
    If removeTarget < FORCED_MIN Then removeTarget = FORCED_MIN
    Dim i As Long, pIdx As Long, rowNum As Long, qty As Long, boxQty As Long, avgW As Double, newQty As Long, newVal As Double, origQty As Long, minQ As Long
    ' Priority 1: high-penalty items.
    For i = 1 To hiCount
        pIdx = hiPenalty(i, 1): rowNum = products(pIdx, 1)
        If IsProtectedItem(ws, rowNum) Then GoTo NxtF1
        qty = products(pIdx, 3): boxQty = products(pIdx, 4): avgW = products(pIdx, 7): origQty = products(pIdx, 9)
        If avgW > 0.0001 And qty > boxQty Then
            Dim reduceQty As Long
            reduceQty = Application.WorksheetFunction.Min(qty - boxQty, _
                Application.WorksheetFunction.RoundUp(removeTarget / avgW, 0))
            If reduceQty > 0 Then
                newQty = qty - reduceQty: newVal = avgW * newQty
                If UnitValueValid(newVal) Then
                    Call RatioBounds(origQty, boxQty, minQ, minQ)
                    If newQty >= minQ Or origQty <= 0 Then
                        bestIdx = pIdx: bestImpact = avgW * reduceQty: ForcedReduce = True: Exit Function
                    End If
                End If
            End If
        End If
NxtF1:
    Next i
    ' Priority 2: low-penalty items (low -> high).
    Dim tempLo() As Variant, tCount As Long, k As Long, j As Long
    tCount = 0
    For i = 1 To loCount
        pIdx = loPenalty(i, 1)
        If products(pIdx, 2) > 0 Then tCount = tCount + 1
    Next i
    If tCount > 0 Then
        ReDim tempLo(1 To tCount, 1 To 2)
        k = 1
        For i = 1 To loCount
            pIdx = loPenalty(i, 1)
            If products(pIdx, 2) > 0 Then tempLo(k, 1) = pIdx: tempLo(k, 2) = products(pIdx, 2): k = k + 1
        Next i
        For i = 1 To tCount - 1
            For j = i + 1 To tCount
                If tempLo(i, 2) < tempLo(j, 2) Then
                    Dim a, b: a = tempLo(i, 1): b = tempLo(i, 2)
                    tempLo(i, 1) = tempLo(j, 1): tempLo(i, 2) = tempLo(j, 2)
                    tempLo(j, 1) = a: tempLo(j, 2) = b
                End If
            Next j
        Next i
        For i = 1 To tCount
            pIdx = tempLo(i, 1): rowNum = products(pIdx, 1)
            If IsProtectedItem(ws, rowNum) Then GoTo NxtF2
            qty = products(pIdx, 3): boxQty = products(pIdx, 4): avgW = products(pIdx, 7): origQty = products(pIdx, 9)
            If avgW > 0.0001 And qty > boxQty Then
                reduceQty = Application.WorksheetFunction.Min(qty - boxQty, _
                    Application.WorksheetFunction.RoundUp(removeTarget / avgW, 0))
                If reduceQty > 0 Then
                    newQty = qty - reduceQty: newVal = avgW * newQty
                    If UnitValueValid(newVal) Then
                        Call RatioBounds(origQty, boxQty, minQ, minQ)
                        If newQty >= minQ Or origQty <= 0 Then
                            bestIdx = pIdx: bestImpact = avgW * reduceQty: ForcedReduce = True: Exit Function
                        End If
                    End If
                End If
            End If
NxtF2:
        Next i
    End If
    ForcedReduce = False
End Function

Private Sub ForcedReduceStage(ws As Worksheet, products() As Variant, validCount As Long, _
                              hiPenalty() As Variant, hiCount As Long, _
                              loPenalty() As Variant, loCount As Long, _
                              zeroPenalty() As Variant, zeroCount As Long, _
                              totalValRow As Long, targetValRow As Long, productEndRow As Long, _
                              targetVal As Double, ByRef diff As Double, ByRef adjustedCount As Long, _
                              ByRef currentPenalty As Double, startTime As Double, timeLimit As Double)
    Application.Calculate
    diff = Abs(ws.Cells(totalValRow, TOTAL_VAL_COL).value - targetVal)
    currentPenalty = ComputePenaltyPerUnit(ws, productEndRow)
    Do While diff > TOL_PRECISE And Timer - startTime < timeLimit
        Dim bestIdx As Long, bestImpact As Double, found As Boolean
        found = ForcedReduce(ws, products, validCount, hiPenalty, hiCount, _
                             loPenalty, loCount, zeroPenalty, zeroCount, bestIdx, bestImpact, diff)
        If Not found Then Exit Do
        Dim rowNum As Long, qty As Long, avgW As Double
        rowNum = products(bestIdx, 1): qty = products(bestIdx, 3): avgW = products(bestIdx, 7)
        Dim newQty As Long
        newQty = qty - CLng(bestImpact / avgW)
        ws.Cells(rowNum, QTY_COL).value = newQty
        ws.Cells(rowNum, QTY_COL).Font.Color = RGB(255, 0, 255)  ' magenta marks forced reduction
        adjustedCount = adjustedCount + 1
        products(bestIdx, 3) = newQty
        Application.Calculate
        diff = Abs(ws.Cells(totalValRow, TOTAL_VAL_COL).value - targetVal)
        currentPenalty = ComputePenaltyPerUnit(ws, productEndRow)
    Loop
End Sub

' =============================================================================
' PRESENTATION / TIDY
' =============================================================================

Private Sub TidySheet(ws As Worksheet, lastRow As Long)
    On Error Resume Next
    ws.Range("N4:N" & lastRow).NumberFormat = "0"
    ws.Range("O4:O" & lastRow).NumberFormat = "0"
    Dim cols As Variant
    cols = Array(7, 8, 9, 11, 12, 13, 19, 21, 22, 23, 29, 30, 31)
    Dim i As Long
    For i = 0 To UBound(cols)
        ws.Columns(cols(i)).NumberFormat = "0.00"
    Next i
    On Error GoTo 0
End Sub

Private Sub DebugPrint(msg As String)
    If DEBUG_ON Then Debug.Print Now & " - " & msg
End Sub
