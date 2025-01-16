' ------------------------------------------
' Title: Simulation for Validating IVN-ISM and Fuzzy ISM
' Author: [Your Name]
' Description: This VBA code conducts a simulation to compare IVN-ISM and Fuzzy ISM
' methodologies using the Dice-Sørensen similarity index.
' ------------------------------------------

' Type definitions for Interval-Valued Neutrosophic (IVN) and Fuzzy numbers
Type Bound
    U As Double ' Upper bound
    L As Double ' Lower bound
End Type

Type IVN
    Tr As Bound ' Truth membership
    In As Bound ' Indeterminacy membership
    Fa As Bound ' Falsity membership
End Type

Type SF
    Mf As Double ' Membership function
    NMf As Double ' Non-membership function
    Hf As Double ' Hesitancy degree
End Type

Sub IVN_ISM_Simulation()

    ' Request number of replications from the user
    Dim r As Integer
    r = InputBox("Enter number of replications")

    ' Initialize arrays for storing results
    Dim DSS() As Double
    ReDim DSS(1 To r)

    ' Simulation loop
    Dim w As Integer
    For w = 1 To r
        Call RunSimulation(w, DSS)
    Next w

    ' Calculate average Dice-Sørensen similarity index
    Dim AveDSS As Double
    AveDSS = Application.WorksheetFunction.Average(DSS)

    ' Calculate standard deviation
    Dim SE As Double
    SE = Application.WorksheetFunction.StDev(DSS)

    ' Output results to Excel sheet
    Cells(2, 3).Value = "Average Dice-Sørensen Similarity Index"
    Cells(2, 4).Value = AveDSS
    Cells(3, 3).Value = "Standard Deviation"
    Cells(3, 4).Value = SE

    MsgBox "Simulation completed! Results saved to the worksheet."

End Sub

Sub RunSimulation(w As Integer, DSS() As Double)

    ' Randomly determine the number of factors (x) and experts (z)
    Dim x As Integer, z As Integer
    z = Int((20 - 5 + 1) * Rnd + 5) ' Random experts: 5 to 20
    x = Int((30 - 10 + 1) * Rnd + 10) ' Random factors: 10 to 30

    ' Initialize matrices
    Dim rndmat() As Integer, IVNExpert() As IVN, FExpert() As SF
    Dim IVNDec() As IVN, FDec() As SF
    Dim RM() As Double, FRM() As Double, BRM() As Double, FBRM() As Double

    ReDim rndmat(1 To x, 1 To x, 1 To z)
    ReDim IVNExpert(1 To x, 1 To x, 1 To z)
    ReDim FExpert(1 To x, 1 To x, 1 To z)
    ReDim IVNDec(1 To x, 1 To x)
    ReDim FDec(1 To x, 1 To x)
    ReDim RM(1 To x, 1 To x), FRM(1 To x, 1 To x)
    ReDim BRM(1 To x, 1 To x), FBRM(1 To x, 1 To x)

    ' Generate random expert matrices
    Call GenerateRandomMatrices(x, z, rndmat, IVNExpert, FExpert)

    ' Calculate decision matrices
    Call CalculateDecisionMatrices(x, z, IVNExpert, FExpert, IVNDec, FDec, RM, FRM)

    ' Generate binary reachability matrices and compute Dice-Sørensen similarity
    DSS(w) = ComputeDiceSorensen(x, RM, FRM, BRM, FBRM)

End Sub

Sub GenerateRandomMatrices(x As Integer, z As Integer, rndmat() As Integer, _
    IVNExpert() As IVN, FExpert() As SF)

    Dim i As Integer, j As Integer, t As Integer
    For t = 1 To z
        For i = 1 To x
            For j = 1 To x
                If i = j Then
                    rndmat(i, j, t) = 4
                Else
                    rndmat(i, j, t) = Int(5 * Rnd)
                End If

                Select Case rndmat(i, j, t)
                    Case 0
                        IVNExpert(i, j, t).Tr.L = 0
                        IVNExpert(i, j, t).Tr.U = 0
                        IVNExpert(i, j, t).In.L = 0
                        IVNExpert(i, j, t).In.U = 0
                        IVNExpert(i, j, t).Fa.L = 1
                        IVNExpert(i, j, t).Fa.U = 1
                        FExpert(i, j, t).Mf = 0
                        FExpert(i, j, t).NMf = 0
                        FExpert(i, j, t).Hf = 0.25
                    ' Additional cases for 1 to 4 go here...
                End Select
            Next j
        Next i
    Next t

End Sub

Sub CalculateDecisionMatrices(x As Integer, z As Integer, IVNExpert() As IVN, _
    FExpert() As SF, IVNDec() As IVN, FDec() As SF, RM() As Double, FRM() As Double)
    ' Add logic for aggregating decision matrices
End Sub

Function ComputeDiceSorensen(x As Integer, RM() As Double, FRM() As Double, _
    BRM() As Double, FBRM() As Double) As Double
    Dim i As Integer, j As Integer
    Dim JaccA As Double, JaccB As Double, Joint As Double

    JaccA = 0: JaccB = 0: Joint = 0
    For i = 1 To x
        For j = 1 To x
            If RM(i, j) > 0 Then BRM(i, j) = 1 Else BRM(i, j) = 0
            If FRM(i, j) > 0 Then FBRM(i, j) = 1 Else FBRM(i, j) = 0

            JaccA = JaccA + BRM(i, j)
            JaccB = JaccB + FBRM(i, j)
            If BRM(i, j) = FBRM(i, j) Then Joint = Joint + BRM(i, j)
        Next j
    Next i

    ComputeDiceSorensen = 2 * Joint / (JaccA + JaccB)
End Function
