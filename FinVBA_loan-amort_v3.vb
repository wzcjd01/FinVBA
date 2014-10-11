
Sub Loan_Amort_V3()

    'this program creates a year-by-year repayment schedule for a fixed rate 'loan that is to be repayed in equal annual installment payable at the end of 'each year. determine the required annual payment amount by iteration instead of using VBA's pmt function. use arrays for all intermediate calculations to speed up the model's operation

    Dim intRate As Single, initLoanAmnt As Single, loanLife As Single
    Dim yrBegBal As Single, yrEndBal As Single
    Dim annualPmnt As Single, annualIncr as single,numIterations as Integer
    dim intComp as single, princRepay As Single
    Dim outRow As Integer, rowNum As Integer, outSheet As String

    '**************************************************
    ' programmer inputs
    '**************************************************
    outRow = 5 'Used to control where the output table will start
    outSheet = "Loan Amort"

    Worksheets(outSheet).Activate

    'clear previous data
    Rows((outRow + 4) & ":" & (outRow + 300)).Select
    Selection.Clear
    Range("b2:b4").ClearContents

    '**************************************************
    ' get user inputs
    '**************************************************
    'the user provides these input data through dialog boxes. input data not
    'meeting specified criterea are not accepted

    Do
        intRate = InputBox("Enter interest rate in percent" _
        & "without % sign. It must be between 0 and 15")

        If intRate < 0 Or intRate > 15 Then
            MsgBox ("interest rate must be between 0% and 15%.")
        Else
            Exit Do
        End If
    Loop

    intRate = intRate / 100

    Do
        loanLife = InputBox("enter loan life in years. " _
        & "loan life must be a whole number.")
        If loanLife < 0 Or (loanLife - Round(loanLife)) <> 0 Then
            MsgBox ("loan life must be a whole number.")
        Else: Exit Do
        End If
    Loop

GetLoanAmnt:
    initLoanAmnt = InputBox("enter loan amount. Loan amount " _
    & "must be a positive whole number.")

    If initLoanAmnt < 0 Or (initLoanAmnt - Round(initLoanAmnt) <> 0) Then
        MsgBox ("loan amount must be a positve whole number.")
        GoTo GetLoanAmnt
    End If


    '**************************************************
    ' write out the input data on the output sheet
    '**************************************************
    Cells(2, 2).Value = intRate
    Cells(3, 2).Value = loanLife
    Cells(4, 2).Value = initLoanAmnt

    '**************************************************
    ' compute and output results
    '**************************************************
' set annual payment upfront with the vale of (initLoanAmnt/loanLife)
    annualPmnt = initLoanAmnt / loanLife 
    annualIncr = initLoanAmnt / 1000 'step size for payment guess
    numIterations = 0
    
    do

    'initialize beginning balance for year 1
    yrBegBal = initLoanAmnt

    'loop to calculate and output year-by-year amort. table
    For rowNum = 1 To loanLife
        intComp = yrBegBal * intRate
        princRepay = annualPmnt - intComp
        yrEndBal = yrBegBal - princRepay

        Cells(outRow + 3 + rowNum, 3).Value = rowNum  'year number
        Cells(outRow + 3 + rowNum, 4).Value = yrBegBal
        Cells(outRow + 3 + rowNum, 5).Value = annualPmnt
        Cells(outRow + 3 + rowNum, 6).Value = intComp
        Cells(outRow + 3 + rowNum, 7).Value = princRepay
        Cells(outRow + 3 + rowNum, 8).Value = yrEndBal

        yrBegBal = yrEndBal

    Next rowNum

    annualPmnt = annualPmnt + anualIncr
    nunOfIterations = numOfIterations + 1
loop while yrEndBal >= 0

'write out the number of iterations used
cells(outRow + loanLife + 4, 1).value = "no. of iterations ="
cells(outRow + loanLife + 4, 2).value = numOfIterations


    '**************************************************
    ' format the output data in the table
    '**************************************************
    Range(Cells(outRow + 4, 4), Cells(outRow + 3 + loanLife, 8)) _
    .Select
    Selection.NumberFormat = "$#,##0.00"

End Sub


'/* vim: set filetype=vb : */



