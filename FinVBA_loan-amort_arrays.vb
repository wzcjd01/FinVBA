option base 1
Sub Loan_Amort_Arrays()

    'this program creates a year-by-year repayment schedule for a fixed rate 'loan that is to be repayed in equal annual installment payable at the end of 'each year. determine the required annual payment amount by iteration instead of using VBA's pmt function. use arrays for all intermediate calculations to speed up the model's operation

    Dim intRate As Single, initLoanAmnt As Single, loanLife As Single
    Dim yrBegBal(100) As Single, yrEndBal() As Single, finalBal as single
    dim ipPay(1 to 100, 1 to 2) as single
    Dim annualPmnt As Single, aPmtOld
    dim numOfIterations as Integer, balTolerance as single
    Dim outRow As Integer, rNum, iCol, pCol, outSheet As String


    '**************************************************
    ' user inputs
    '**************************************************
    ' read in from data entered by user on worksheet
    intRate = Cells(2,2).value  'in decimals
    loanLife = Cells(3,2).value  'in full years
    initLoanAmnt = Cells(4,2).value


    '**************************************************
    ' programmer inputs
    '**************************************************
    outSheet = "Loan Amort"
    balTolerance = 1 'specifies desired accuracy
    iCol = 1
    pCol = 2
    outRow = 8 'row below which repayment schedule would start

    '**************************************************
    'preliminaries
    '**************************************************
    'make the outSheet the active sheet
    worksheets(outSheet).activate

    'clear previous data
    Rows((outRow + 1) & ":" & (outRow + 300)).select
    selection.clear

    redim yrEndBal(loanLife)  'redimension the array


    '**************************************************
    ' compute and output results
    '**************************************************
    annualPmnt = initLoanAmnt * intRate 
    numOfIterations = 0

    'this do loop controls the iteration
    do

        'initialize beginning balance for year 1
        yrBegBal(1) = initLoanAmnt

        'loop to calculate and output year-by-year amort. table
        For rNum = 1 To loanLife
            ipPay(rNum, iCol) = yrBegBal(rNum) * intRate
            ipPay(rNum, pCol) = annualPmnt -  ipPay(rNum, iCol)
            yrEndBal(rNum) = yrBegBal(rNum) - ipPay(rNum, pCol)


            yrBegBal(rNum + 1) = yrEndBal(rNum)

        Next rNum

        finalBal = yrEndBal(loanLife)
        aPmtOld = annualPmnt

        annualPmnt = annualPmnt + (finalBal * (1 + intRate)^ _
        (-loanLife)) / loanLife
        numOfIterations = numOfIterations + 1
    loop while finalBal>= balTolerance


    '**************************************************
    ' output data to worksheet
    '**************************************************
    for rNum = 1 to loanLife
        cells(outRow + rNum, 3).value = rNum 'year number
        cells(outRow + rNum, 4).value = yrBegBal(rNum)
        cells(outRow + rNum, 5).value = annualPmnt
        cells(outRow + rNum, 6).value = ipPay(rNum, iCol)
        cells(outRow + rNum, 7).value = ipPay(rNum, pCol)
        cells(outRow + rNum, 8).value = yrEndBal(rNum)
    next rNum

    'write out the number of iterations used
    cells(outRow + loanLife + 4, 1).value = "no. of iterations ="
    cells(outRow + loanLife + 4, 2).value = numOfIterations

    '**************************************************
    ' format the output data in the table
    '**************************************************
    Range(Cells(outRow + 1, 4), Cells(outRow + loanLife, 8)) _
    .Select
    Selection.NumberFormat = "$#,##0.00"

End Sub


'/* vim: set filetype=vb : */



