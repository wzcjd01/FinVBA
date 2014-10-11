sub Loan_Amort_GoalSeek()
    'this version creates an amortization table with formulas and then uses
    'goal seek to find the answer

    dim intRate, initLoanAmnt, loanLife, outRow, outSheet

    '**************************************************
    ' user inputs
    '**************************************************
    ' read in from data entered by user on worksheet
    worksheets("Loan Amort").activate
    intRate = Cells(2,2).value 'in decimals
    loanLife = Cells(3,2).value 'in full years
    initLoanAmnt = Cells(4,2).value 'initial loan balance


    '**************************************************
    ' Programmer inputs
    '**************************************************
    outSheet = "Loan Amort"
    outRow = 8 'row below which repayment schedule would start


    '**************************************************
    ' preliminaries
    '**************************************************
    ' make the outSheet the active sheet
    worksheets(outSheet).activate

    ' clear previous data
    rows((outRow+1) & ":" &(outRow+300)).select
    selection.clear

    Cells(outRow + 1, 3).value = 1
    Cells(outRow + 1, 4).value = initLoanAmnt
    ' Cells(outRow + 1, 5).value = 
    Cells(outRow + 1, 6).formula ="=D9*$B$2"
    Cells(outRow + 1, 7).formula = "=E9-F9"
    Cells(outRow + 1, 8).formula = "=D9-G9"


    Cells(outRow + 2, 3).formula = "=C9+1"
    Cells(outRow + 2, 4).formula = "=H9"
    Cells(outRow + 2, 5).formula = "=E9"
    Cells(outRow + 2, 6).formula ="=D10*$B$2"
    Cells(outRow + 2, 7).formula = "=E10-F10"
    Cells(outRow + 2, 8).formula = "=D10-G10"

    range("C" & (outRow + 2) & ":H" & (outRow + 2)).copy _
    range("C" & (outRow + 3) & ":H" & (outRow + loanLife))


    '**************************************************
    ' use Goal Seek to find answer
    '**************************************************
    range("H" & (outRow + loanLife)).goalseek _
    goal:=0, changingcell:=range("E9")


    '**************************************************
    ' format data in table
    '**************************************************
    range(Cells(outRow + 1, 4), Cells(outRow + loanLife, 8)).select
    selection.numberformat = "$#,##0.00"

end sub
