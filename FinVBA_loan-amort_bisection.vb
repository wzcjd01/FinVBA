option base 1
dim yrBegBal(100), yrEndBal(100)
dim ipPay(1 to 100, 1 to 3), tCol, iCol, pcol

sub Loan_Amort_Bisection()
    'this version uses the bisection method for the iteration. we start with
    'a good initial guess and a fair large initial step and the bisection
    'method takes over, halving the step size with each iteration.

    dim intRate, loanSize, loanLife
    dim annaulPmnt, finalBal
    dim outRow1, outRow2, rNum, numOfIt, errMar
    dim outSheet1, outSheet2, aPmnt, aStep, fBalL, fBalH
    dim itStage, itTrack(100,6)  'to track iteration results

    '**************************************************
    'programmer inputs
    '**************************************************
    outSheet1 = "Loan Amort"
    outSheet2 = "Iterations"
    tCol = 1 'column number for total payment
    iCol = 2 'column number for interest payment
    pCol = 2 'column number for pincipal component
    outRow1 = 8 'row below which repayment schedule would start
    outRow2 = 5
    numOfIt = 0 'couter for number of iterations
    itStage = 1 'used to distinguish stage 1 from stage 2

    errMar = 1 'specified desired accuracy
    aStep = 1000 'initial step size

    '**************************************************
    'preliminaries
    '**************************************************
    worksheets(outSheet2).activate

    'clear previous data from iterations sheet
    rows((outRow2 + 1) & ":" & (outRow2 + 300)).select
    selection.clear

    worksheets(outSheet1).activate

    'clear previous data from Loan Amort sheet
    rows((outRow1 + 1) & ":" & (outRow1 + 300)).select
    selection.clear


    '**************************************************
    'user inputs
    '**************************************************
    'read in from data entered by user on worksheet
    intRate = cells(2,2).value 'in decimals
    loanLife = cells(3,2).value  'in full years
    loansize = cells(4,2).value  'initial loan balance

    '**************************************************
    'computations
    '**************************************************
    aPmnt = loanSize * intRate  'initial guess for annual payment

    call Calc_Table(loanSize, aPmnt, intRate, loanLife, finalBal)
    fBalL = finalBal  'final balance for lower end of interval

    do
        numOfIt = numOfIt + 1
        call Calc_Table(loanSize, aPmnt + aStep, intRate, loanLife, finalBal)
        fBalH = finalBal 'final bal. for higher end of interval

        itTrack(numOfIt, 1) = numOfIt 
        itTrack(numOfIt, 2) = aPmnt 
        itTrack(numOfIt, 3) = aStep 
        itTrack(numOfIt, 4) = aPmnt + aStep 
        itTrack(numOfIt, 5) = fBalL 
        itTrack(numOfIt, 6) = fBalH 

        if (fBalL * fBalH) > 0 then
            aPmnt = aPmnt + aStep
            fBalL = fBalH
            if itStage <> 1 then aStep = aStep / 2
        else
            itStage = 2
            aStep = aStep / 2
        end if 

    loop while (abs(fBalH) - errMar) > 0

    '**************************************************
    ' output data to worksheet
    '**************************************************
    for rNum = 1 to loanLife
        cells(outRow1 + rNum, 3).value = rNum 'year number
        cells(outRow1 + rNum, 4).value = yrBegBal(rNum)
        cells(outRow1 + rNum, 5).value = ipPay(rNum, tCol)
        cells(outRow1 + rNum, 6).value = ipPay(rNum, iCol)
        cells(outRow1 + rNum, 7).value = ipPay(rNum, pCol)
        cells(outRow1 + rNum, 8).value = yrEndBal(rNum)
    next rNum

    range(cells(outRow1 + 1, 4), cells(outRow1 + loanLife, 8)).select
    selection.numberformat = "$#,##0"

    sheets(outSheet2).activate

    for rNum = 1 to numOfIt
        cells(outRow2 + rNum, 1).value = itTrack(rNum, 1)
        cells(outRow2 + rNum, 2).value = itTrack(rNum, 2)
        cells(outRow2 + rNum, 3).value = itTrack(rNum, 3)
        cells(outRow2 + rNum, 4).value = itTrack(rNum, 4)
        cells(outRow2 + rNum, 5).value = itTrack(rNum, 5)
        cells(outRow2 + rNum, 6).value = itTrack(rNum, 6)
    next rNum

    range(cells(outRow2 + 1, 1), cells(outRow2 + numOfIt, 1)).select
    selection.numberformat = "0"

    range(cells(outRow2 + 1, 2), cells(outRow2 + numOfIt, 6)).select
    selection.numberFormat = "$#,##0.00_);($#,##0.00)"

end sub


'--------------------------------------------------
'--------------------------------------------------
sub Calc_Table(loanSize,annualPmnt,intRate,loanLife,fBal)
    ' this program generates an amortization table and the 
    ' remaining balance at the end of the final period.

    dim rNum

    yrBegBal(1) = loanSize
    for rNum = 1 to loanLife
        ipPay(rNum, tCol) = annualPmnt
        ipPay(rNum, iCol) = yrBegBal(rNum) * intRate
        ipPay(rNum, pCol) = annualPmnt - ipPay(rNum, iCol)
        yrEndBal(rNum) = yrBegBal(rNum) - ipPay(rNum, pCol)
        yrBegBal(rNum + 1) = yrEndBal(rNum)
    next rNum

    fBal = yrEndBal(loanLife) 'final balance

end sub
