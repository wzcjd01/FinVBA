option base 1

function MyPmt(intRate, loanLife as Integer, initLoanAmnt)

    'this program creates a user-defined function from a loan amortization
    'program we wrote before. it can be used as a worksheet function to 
    'calculate the periodic payment for an amortizing loan.

    dim yrBegBal(100), yrEndBal(), finalBal
    dim ipPay(1 to 100, 1 to 2)
    dim annaulPmnt, aPmtOld
    dim rNum, balTolerance
    dim iCol, pCol

    '**************************************************
    ' programmer inputs
    '**************************************************
    balTolerance = 1  'specified desired accuracy
    iCol = 1
    pCol = 2

    redim  yrEndBal(loanLife)


    '**************************************************
    ' computations and output
    '**************************************************
    annualPmnt = initLoanAmnt * intRate

    'this do loop controls the iteration
    do
        'initialize balance at the beginning of year 1
        yrBegBal(1) = initLoanAmnt

        'loop to calculate and output year-by-year data
        for rNum = 1 to loanLife
            ipPay(rNum, iCol) = yrBegBal(rNum) * intRate
            ipPay(rNum, pCol) = annualPmnt - ipPay(rNum, iCol)

            yrBegBal(rNum + 1) = yrEndBal(rNum)
        next rNum

        finalBal = yrEndBal(loanLife)
        aPmtOld = annualPmnt

        'calculate the next annual payment to try
        annualPmnt = annualPmnt + (finalBal * (1 + intRate)^ _
            (-loanLife)) / loanLife
        loop while finalBal >= balTolerance

        MyPmt = aPmtOld

    end function
            
