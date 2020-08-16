Module AlcoverR2LoanAccountReport


    '                                                                START OF PROGRAM

    Private LoanAccountReportFile As New Microsoft.
        VisualBasic.FileIO.TextFieldParser("LOANS2019.TXT")

    Private CurentRecord() As String
    '            NOW WE'LL DECLARE THE FILE WE'LL
    '            WE USE IN THE PROGRAM AND ASSICIATE IT 
    '            WITH THE ACTUAL FILE NAME, WHERE THE DATA IS STORED


    '                                                                   INPUT VARIABLES/FLIEDS:
    Private CustomerNumberDecimal As Decimal
    Private CustomerNameString As String
    Private MonthlyBeginningBalanceDecimal As Decimal
    Private MonthlyChargesDecimal As Decimal
    Private MonthlyPaymentDecimal As Decimal

    '                                                                   CALCULATED FIELDS:
    Private OutstandingAmountDecimal As Decimal  '                      EACH SALESPERSON /RECORD
    Private InterestAddedDecimal As Decimal
    Private EndingBalanceDecimal As Decimal
    '                                                               CALCULATED FOR END OF REPORT
    Private AccumBeginningBalanceDecimal As Decimal = 0
    Private AccumMonthlyPaymentDecimal As Decimal = 0
    Private AccumInterestAddedPaymentDecimal As Decimal = 0
    Private AccumEndingBalancePaymentDecimal As Decimal = 0
    Private AccumCustomerNumberDecimal As Decimal = 0


    Private AverageBeginningBalanceDecimal As Decimal
    Private AverageMonthlyPaymentDecimal As Decimal
    Private AverageInterestAddedPaymentDecimal As Decimal
    Private AverageEndingBalancePaymentDecimal As Decimal
    '                                                            ACCUMULATED FINALS TOTALS
    '                                                            FINAL  TOTALS FOR END OF REPORT

    Private AccumCustomerNumberInteger As Integer = 0    '       COUNTER NUMBER OF EMPLOYEES
    '                                                            CONSTANT FIELD:
    Private Const INTEREST_RATE As Decimal = 0.3775     '        INTEREST RATE

    '                                        PAGINATION VARIABLES:

    Private LineCounterInteger As Integer = 99         '              99 FOR HEADINGS ON FIRST PAGE
    '                                                              
    Private Const PAGE_SIZE_INTEGER As Integer = 15
    '                                                              
    Private PageNumberInteger As Integer = 1 '                      Page #'s for headings            
    '                                                      FILE RECORD AND FILE NAME DECLARATIONS:
    '                                                      WHEN THE FILE IS READ, 
    '                                                      THE RECORD IS PLACED IN THIS VARIABLE

    Sub Main()   '                                         PROGRAM EXECUTION LOGIC STARTS.
        Call HouseKeeping()
        Do While Not LoanAccountReportFile.EndOfData
            Call ProcessRecords()
        Loop
        Call EndOfJob()
    End Sub

    Private Sub HouseKeeping()  '                          LEVEL 2 CONTROL MODULES
        Call SetFileDelimiter()

    End Sub

    Private Sub ProcessRecords()
        Call ReadFile()
        Call DetailCalculation()
        Call AccumulateTotals()
        Call WriteDetailLine()
    End Sub

    Private Sub EndOfJob()
        Call SummaryCalculations()
        Call SummaryOutput()
        Call CloseFile()

    End Sub


    Private Sub SetFileDelimiter()
        LoanAccountReportFile.TextFieldType = FileIO.FieldType.Delimited

        LoanAccountReportFile.SetDelimiters(",")
        '                                HOUSEKEAPING MODULES
        '                                DEFINES FILES AS A DELIMITER
        '                                DEFINES DELIMITER AS A COMMA                                                    
    End Sub
    Private Sub ReadFile()
        '         READ WHOLE RECORD AND ASSIGN TO THE CURRENT RECORD VARIABLE
        CurentRecord = LoanAccountReportFile.ReadFields()
        '        PLACE CURRENT RECORDS FIELDS INTO THEIR RESPECTIVE VARIABLES
        '        THE CURRENT RECORD 1 IS SKIP BECAUSE IS NOT USED 
        CustomerNumberDecimal = CurentRecord(0)
        MonthlyBeginningBalanceDecimal = CurentRecord(2)
        MonthlyChargesDecimal = CurentRecord(3)
        MonthlyPaymentDecimal = CurentRecord(4)
    End Sub

    Private Sub DetailCalculation() '                   CALCULATIONS
        OutstandingAmountDecimal = MonthlyBeginningBalanceDecimal + MonthlyChargesDecimal - MonthlyPaymentDecimal
        InterestAddedDecimal = OutstandingAmountDecimal * INTEREST_RATE
        InterestAddedDecimal = (Math.Round(InterestAddedDecimal, 2))
        ' ROUND INTERESTADDED BEFORE CALCULATE ENDING BALANCE
        EndingBalanceDecimal = OutstandingAmountDecimal + InterestAddedDecimal

    End Sub

    Private Sub AccumulateTotals()
        '                 ACCUMULATE FINAL TOTALS LOAN ACCOUNT REPORT
        AccumBeginningBalanceDecimal = AccumBeginningBalanceDecimal + MonthlyBeginningBalanceDecimal
        AccumMonthlyPaymentDecimal = AccumMonthlyPaymentDecimal + MonthlyPaymentDecimal
        AccumInterestAddedPaymentDecimal = AccumInterestAddedPaymentDecimal + InterestAddedDecimal
        AccumEndingBalancePaymentDecimal = AccumEndingBalancePaymentDecimal + EndingBalanceDecimal

        AccumCustomerNumberDecimal = AccumCustomerNumberDecimal + 1
        '                                       COUNT OF # EMPLOYEES
    End Sub



    Private Sub WriteDetailLine()

        If LineCounterInteger >= PAGE_SIZE_INTEGER Then
            Call WriteHeadings()
        End If


        '                         WRITE DETAIL LINE

        Console.WriteLine(Space(1) & CustomerNumberDecimal.ToString("n0").PadLeft(3) &
                          Space(2) & MonthlyBeginningBalanceDecimal.ToString("n").PadLeft(14) &
                          Space(2) & MonthlyChargesDecimal.ToString("n").PadLeft(8) &
                          Space(3) & MonthlyPaymentDecimal.ToString("n").PadLeft(9) &
                          Space(4) & OutstandingAmountDecimal.ToString("n").PadLeft(9) &
                          Space(3) & InterestAddedDecimal.ToString("n").PadLeft(8) &
                          Space(4) & EndingBalanceDecimal.ToString("n").PadLeft(9))
        '                                     LineCounterInteger = LineCounterInteger +1    
        '                                     COUNT THE LINE PRINTED

        LineCounterInteger += 1 ' +=  IS A ' COMBINED OPERATOR', SHORTCUT FOR ACCUMULATION
        '                                     OUTPUT 1 LINE FOR EACH PERSON PROCESSED 
        '                                     TEST FOR PAGINATION


    End Sub

    Private Sub WriteHeadings() ' WRITE HEADINGS MODULE IS PART OF PROCESS RECORD MODULES
        Console.WriteLine() '     AND IS CALL BY WRITE DETAILLINE WEN THE LINE COUNTER 
        Console.WriteLine() '      IS GREATER OR EQUAL TO 15.
        Console.WriteLine()                                 'WRITE REPORTHEADLINES
        Console.WriteLine(Space(24) & "Shark Attack, Inc. Loan Report" &
                          Space(16) & "Page " & PageNumberInteger.ToString("n0".PadLeft(2)))
        Console.WriteLine(Space(20) & "Report Created by Rodrigo Martin Alcover")
        Console.WriteLine()                                 'WRITE COLUM LEADER LINES
        Console.WriteLine(Space(1) & "Cust" & Space(6) & "Beginning" &
                          Space(7) & "New" & Space(5) & "Monthly" &
                          Space(2) & "Outstanding" & Space(3) & "Interest" &
                          Space(7) & "Ending")
        Console.WriteLine(Space(1) & "Num" & Space(9) & "Balance" &
                          Space(3) & "Charges" & Space(5) & "Payment" &
                          Space(7) & "Amount" & Space(6) &
                          "Added" & Space(6) & "Balance")

        Console.WriteLine()
        LineCounterInteger = 0 '               RESET LINE COUNTER &
        PageNumberInteger += 1       '   ADD TO PAGE#     +=  IS CALLED A  COBINED OPERATOR

    End Sub
    '                                          END OF JOBS MODULES
    '                                          FINAL AVERAGES
    Private Sub SummaryCalculations() 'STATE CALCULATIONS

        AverageBeginningBalanceDecimal = MonthlyBeginningBalanceDecimal / AccumCustomerNumberDecimal
        AverageMonthlyPaymentDecimal = AccumMonthlyPaymentDecimal / AccumCustomerNumberDecimal
        AverageInterestAddedPaymentDecimal = AccumInterestAddedPaymentDecimal / AccumCustomerNumberDecimal
        AverageEndingBalancePaymentDecimal = AccumEndingBalancePaymentDecimal / AccumCustomerNumberDecimal

    End Sub

    Private Sub SummaryOutput()
        Console.WriteLine() ' WRITE TOTAL LINE AND MOVE ACCUM
        Console.WriteLine()
        Console.WriteLine(Space(0) & "Totals:  " & AccumBeginningBalanceDecimal.ToString("c").PadLeft(11) &
                          Space(11) & AccumMonthlyPaymentDecimal.ToString("c").PadLeft(11) & Space(13) &
                          AccumInterestAddedPaymentDecimal.ToString("c").PadLeft(11) & Space(2) &
                          AccumEndingBalancePaymentDecimal.ToString("c").PadLeft(11))
        Console.WriteLine() ' WRITE AVERAGE LINE AND MOVE AVERAGES
        Console.WriteLine("Averages: " & AverageBeginningBalanceDecimal.ToString("c").PadLeft(10) &
                          Space(13) & AverageMonthlyPaymentDecimal.ToString("c").PadLeft(9) &
                          Space(14) & AverageInterestAddedPaymentDecimal.ToString("c").PadLeft(10) &
                          Space(3) & AverageEndingBalancePaymentDecimal.ToString("c").PadLeft(10))
        Console.WriteLine() ' WRITE NUMBER OF CUSTOMERS PROCESSED AND MOVE ACCUMCUSTOMER
        Console.WriteLine("Number of customers Processed:  " & AccumCustomerNumberDecimal.ToString("n0"))

    End Sub

    Private Sub CloseFile()                        ' END OF JOB MODULES
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine("Click ENTER Close Output Window")
        Console.ReadKey() ' WRITE MESSAGE FOR PRESS ENTER AND CLOSE THE WINDOW
        '                   PROMPT

    End Sub

End Module


