Module Module1

    'Program:   The Gift Emporium Sales Receipt
    'Programmer: Ryan Roque
    'Date: 1/30/2017
    'Purpose: User inputs purchase amount. Calculates taxes and total amount. Displays receipt.

    Sub Main()

        'Declarataions
        Const dblSTATE_TAX_RATE As Double = 0.035
        Const dblCOUNTY_TAX_RATE As Double = 0.05
        Const dblCITY_TAX_RATE As Double = 0.0125

        Dim strItem As String = ""
        Dim decPrice As Decimal = 0
        Dim decStateTax As Decimal = 0
        Dim decCountyTax As Decimal = 0
        Dim decCityTax As Decimal = 0
        Dim decTotalPurchase As Decimal = 0

        'Get Input
        getPurchase(strItem, decPrice)

        'Calculate Taxes and Total
        calculateTax(decPrice, dblSTATE_TAX_RATE, decStateTax)
        calculateTax(decPrice, dblCOUNTY_TAX_RATE, decCountyTax)
        calculateTax(decPrice, dblCITY_TAX_RATE, decCountyTax)
        calculateTotal(decPrice, decStateTax, decCountyTax, decCityTax, decTotalPurchase)

        'Display Receipt
        displayTitle()
        displayReceipt(strItem, decPrice, decStateTax, decCountyTax, decCityTax, decTotalPurchase)
        terminateProgram()

    End Sub

    Private Sub getPurchase(ByRef strProduct As String, ByRef decAmount As Decimal)

        Console.Write("What is the product description? ")
        strProduct = Console.ReadLine()
        Console.Write("What is the price of the item? ")
        decAmount = CDec(Console.ReadLine())

    End Sub

    Private Sub calculateTax(ByVal decAmount As Decimal, dblRate As Double, ByRef decTax As Decimal)

        decTax = decAmount * CDec(dblRate)

    End Sub

    Private Sub calculateTotal(ByVal decAmount As Decimal, ByVal decState As Decimal, ByVal decCounty As Decimal,
                               ByVal decCity As Decimal, ByRef decTotal As Decimal)

        decTotal = decAmount + decState + decCity

    End Sub

    Private Sub displayTitle()

        Console.Clear()
        Console.WriteLine()
        Console.WriteLine()

    End Sub

    Private Sub displayReceipt(ByVal strProduct As String, ByVal decAmount As Decimal, ByVal decState As Decimal,
                               ByVal decCounty As Decimal, ByVal decCity As Decimal, ByVal decTotal As Decimal)

        Console.WriteLine("Item: " & strProduct & ControlChars.Tab & decAmount.ToString("C"))
        Console.WriteLine("State Tax:" & ControlChars.Tab & decState)
        Console.WriteLine("County Tax:" & ControlChars.Tab & decCity)
        Console.WriteLine("City Tax:" & ControlChars.Tab & decCity)
        Console.WriteLine("Total Purchase:" & ControlChars.Tab & decTotal.ToString("C"))

    End Sub

    Private Sub terminateProgram()

        Console.WriteLine()
        Console.WriteLine("Press the enter key to terminate the program.")
        Console.Read()

    End Sub

End Module