'------------------------------------------------------------
'-                File Name : Program 9                     - 
'-                Part of Project: Class for Sales Record   -
'------------------------------------------------------------
'-                Written By: Frederich Schulz              -
'-                Written On: 4/11/2019                     -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains all of the accessors for the file read-
'- in as well as contains the constructor to read the file  -
'------------------------------------------------------------
Public Class clsDataExtraction
    Private intStoreNumber As Integer
    Private dblJanSales As Double
    Private dblFebSales As Double
    Private dblMarSales As Double
    Private dblAprSales As Double
    Private dblMaySales As Double
    Private dblJunSales As Double
    Private dblJulSales As Double
    Private dblAugSales As Double
    Private dblSepSales As Double
    Private dblOctSales As Double
    Private dblNovSales As Double
    Private dblDecSales As Double
    Private dblTotPrevSales As Double

    'Simple enum created to help w readability of the spliting we will have to do
    'for this program to be successful
    Public Enum SalesData
        StoreNumber
        JanSales
        FebSales
        MarSales
        AprSales
        MaySales
        JunSales
        JulSales
        AugSales
        SepSales
        OctSales
        NovSales
        DecSales
        TotalPrevSales
    End Enum

    '------------------------------------------------------------
    '-          Subprogram Name: New(Line As String)            -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This subs purpose is to read a line in from the file, that-
    '-is comma delimited, then it appends it to the variables we-
    '-made using the enum                                       -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '-Line: takes a line read in from a text file and splits it -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-data: which holds the line to be split                    -
    '------------------------------------------------------------
    Public Sub New(data As String())
        intStoreNumber = data(SalesData.StoreNumber)
        dblJanSales = data(SalesData.JanSales)
        dblFebSales = data(SalesData.FebSales)
        dblMarSales = data(SalesData.MarSales)
        dblAprSales = data(SalesData.AprSales)
        dblMaySales = data(SalesData.AugSales)
        dblJunSales = data(SalesData.JunSales)
        dblJulSales = data(SalesData.JulSales)
        dblAugSales = data(SalesData.AugSales)
        dblSepSales = data(SalesData.SepSales)
        dblOctSales = data(SalesData.OctSales)
        dblNovSales = data(SalesData.NovSales)
        dblDecSales = data(SalesData.DecSales)
        dblTotPrevSales = data(SalesData.TotalPrevSales)
    End Sub

    '------------------------------------------------------------
    '-          Function Name: GetStoreNumber()                 -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the store num, from      -
    '-the data file read in.                                    - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Integer – giving the store number                        -
    '------------------------------------------------------------
    Public Function GetStoreNumber() As Integer
        Return intStoreNumber
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetJanSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for janurary,  -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving janurary sales                           -
    '------------------------------------------------------------
    Public Function GetJanSales() As Double
        Return dblJanSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetFebSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for feburary,  -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving feburary sales                           -
    '------------------------------------------------------------
    Public Function GetFebSales() As Double
        Return dblFebSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetMarSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for March,     -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving march sales                              -
    '------------------------------------------------------------
    Public Function GetMarSales() As Double
        Return dblMarSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetAprSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for April,     -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving April sales                              -
    '------------------------------------------------------------
    Public Function GetAprSales() As Double
        Return dblAprSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetMaySales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for May,       -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving may sales                                -
    '------------------------------------------------------------
    Public Function GetMaySales() As Double
        Return dblMaySales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetJunSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for June,      -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving june sales                               -
    '------------------------------------------------------------
    Public Function GetJunSales() As Double
        Return dblJunSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetJulSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for July,      -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving july sales                               -
    '------------------------------------------------------------
    Public Function GetJulSales() As Double
        Return dblJulSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetAugSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for August,    -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving august sales                             -
    '------------------------------------------------------------
    Public Function GetAugSales() As Double
        Return dblAugSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetSepSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for September, -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving september sales                          -
    '------------------------------------------------------------
    Public Function GetSepSales() As Double
        Return dblSepSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetOctSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for October,   -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving october sales                            -
    '------------------------------------------------------------
    Public Function GetOctSales() As Double
        Return dblOctSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetNovSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for November,  -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving november sales                           -
    '------------------------------------------------------------
    Public Function GetNovSales() As Double
        Return dblNovSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetDecSales()                    -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the sales for December,  -
    '-from the data file read in.                               - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving december sales                           -
    '------------------------------------------------------------
    Public Function GetDecSales() As Double
        Return dblDecSales
    End Function

    '------------------------------------------------------------
    '-          Function Name: GetTotalPrevSales()              -
    '------------------------------------------------------------
    '-                Written By: Frederich Schulz              -
    '-                Written On: 4/11/2019                     -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '-This functions purpose is to get the total sales for,     -
    '-previous year from the data file read in.                 - 
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-N/A                                                       -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double – giving total sales                              -
    '------------------------------------------------------------
    Public Function GetTotalPrevSales() As Double
        Return dblTotPrevSales
    End Function
End Class
