Attribute VB_Name = "User-Functions"
Option Compare Database
Option Explicit


Public Function BenOption(DedCode, Amount) As String
    Select Case Trim(DedCode)
        Case Is = "LIFED"
            If Amount = 0.96 Then
                BenOption = "DEPL2"
            ElseIf Amount = 0.48 Then
                BenOption = "DEPL1"
            ElseIf Amount = 0 Then
                BenOption = "DEPL1"
             Exit Function
            End If
            
        Case Is = "LEGAL"
            If Amount = 5.98 Then
                BenOption = "LEGIDO"
            ElseIf Amount = 11.95 Then
                BenOption = "LEGIDT"
            ElseIf Amount = 7.36 Then
                BenOption = "LEGONL"
             Exit Function
            End If
        Case Is = "2"
            BenOption = "NSMOKE"
        Case Is = "3"
            BenOption = "HASSES"
        Case Is = "4", "14", "23"
            BenOption = "EE"
        Case Is = "16", "19", "24"
            BenOption = "EEC"
        Case Is = "20", "25", "28"
            BenOption = "EES"
        Case Is = "1", "26", "27"
            BenOption = "EEF"
        Case Else
            BenOption = ""
    End Select
    
        
End Function



Public Function NumUnPuncher(iNumber)
Dim x As String
Dim N As String
    
    x = Right(Trim(iNumber), 1)
    
    Select Case Right(Trim(iNumber), 1)
        Case Is = "}"
            N = "0"
        Case Is = "J"
            N = "1"
        Case Is = "K"
            N = "2"
        Case Is = "L"
            N = "3"
        Case Is = "M"
            N = "4"
        Case Is = "N"
            N = "5"
        Case Is = "O"
            N = "6"
        Case Is = "P"
            N = "7"
        Case Is = "Q"
            N = "8"
        Case Is = "R"
            N = "9"
        Case Else
            x = "FAIL"
    End Select
                                   
NumUnPuncher = Replace(iNumber, x, N)

End Function
Public Function GTDAmt(DedAmt As Double, Goal As Double, GoalToDate As Double) As Double
    If Goal <> 0 Then
        If Goal = GoalToDate Then
            GTDAmt = GoalToDate
        ElseIf GoalToDate + DedAmt > Goal Then
            GTDAmt = Goal
        Else
            GTDAmt = GoalToDate + DedAmt
        End If
    End If
End Function

Public Function PadL(padString As String, padChar As String, Optional padSize As Integer, Optional padHowMany As Integer) As String
    ' padSize is the size of the return string
    ' padHowMany is how many of padChar to use
    ' cannot have both, padSize takes precendence
    Dim t
    padChar = Left(padChar, 1) ' it's a char, after all
    padString = Trim(padString)
    If padSize = 0 And padHowMany = 0 Then
        ' they didn't say, so it's just a TRIM()
        PadL = padString
        Exit Function
    End If
    If padSize = 0 Then
        ' use padHowMany
        PadL = String(padHowMany, padChar) & padString
        Exit Function
    End If
    ' treat it as a padSize, even though thay may have
    ' erroneously supplied a padHowmany
    For t = 1 To (padSize - Len(padString))
        padString = padChar & padString
    Next t
    PadL = padString

End Function
Public Function PadR(padString As String, padChar As String, Optional padSize As Integer, Optional padHowMany As Integer) As String
    ' padSize is the size of the return string
    ' padHowMany is how many of padChar to use
    ' cannot have both, padSize takes precendence
    Dim t
    padChar = Left(padChar, 1) ' it's a char, after all
    padString = Trim(padString)
    If padSize = 0 And padHowMany = 0 Then
        ' they didn't say, so it's just a TRIM()
        PadR = padString
        Exit Function
    End If
    If padSize = 0 Then
        ' use padHowMany
        PadR = padString & String(padHowMany, padChar)
        Exit Function
    End If
    ' treat it as a padSize, even though thay may have
    ' erroneously supplied a padHowmany
    For t = 1 To (padSize - Len(padString))
        padString = padString & padChar
    Next t
    PadR = padString
    Exit Function
End Function
Public Function Strip(ByVal stripString As String, ByVal stripChar As String) As String
    Dim A, x, y, z As Integer
    Dim StripStringValue, StripCharValue, FixedString As String
    A = 1
    For x = 1 To Len(stripChar)
        StripCharValue = Mid(stripChar, A, 1)
        If InStr(1, stripString, StripCharValue, vbTextCompare) Then
            z = 1
            For y = 1 To Len(stripString)
                StripStringValue = Mid(stripString, z, 1)
                If StripStringValue = StripCharValue Then
                    'do nothing
                Else
                    FixedString = FixedString & StripStringValue
                End If
                z = z + 1
            Next y
            If Len(stripChar) > 1 Then
                stripString = FixedString
                FixedString = ""
                Strip = stripString
            Else
                Strip = FixedString
            End If
        Else
            Strip = stripString
        End If
        A = A + 1
    Next x
End Function
Public Function City(ByVal iCityState As String) As String
    Dim A, x, y, z As Integer
    Dim StripStringValue, StripCharValue, AddressCity, FindComma As String
    StripCharValue = ""
    x = Len(iCityState)
    FindComma = ","
    A = 1
    While StripCharValue <> FindComma
        StripCharValue = Mid(iCityState, A, 1)
        If InStr(1, StripCharValue, FindComma, vbTextCompare) Then
            City = Proper(AddressCity)
            Exit Function
        Else
            AddressCity = AddressCity & StripCharValue
        End If
        A = A + 1
        If A > x Then
            City = Proper(Left(AddressCity, A - 3))
            Exit Function ' did not find comma
        End If
    Wend
End Function
Public Function State(ByVal iCityState As String) As String
    State = UCase(Right(Trim(iCityState), 2))
    
End Function
Public Function StateCode(ByVal iStateWorkIn As String) As String
    Select Case iStateWorkIn
        Case Is = "01"
            StateCode = "NYSIT"
        Case Is = "02"
            StateCode = "MASIT"
        Case Is = "03"
            StateCode = "ALSIT"
        Case Is = "04"
            StateCode = "ERR04"
        Case Is = "05"
            StateCode = "MDSIT"
        Case Is = "06"
            StateCode = "ERR06"
        Case Is = "07"
            StateCode = "ERR07"
        Case Is = "08"
            StateCode = "ERR08"
        Case Is = "09"
            StateCode = "ERR09"
        Case Is = "10"
            StateCode = "ERR10"
        Case Is = "11"
            StateCode = "ERR11"
        Case Is = "12"
            StateCode = "NCSIT"
        Case Is = "13"
            StateCode = "ERR13"
        Case Is = "14"
            StateCode = "WISIT"
        Case Is = "15"
            StateCode = "COSIT"
        Case Is = "16"
            StateCode = "ERR16"
        Case Is = "17"
            StateCode = "VASIT"
        Case Is = "18"
            StateCode = "ARSIT"
        Case Is = "19"
            StateCode = "ERR19"
        Case Is = "20"
            StateCode = "MNSIT"
        Case Is = "21"
            StateCode = "INSIT"
        Case Is = "22"
            StateCode = "MOSIT"
        Case Is = "23"
            StateCode = "GASIT"
        Case Is = "24"
            StateCode = "ERR24"
        Case Is = "25"
            StateCode = "CASIT"
        Case Is = "26"
            StateCode = "WVSIT"
        Case Is = "27"
            StateCode = "CTSIT"
        Case Is = "28"
            StateCode = "ERR28"
        Case Is = "29"
            StateCode = "ERR29"
        Case Is = "30"
            StateCode = "OHSIT"
        Case Is = "31"
            StateCode = "KYSIT"
        Case Is = "32"
            StateCode = "ERR32"
        Case Is = "33"
            StateCode = "ERR33"
        Case Is = "34"
            StateCode = "PRSIT"
        Case Is = "35"
            StateCode = "SCSIT"
        Case Is = "36"
            StateCode = "OKSIT"
        Case Is = "37"
            StateCode = "AZSIT"
        Case Is = "38"
            StateCode = "ERR38"
        Case Is = "39"
            StateCode = "ORSIT"
        Case Is = "40"
            StateCode = "IDSIT"
        Case Is = "41"
            StateCode = "ERR41"
        Case Is = "42"
            StateCode = "FLSIT"
        Case Is = "43"
            StateCode = "ILSIT"
        Case Is = "44"
            StateCode = "KSSIT"
        Case Is = "45"
            StateCode = "MESIT"
        Case Is = "46"
            StateCode = "MSSIT"
        Case Is = "47"
            StateCode = "ERR47"
        Case Is = "48"
            StateCode = "ERR48"
        Case Is = "49"
            StateCode = "ERR49"
        Case Is = "50"
            StateCode = "ERR50"
        Case Is = "51"
            StateCode = "SDSIT"
        Case Is = "52"
            StateCode = "TNSIT"
        Case Is = "53"
            StateCode = "TXSIT"
        Case Is = "54"
            StateCode = "WASIT"
        Case Is = "55"
            StateCode = "ERR55"
        Case Is = "56"
            StateCode = "ERR56"
        Case Is = "57"
            StateCode = "ERR57"
        Case Is = "58"
            StateCode = "NHSIT"
        Case Is = "59"
            StateCode = "PASIT"
        Case Is = "60"
            StateCode = "MISIT"
        
        Case Else
            StateCode = "ERROR"
    End Select
End Function
Public Function SUICode(ByVal iSUIDisCode As String) As String
    Select Case iSUIDisCode
        Case Is = "01"
            SUICode = "ERR01"
        Case Is = "02"
            SUICode = "MASUIER"
        Case Is = "03"
            SUICode = "ALSUIER"
        Case Is = "04"
            SUICode = "ERR04"
        Case Is = "05"
            SUICode = "MDSUIER"
        Case Is = "06"
            SUICode = "ERR06"
        Case Is = "07"
            SUICode = "ERR07"
        Case Is = "08"
            SUICode = "ERR08"
        Case Is = "09"
            SUICode = "ERR09"
        Case Is = "10"
            SUICode = "ERR10"
        Case Is = "11"
            SUICode = "ERR11"
        Case Is = "12"
            SUICode = "NCSUIER"
        Case Is = "13"
            SUICode = "ERR13"
        Case Is = "14"
            SUICode = "ERR14"
        Case Is = "15"
            SUICode = "COSUIER"
        Case Is = "16"
            SUICode = "ERR16"
        Case Is = "17"
            SUICode = "VASUIER"
        Case Is = "18"
            SUICode = "ARSUIER"
        Case Is = "19"
            SUICode = "NYSUIER"
        Case Is = "20"
            SUICode = "MNSUIER"
        Case Is = "21"
            SUICode = "ERR21"
        Case Is = "22"
            SUICode = "ERR22"
        Case Is = "23"
            SUICode = "GASUIER"
        Case Is = "24"
            SUICode = "ERR24"
        Case Is = "25"
            SUICode = "ERR25"
        Case Is = "26"
            SUICode = "WVSUIER"
        Case Is = "27"
            SUICode = "CTSUIER"
        Case Is = "28"
            SUICode = "UTSUIER"
        Case Is = "29"
            SUICode = "PRSUIER"
        Case Is = "30"
            SUICode = "OHSUIER"
        Case Is = "31"
            SUICode = "KYSUIER"
        Case Is = "32"
            SUICode = "ERR32"
        Case Is = "33"
            SUICode = "ERR33"
        Case Is = "34"
            SUICode = "ERR34"
        Case Is = "35"
            SUICode = "SCSUIER"
        Case Is = "36"
            SUICode = "ERR36"
        Case Is = "37"
            SUICode = "AZSUIER"
        Case Is = "38"
            SUICode = "ERR38"
        Case Is = "39"
            SUICode = "ERR39"
        Case Is = "40"
            SUICode = "ERR40"
        Case Is = "41"
            SUICode = "ERR41"
        Case Is = "42"
            SUICode = "FLSUIER"
        Case Is = "43"
            SUICode = "ILSUIER"
        Case Is = "44"
            SUICode = "KSSUIER"
        Case Is = "45"
            SUICode = "MESUIER"
        Case Is = "46"
            SUICode = "MSSUIER"
        Case Is = "47"
            SUICode = "ERR47"
        Case Is = "48"
            SUICode = "ERR48"
        Case Is = "49"
            SUICode = "ERR49"
        Case Is = "50"
            SUICode = "ERR50"
        Case Is = "51"
            SUICode = "SDSUIER"
        Case Is = "52"
            SUICode = "TNSUIER"
        Case Is = "53"
            SUICode = "TXSUIER"
        Case Is = "54"
            SUICode = "WASUIER"
        Case Is = "55"
            SUICode = "ERR55"
        Case Is = "56"
            SUICode = "ERR56"
        Case Is = "57"
            SUICode = "ERR57"
        Case Is = "58"
            SUICode = "ERR58"
        Case Is = "59"
            SUICode = "PASUIER"
        Case Is = "60"
            SUICode = "MISUIER"
        Case Is = "61"
            SUICode = "ERR61"
        Case Is = "62"
            SUICode = "ERR62"
        Case Is = "63"
            SUICode = "ERR63"
        Case Is = "64"
            SUICode = "ERR64"
        Case Is = "65"
            SUICode = "ERR65"
        Case Is = "66"
            SUICode = "ERR66"
        Case Is = "67"
            SUICode = "ERR67"
        Case Is = "68"
            SUICode = "ERR68"
        Case Is = "69"
            SUICode = "ERR69"
        Case Is = "70"
            SUICode = "ERR70"
        Case Is = "71"
            SUICode = "ERR71"
        Case Is = "72"
            SUICode = "ERR72"
        Case Is = "73"
            SUICode = "ERR73"
        Case Is = "74"
            SUICode = "ERR74"
        Case Is = "75"
            SUICode = "CASUIER"
        Case Is = "76"
            SUICode = "ERR76"
        Case Is = "77"
            SUICode = "ERR77"
        Case Is = "78"
            SUICode = "ERR78"
        Case Is = "79"
            SUICode = "ERR79"
        Case Is = "80"
            SUICode = "ERR80"
        Case Is = "81"
            SUICode = "ERR81"
        Case Is = "82"
            SUICode = "ERR82"
        Case Is = "83"
            SUICode = "ERR83"
        Case Is = "84"
            SUICode = "ERR84"
        Case Is = "85"
            SUICode = "ERR85"
        Case Is = "86"
            SUICode = "ERR86"
        Case Is = "87"
            SUICode = "WISUIER"
        Case Is = "88"
            SUICode = "ERR88"
        Case Is = "89"
            SUICode = "INSUIER"
        Case Is = "90"
            SUICode = "MOSUIER"
        Case Is = "91"
            SUICode = "ERR91"
        Case Is = "92"
            SUICode = "ERR92"
        Case Is = "93"
            SUICode = "ORSUIER"
        Case Is = "94"
            SUICode = "IDSUIER"
        Case Is = "95"
            SUICode = "ERR95"
        Case Else
            SUICode = "ERROR"
    End Select
End Function
Public Function LITCounty(ByVal iLocalTax As String, ByVal iStateTax As String, ByVal iCallingField As String) As String
    Select Case iLocalTax
    
        Case Is = "110H", "110F", "215H", "5071"
            If iStateTax = "30" And iCallingField = "Resident" Then
                LITCounty = ""
            Else
                LITCounty = "MERCER"
            End If
        Case Is = "110P"
            LITCounty = "ALLEGHENY"
        Case Is = Null
            LITCounty = ""
        Case Else
            LITCounty = ""
    End Select
End Function

Public Function LITWkIn(ByVal iLocalTax As String, ByVal iStateTax As String) As String
    Select Case iLocalTax
        Case Is = "110H"
            Select Case iStateTax
                Case Is = "30"
                    LITWkIn = "PA110475"
                Case Else
                    LITWkIn = "PA100214"
            End Select
        Case Is = "215H"
            LITWkIn = "PA103050"
        Case Is = "5071"
            LITWkIn = "PA111397"
        Case Is = "110F"
            LITWkIn = "PA113551"
        Case Is = "110P"
            LITWkIn = "PA100495"
        Case Is = Null
            LITWkIn = ""
        Case Else
            LITWkIn = ""
    End Select
End Function
Public Function Location(ByVal iHomeDept As String, Optional ByVal iLocation As String) As String
    If Mid(iHomeDept, 3, 1) = "L" And Mid(iHomeDept, 4, 1) = "G" And Trim(iLocation) = "GBO" Then
        Location = "GBL"
        Exit Function
    ElseIf Mid(iHomeDept, 3, 1) = "L" And Mid(iHomeDept, 4, 1) = "C" And Trim(iLocation) = "CHLT" Then
        Location = "LOC"
        Exit Function
    End If
    If Mid(iHomeDept, 4, 1) = "P" And Trim(iLocation) = "STV" Then
        Location = "SVP"
        Exit Function
    ElseIf Mid(iHomeDept, 4, 1) = "E" And Trim(iLocation) = "GBO" Then
        Location = "GBE"
        Exit Function
    End If
    Dim db As Database
    Dim vloc, SQL As String
    Dim rst As Recordset
    
    Set db = CurrentDb
    
    SQL = "SELECT UltiLoc from LocListing where [Old Name] = '" & iLocation & "'"
    Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
    If rst.RecordCount = 0 Then
        Location = "ERROR"
        GoTo CleanClose
    Else
        Location = rst![UltiLoc]
        
    End If
CleanClose:
    rst.Close
    db.Close
    Set rst = Nothing
    Set db = Nothing
    
End Function
Public Function LITSD(ByVal iLocalTax As String, ByVal iStateTax As String) As String
    Select Case iLocalTax
        Case Is = "110H"
            If iStateTax = "30" Then
                LITSD = ""
            Else
                LITSD = "PA119371"
            End If
        Case Is = "215H"
            LITSD = "PA103092"
        Case Is = "5071"
            LITSD = "PA110476"
        Case Is = "110F"
            LITSD = "PA111475"
        Case Is = "110P"
            LITSD = "PA110843"
        Case Is = Null
            LITSD = ""
        Case Else
            LITSD = ""
    End Select
End Function
Public Function MakeDate(ByVal iDate As String) As Date
    If Trim(iDate) = "" Then
        'Do nothing  Invalid date
    ElseIf iDate = "0000000000" Then
        'Do nothing  Invalid date
        
    Else
        MakeDate = CDate(Mid(iDate, 3, 2) & "/" & Mid(iDate, 5, 2) & "/" & Mid(iDate, 7, 4))
        MakeDate = FixDate(MakeDate)
    End If
    
End Function
Public Function LastName(ByVal iName As String) As String
    Dim A, x, y, z As Integer
    Dim StripStringValue, StripCharValue, Lname, FindComma As String
    StripCharValue = ""
    FindComma = ","
    A = 1
   While StripCharValue <> FindComma
        StripCharValue = Mid(iName, A, 1)
        If InStr(1, FindComma, StripCharValue, vbTextCompare) Then
            Lname = RemoveSuffix(Strip(Lname, " "))
            LastName = Proper(Lname)
            Exit Function
        Else
            Lname = Lname & StripCharValue
        End If
        A = A + 1
    Wend
    LastName = Proper(Lname)
End Function
Public Function FirstName(ByVal iName As String) As String
    Dim A, x, y, z As Integer
    Dim StripStringValue, StripCharValue, Fname, FindSpace As String
    Dim Getname As Boolean
    StripCharValue = ""
    FindSpace = " "
    A = 1
    iName = RemoveSuffix(iName)
    Getname = False
KeepLooking:
    While StripCharValue <> FindSpace
        StripCharValue = Mid(iName, A, 1)
        If StripCharValue = "," Then
                Getname = True
                GoTo reread
        End If
        If InStr(1, FindSpace, StripCharValue, vbTextCompare) And Getname Then
            FirstName = Proper(Fname)
            Exit Function
        Else
            If Getname Then
                Fname = Fname & StripCharValue
            Else
                'do nothing
            End If
        End If
reread:
    A = A + 1
    Wend
     If Getname = False Then
        A = A + 1
        StripCharValue = "Z"
        GoTo KeepLooking
    End If
    FirstName = Proper(Fname)
End Function
Public Function MiddleName(ByVal iName As String) As String
    Dim A, x, y, z As Integer
    Dim StripStringValue, StripCharValue, Mname, FindSpace As String
    Dim Getname, FoundComma As Boolean
    StripCharValue = ""
    FindSpace = " "
    A = 1
    x = 1
    iName = RemoveSuffix(iName)
    Getname = False
    FoundComma = False
KeepLooking:
    While x <= Len(iName)
        StripCharValue = Mid(iName, A, 1)
        If StripCharValue = "," Then
            FoundComma = True
        End If
        If StripCharValue = " " And FoundComma Then
                Getname = True
                GoTo reread
        End If
        If InStr(1, FindSpace, StripCharValue, vbTextCompare) And Getname Then
            MiddleName = Proper(Mname)
            Exit Function
        Else
            If Getname Then
                Mname = Mname & StripCharValue
            Else
                'do nothing
            End If
        End If
reread:
        A = A + 1
        x = x + 1
    Wend
    'If Getname = False Then
    '    A = A + 1
    '    x = x + 1
    '    StripCharValue = "Z"
    '    GoTo KeepLooking
    'End If
    MiddleName = Proper(Mname)
End Function
Function COUNTRY(COUNTRYORG) As String
    If COUNTRYORG = "CU" Then
        COUNTRY = "CUB"
        ElseIf COUNTRYORG = "NON" Then
        COUNTRY = "Z"
        Else
        COUNTRY = "USA"
    End If
End Function
Public Function SalaryNonExempt(FileNo) As String
    Select Case FileNo
        Case Is = "040577", "040677", "406797", "003048", "003022", "040531", "406799", "040695", "406835"
            SalaryNonExempt = "True"
    Case Else
        SalaryNonExempt = "False"
    End Select
End Function
Function YesOrNo(Incoming) As String
    'Takes a Yes or No field equivalent and converts it to either yes or no
    If Incoming = "-1" Then
        YesOrNo = "YES"
        ElseIf Incoming = "0" Then
        YesOrNo = "NO"
        Else
        YesOrNo = "ERROR"
    End If
End Function
Function YOrN(Incoming) As String
    'Takes a Yes or No field equivalent and converts it to either y or n
    If Incoming = "-1" Then
        YOrN = "Y"
        ElseIf Incoming = "0" Then
        YOrN = "N"
        Else
        YOrN = "N"
    End If
End Function
Public Function GetSuffix(stripString As String)
    'looks at a name field and extracts the suffix for a name
    Dim A, x, y, z As Integer
    Dim StrippedString, StripStringValue, StripCharValue, FixedString As String
    Dim ending As String
    Dim suffix(1 To 7) As String
    suffix(1) = "JR"
    suffix(2) = "SR"
    suffix(3) = "III"
    suffix(4) = "II"
    suffix(5) = "MD"
    suffix(6) = "PHD"
    suffix(7) = "IV"
    
    'StrippedString = Trim(Strip(stripString, ",."))     'uses the strip function to take out commas and periods
    StrippedString = Strip(stripString, ".")
    y = 0    'goes through the field and looks for the suffixes
    Do
        y = y + 1   'array iterations
        A = Len(suffix(y))
        'x = InStr(1, Right(StrippedString, A), suffix(y), 1)  'determines if the suffix is in the field
        x = InStr(1, StrippedString, suffix(y), 1)  'determines if the suffix is in the field
        If x > 0 Then
            GetSuffix = suffix(y)
            If GetSuffix = "JR" Then GetSuffix = "Jr"
            If GetSuffix = "SR" Then GetSuffix = "Sr"
            Exit Do
            'if the suffix is found then it returns that suffix and exits the loop
        End If
    Loop Until y = 7  'goes until there are no more suffixes to try, then moves to the next record
End Function
Public Function RemoveSuffix(ByVal stripString As String)
    Dim A, x, y, z As Integer
    Dim StrippedString, StripStringValue, StripCharValue, FixedString As String
    Dim ending As String
        
    Dim suffix(1 To 7) As String  'suffixes to check for
    suffix(1) = "JR"
    suffix(2) = "SR"
    suffix(3) = "III"
    suffix(4) = "II"
    suffix(5) = "MD"
    suffix(6) = "PHD"
    suffix(7) = "IV"
        
    StrippedString = Trim(Strip(stripString, "."))  'get rid of unwanted characters
    A = Len(StrippedString)
    
    If A < 4 Then
        RemoveSuffix = StrippedString  'I don't want any null fields so I leave what's there if less than 4 chars
        Exit Function
    End If
    y = 0
    Do
        y = y + 1 'y is the variable for the array iterations
        A = Len(suffix(y))
        z = Len(StrippedString) - Len(suffix(y))  'z is the variable that tells how much of the string to keep for fixedstring
        x = InStr(1, Right(StrippedString, A), suffix(y), 1)  'x just tells if the suffix is found
        If x > 0 Then
            FixedString = Left(StrippedString, z)
            'RemoveSuffix = FixedString & Strip(Right(StrippedString, 3), suffix(y)) 'add the stripped string back to the part that didn't change
            RemoveSuffix = FixedString
            
            Exit Do
        Else
            RemoveSuffix = StrippedString  'if not found in the field
        End If
    Loop Until y = 7  'check for each type of suffix
    
End Function

Function FullOrPart(Incoming) As String
    If Incoming = "RPT" Then
        FullOrPart = "P"
    Else
        FullOrPart = "F"
    End If
End Function
Function RegEEType(Incoming) As String
    If Left(Incoming, 1) = "R" Then
        RegEEType = "REG"
    Else
        RegEEType = "REG"
    End If
End Function
Function MilitaryEra(ByVal iMilitary) As String
    Select Case Trim(iMilitary)
        Case Is = 1, 2, 5, 6
            MilitaryEra = "Z"
        Case Is = 3
            MilitaryEra = "VIET"
        Case Is = 4
            MilitaryEra = "OTHVET"
        Case Else
            'must be an error
            MilitaryEra = "ERROR"
    End Select
End Function
Function MilitaryBranch(ByVal iMilitary) As String
    Select Case Trim(iMilitary)
        Case Is = 1, 2, 3, 4
            MilitaryBranch = "Z"
        Case Is = 5
            MilitaryBranch = "AR"
        Case Is = 6
            MilitaryBranch = "IAR"
        Case Else
            'must be an error
            MilitaryBranch = "ERROR"
    End Select
End Function

Function F_Status(ByVal iState, ByVal FILSTAT_S1) As String
    'Exit Function
    
    If IsNull(FILSTAT_S1) Then
        FILSTAT_S1 = "S"
       
    End If
    
    Select Case Trim(iState)
        Case Is = "GA"
            If FILSTAT_S1 = "A" Or FILSTAT_S1 = "B" Then
                F_Status = "S"
                ElseIf FILSTAT_S1 = "C" Or FILSTAT_S1 = "D" Or FILSTAT_S1 = "E" Then
                F_Status = "M"
                ElseIf FILSTAT_S1 = "F" Or FILSTAT_S1 = "G" Then
                F_Status = "N"
                ElseIf FILSTAT_S1 = "I" Or FILSTAT_S1 = "J" Or FILSTAT_S1 = "K" Then
                F_Status = "H"
                Else
                F_Status = FILSTAT_S1
            End If
        Case Is = "MS"
            If FILSTAT_S1 = "S" Then
                F_Status = "A"
            ElseIf FILSTAT_S1 = "M" Then
                F_Status = "B"
            Else: F_Status = FILSTAT_S1
            End If
        Case Is = "AZ"
            If FILSTAT_S1 = "S" Then
                F_Status = "A"
            ElseIf FILSTAT_S1 = "M" Then
                F_Status = "B"
            Else: F_Status = FILSTAT_S1
            End If
        Case Is = "CT"
            If FILSTAT_S1 = "S" Then
                F_Status = "F"
            ElseIf FILSTAT_S1 = "M" Then
                F_Status = "A"
            Else: F_Status = FILSTAT_S1
            End If
        Case Is = "NC"
            If FILSTAT_S1 = "N" Then
                F_Status = "S"
            ElseIf FILSTAT_S1 = "A" Then
                F_Status = "S"
            Else
                F_Status = FILSTAT_S1
            End If
        Case Is = "SC"
            If FILSTAT_S1 = "N" Then
                F_Status = "S"
            ElseIf FILSTAT_S1 = "A" Then
                F_Status = "S"
            Else
                F_Status = FILSTAT_S1
            End If
        Case Is = "TN"
            If FILSTAT_S1 = "A" Or FILSTAT_S1 = "B" Or FILSTAT_S1 = "F" Or FILSTAT_S1 = "G" Then
                F_Status = "S"
            ElseIf FILSTAT_S1 = "C" Or FILSTAT_S1 = "D" Or FILSTAT_S1 = "E" Or FILSTAT_S1 = "I" Or FILSTAT_S1 = "J" Or FILSTAT_S1 = "K" Then
                F_Status = "M"
            Else
                F_Status = FILSTAT_S1
            End If
        Case Else
            F_Status = FILSTAT_S1
            'F_Status = "X"
    End Select
End Function
Function DedGroup(ByVal iSalOrHour, ByVal iForP, ByVal FileNo) As String
    Select Case Trim(iSalOrHour)
        Case Is = "E"
            If iForP = "F" Then
                DedGroup = Trim(CompanyCode(FileNo)) & "SAL"
            Else
                'Must not be full time
                DedGroup = Trim(CompanyCode(FileNo)) & "PT"
            End If
        Case Is = "H"
            If iForP = "F" Then
                DedGroup = Trim(CompanyCode(FileNo)) & "HR"
            Else
                'Must not be full time
                DedGroup = Trim(CompanyCode(FileNo)) & "PT"
            End If
    End Select
End Function
Function EarnGroup(ByVal iSalOrHour, ByVal FileNo) As String
    Select Case Trim(iSalOrHour)
        Case Is = "E"
            EarnGroup = Trim(CompanyCode(FileNo)) & "SAL"
        Case Is = "H"
            EarnGroup = Trim(CompanyCode(FileNo)) & "HR"
    End Select
End Function
Function PayGroup(ByVal iPayGroup, ByVal FileNo) As String
   Select Case Trim(iPayGroup)
        Case Is = "X"
            PayGroup = Trim(CompanyCode(FileNo)) & "SLRY"
        Case Else
            PayGroup = Trim(CompanyCode(FileNo)) & "HRLY"
    End Select
End Function
Function Proper(ByVal iString As String, Optional ByVal iChar As String) As String
    Dim A, B, C, D, x, y, z As Integer
    Dim WorkStr, vCharValue As String
    Dim QuitFunction As Boolean
    If Trim(Len(iString)) = 0 Then
        Proper = iString
        Exit Function
    End If
    z = Len(iString)
    y = 0
    A = 1
    QuitFunction = False
    Do
        If InStr(A, iString, " ", vbTextCompare) Then
            y = InStr(A, iString, " ", vbTextCompare)
            B = y - A
            WorkStr = WorkStr & UCase(Mid(iString, A, 1)) & LCase(Mid(iString, A + 1, B))
        Else
            WorkStr = WorkStr & UCase(Mid(iString, A, 1)) & LCase(Mid(iString, A + 1, z - y))
            QuitFunction = True
        
        End If
        A = y + 1
    Loop Until QuitFunction
    'Reset this boolean so it can be used again
    QuitFunction = False
    y = 0
    A = 1
    D = 1
    'Takes input parameters and puts CAPITAL letter after the Char that is passed
    If Len(Trim(iChar)) = 0 Then
        Proper = Trim(WorkStr)
        Exit Function '
    Else
        For C = 1 To Len(iChar)
            vCharValue = Mid(iChar, D, 1)
            Do
                If InStr(A, WorkStr, vCharValue, vbTextCompare) Then
                    y = InStr(A, WorkStr, vCharValue, vbTextCompare)
                    B = y + 1
                    WorkStr = Left(WorkStr, y) & UCase(Mid(WorkStr, B, 1)) & (Mid(WorkStr, B + 1, z - (y - 1)))
                Else
                    QuitFunction = True
                End If
                A = y + 1
            Loop Until QuitFunction
            D = D + 1 'increment this to move to next char in ichar
            'Reset these variables to perform Instr function with next char
            y = 0
            A = 1
            QuitFunction = False
        Next C
    End If
    Proper = Trim(WorkStr)
End Function
Public Function CompanyCode(EmpNo) As String
    CompanyCode = "JMS  "
End Function
Public Function CompanyEmpNo(EmpNo) As String
    CompanyEmpNo = "JMS  " & Trim(EmpNo)
End Function

Public Function fnGetTimedKey(EmpNo) As String
    Dim AServer As String
    'Exit Function
    'AServer = "Ultiprosvr"
    'AServer = "PRONETSVR7\VS2005"
    'AServer = "AMICKFS2"
    'AServer = "ACDemo"
    'AServer = "192.168.125.26"
    AServer = "ProNetSvr2"
    
    On Error GoTo TimedKeyError
    ' this executes the UltiPro extended stored procedure
    Dim conn1 As New ADODB.Connection
    Dim rsKey As ADODB.Recordset
    Dim vkey As String
    conn1.ConnectionString = "driver={SQL Server};" & _
        "server=" & AServer & ";uid=sa;pwd=nowayman;" & _
        "database=master"
    conn1.ConnectionTimeout = 15
    conn1.CommandTimeout = 20
    conn1.Open
    DoEvents
    Set rsKey = conn1.Execute("declare @vkey char(12) exec master..xp_usg_gettimedkey_VAL @vKey OUTPUT SELECT @vKey") 'master..xp_usg_gettimedkey_VAL @v_Key
    fnGetTimedKey = rsKey.Fields.Item(0)
    DoEvents
    conn1.Close
    Set rsKey = Nothing
    Set conn1 = Nothing
    DoEvents
    Exit Function
TimedKeyError:
    MsgBox Err.Number
    MsgBox Err.Description
End Function

Function CalcEligDate(iType, iDate As String) As String
    Dim BegDate, EligDate As Date
    BegDate = MakeDate(iDate)
    If iType = "I" Then
        EligDate = BegDate + 90
        'Go to beginning of  month
        CalcEligDate = Month(EligDate) & "/01/" & Year(EligDate)
        Exit Function
    End If
    If iType = "DU" Then
        EligDate = BegDate + 120
        'Go to beginning of  month
        CalcEligDate = Month(EligDate) & "/01/" & Year(EligDate)
        Exit Function
    End If
    If iType = "DM" Then
        EligDate = BegDate + 395
        'Go to beginning of  month
        CalcEligDate = Month(EligDate) & "/01/" & Year(EligDate)
        Exit Function
    End If
End Function
Function CalcHrlyPayRate(iRate, Optional iFlag, Optional iShift) As Double
    Dim SubAmt As Single
    SubAmt = 0
    'If iFlag = "F" Then
     '   SubAmt = 0.5
    'End If
    If iShift = "2" Then
        SubAmt = SubAmt + 0.5
    End If
    If iShift = "3" Then
        SubAmt = SubAmt + 1
    End If
    CalcHrlyPayRate = (iRate / 10000) - SubAmt
End Function


Function CalcStartDate(iType, iDate As String) As String
    Dim BegDate, EligDate As Date
    BegDate = MakeDate(iDate)
    If iType = "I" Then
        EligDate = BegDate + 90
        'Go to beginning of  month
        CalcStartDate = Month(EligDate) & "/01/" & Year(EligDate)
        If CalcStartDate < #1/1/1990# Then
            CalcStartDate = #1/1/1990#
        End If
        Exit Function
    End If
    If iType = "DU" Then
        EligDate = BegDate + 120
        'Go to beginning of  month
        CalcStartDate = Month(EligDate) & "/01/" & Year(EligDate)
        If CalcStartDate < #1/1/1990# Then
            CalcStartDate = #1/1/1990#
        End If
        Exit Function
    End If
    If iType = "DM" Then
        EligDate = BegDate + 395
        'Go to beginning of  month
        CalcStartDate = Month(EligDate) & "/01/" & Year(EligDate)
        If CalcStartDate < #1/1/1990# Then
            CalcStartDate = #1/1/1990#
        End If
        Exit Function
    End If
End Function
Function BenCalcRule(iDedCode) As String
    Select Case Trim(iDedCode)
        Case Is = "15", "18", "26"
            BenCalcRule = "43" 'Annual Salary  * Pct
        'Case Is = "17"
        '    BenCalcRule = "41" 'Weekly Salary  * Pct
        Case Is = "29" ', "21", "22"
            BenCalcRule = "30" 'Flat Amount
        Case Else
            BenCalcRule = "00"
    End Select
End Function
Function Delete_TRADD(EmpNo) As String
Dim db As Database
Dim SQL As String

Set db = CurrentDb
SQL = " DELETE * " & _
"FROM dbo_LodEDed " & _
"WHERE EedDedCode = 'TRADD' AND TRIM(EedEEID) = '" & EmpNo & "'"

db.Execute SQL, dbFailOnError
db.Close
Set db = Nothing
End Function

Function EECalcRule(iDedCode) As String

'15 Gross - Taxes * percent
    Select Case Trim(iDedCode)
        Case Is = "81", "83"
            EECalcRule = "11" 'Def Comp $ * Percent
        Case Is = "92", "93"
            EECalcRule = "60" '% of disposable income
        Case Else
            EECalcRule = "20"
    End Select
    '        Select Case Trim(iDedCode)
    '            Case Is = "3", "5", "6", "8", "19"
    '                EECalcRule = "21" 'Option rate schedule
    '            Case Is = "17"
    '                EECalcRule = "30" 'Benefit Amount * Pct
    '            Case Is = "15", "18", "26"
    '                EECalcRule = "31" 'Benefit * Age Graded Rate Tables
    '            'Case Is = "39", "74", "85", "88", "89"
    '            '    EECalcRule = "99" 'No calc rule  Tax Levy
    '        End Select
    '    Case Else
    '        EECalcRule = "99" 'None Specified
    'End Select
End Function
Function ERCalcRule(iDolOrPct, iDedCode) As String
'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
'15 Gross - Taxes * percent

    Select Case Trim(iDolOrPct)
        Case Is = "D"
            ERCalcRule = "99" 'none Specified
        Case Is = "P"
            Select Case Trim(iDedCode)
                Case Is = "81"
                    ERCalcRule = "22" 'Percent & EE Calculation
                Case Else
                    ERCalcRule = "99" 'none Specified
            End Select
        Case Is = "N"
            Select Case Trim(iDedCode)
                Case Is = "3", "5", "6", "8", "19"
                    ERCalcRule = "21" 'Option rate schedule
                Case Is = "15", "18", "26"
                    ERCalcRule = "31" 'Benefit * Age Graded Rate Tables
                Case Else
                    ERCalcRule = "99" 'No calc rule  Tax Levy
            End Select
        Case Else
            ERCalcRule = "99" 'None Specified
    End Select
End Function
Function EEPeriodPayCap(iDedCode) As String
    Select Case Trim(iDedCode)
        Case Is = "85", "88", "89"
            EEPeriodPayCap = "50" 'Def comp incl$*Cap pct
        Case Is = ""
            EEPeriodPayCap = "51" 'Def comp incl$-125*Cap pct
        Case Is = "39", "74"
            EEPeriodPayCap = "Y"  'Flat Amount
        Case Else
            EEPeriodPayCap = "N"
    End Select
    
End Function
Function ERPeriodPayCap(iDedCode) As String
    Select Case Trim(iDedCode)
        Case Is = "81"
            ERPeriodPayCap = "50" 'Def comp incl$*Cap pct
        Case Else
            EEPeriodPayCap = "00"
    End Select
    
End Function
Function fnRunProc()
    Append_LodEDep
    Gen_DedCodeList
    GetEarnCode
End Function

Public Function FixDate(iField)
Dim iyear As String
    If iField = #12:00:00 AM# Or Trim(iField) = "" Or IsNull(iField) Then
        IsNull (FixDate)
    
    ElseIf Len(Trim(iField)) = 11 Then  'Corrects dates formated as 01 JAN 2008
        FixDate = CDate(getMonthNum(Mid(iField, 4, 3)) & "/" & Mid(iField, 1, 2) & "/" & Mid(iField, 8, 4))
    
    ElseIf Len(Trim(iField)) = 6 Then   'Corrects PAYCHEX 6 digit date
                    If Mid(iField, 5, 2) < 20 Then
                        iyear = "20" & Mid(iField, 5, 2)
                    Else
                        iyear = "19" & Mid(iField, 5, 2)
                    End If
        FixDate = CDate(Mid(iField, 1, 2) & "/" & Mid(iField, 3, 2) & "/" & iyear)
                
    Else
        FixDate = CDate(iField)
    End If
End Function

Public Function getMonthNum(MonthName As String) As Integer
Dim retVal As Integer
    
    Select Case MonthName
        Case "January", "Jan"
            retVal = 1
        Case "February", "Feb"
            retVal = 2
        Case "March", "Mar"
            retVal = 3
        Case "April", "Apr"
            retVal = 4
        Case "May"
            retVal = 5
        Case "June", "Jun"
            retVal = 6
        Case "July", "Jul"
            retVal = 7
        Case "August", "Aug"
            retVal = 8
        Case "September", "Sep", "Sept"
            retVal = 9
        Case "October", "Oct"
            retVal = 10
        Case "November", "Nov"
            retVal = 11
        Case "December", "Dec"
            retVal = 12
        End Select
        
        getMonthNum = retVal
 
End Function



Function TermReasons(Incoming As String) As String
    Incoming = Trim(Incoming)
    If Len(Trim(Incoming)) = 1 Then
        TermReasons = PadL(Incoming, "0", 2)
    Else
        If Trim(Incoming) = "QUITW" Then
            TermReasons = "01"
        ElseIf Trim(Incoming) = "QUITWO" Then
                TermReasons = "02"
        ElseIf Trim(Incoming) = "POINTS" Then
                TermReasons = "04"
        ElseIf Trim(Incoming) = "NONE" Then
                TermReasons = "Z"
        ElseIf Trim(Incoming) = "TRANS" Then
                TermReasons = "TRO"
        ElseIf Trim(Incoming) = "" Then
                TermReasons = "Z"
        Else
            TermReasons = Incoming
        End If
    End If
        
End Function

Function Message()
    
    MsgBox "Add ID field to AllEmpnums_SSN_Conversion and Run Update_EmplMast_Location before first run.", vbOKCancel + vbCritical, "Add field to table"
           
End Function

'add an ID column to a table
Function CreateAutoNumberFunction()
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim tdf As TableDef
    Dim fld As DAO.Field
    Set tdf = dbs.TableDefs("AllEmpnums_SSN_Conversion")
 
    Set fld = tdf.CreateField("ID", dbLong)
    With fld
        '   Appending dbAutoIncrField to Attributes
        '   tells Jet that it's an Autonumber field
        .Attributes = .Attributes Or dbAutoIncrField 'I don't understand this
    End With
    With tdf.Fields
        .Append fld
        .Refresh
    End With
    
    Set tdf = Nothing
    Set dbs = Nothing
    Exit Function
End Function

Function Magic(iCityWorkIn As String, iState As String, iCallingField As String) As String
    Dim db As Database
    Dim rst As Recordset
    Dim rstNewNum As Recordset
    Dim SQL As String
    Dim NewNum As Integer
    'If iState = "OR" Then MsgBox iState
    Select Case iCallingField
        Case Is = "Location"
            Set db = CurrentDb
            If iCityWorkIn = "" Then
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = rst![loccode]
                End If
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = rst![loccode]
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "OccCode"
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITOCCCode]), "", rst![LocLITOCCCode])
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "OtherCode"
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITOtherCode]), "", rst![LocLITOtherCode])
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "ResidentCode"
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITResWorkInCode]), "", rst![LocLITResWorkInCode])
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "ResidentCounty"
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITResWorkInCode]), "", IIf(IsNull(rst![LocLITWorkInCounty]), "", rst![LocLITWorkInCounty]))
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "SDCode"
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITSDCode]), "", rst![LocLITSDCode])
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "WCCCode"
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITWCCCode]), "", rst![LocLITWCCCode])
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "WorkInCode"
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITNonResWorkInCode]), "", rst![LocLITNonResWorkInCode])
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "WorkInCounty"
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITNonResWorkInCode]), "", IIf(IsNull(rst![LocLITWorkInCounty]), "", rst![LocLITWorkInCounty]))
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "TransTypeLITR" 'Res local Tax Code
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITResWorkInCode]), "", "A")
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "TransTypeLITS" 'school district
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITSDCode]), "", "A")
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "TransTypeLITC" 'WCC
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITWCCCode]), "", "A")
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
        Case Is = "TransTypeLITH" 'Other
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITOtherCode]), "", "A")
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "TransTypeLITO" 'OCC
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITOCCCode]), "", "A")
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
    Case Is = "TransTypeLITW" 'LIT Work-in
            Set db = CurrentDb
            If iCityWorkIn = "" Then 'do nothing
                Magic = ""
                Exit Function
            Else
                SQL = " SELECT * from Locations where Left([LocSITWorkInStateCode], 2) = '" & iState & "' AND " & _
                    "ADPLocalCode = '" & iCityWorkIn & "'"
                Set rst = db.OpenRecordset(SQL, dbOpenDynaset)
                If rst.RecordCount = 0 Then
                    rst.Close
                    db.Close
                    Set rst = Nothing
                    Set db = Nothing
                    Magic = "ERROR"
                    Exit Function
                Else
                    Magic = IIf(IsNull(rst![LocLITNonResWorkInCode]), "", "A")
                End If
            End If
            rst.Close
            db.Close
            Set rst = Nothing
            Set db = Nothing
        
        Case Else
            Magic = "Invalid Calling Field"
    End Select
  
End Function

