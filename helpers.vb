Module helpers
    Friend Function BooleanToString(ByVal boolValue As Boolean, ByVal TrueStr As String, ByVal FalseStr As String) As String
        ' checked: v2.0
        If boolValue Then
            Return TrueStr
        Else
            Return FalseStr
        End If
    End Function

    Friend Function protectedDivisionDouble(ByVal theNominator As Double, ByVal theDenominator As Double) As Double
        ' checked: v2.0
        ' this function handles divisionPerZero checks
        ' - return theNominator / theDenominator if all OK
        ' - return divPerZeroIndicatorFlag if (theDenominator=0.0 or divPerZeroIndicatorFlag) or (theNominator=divPerZeroIndicatorFlag)

        If (theNominator = divPerZeroIndicatorFlag) _
        Or (theDenominator = divPerZeroIndicatorFlag) _
        Or (theDenominator = 0.0) Then
            Return divPerZeroIndicatorFlag
        Else
            Return theNominator / theDenominator
        End If
    End Function

    Friend Function getStrAfterProtectedDivision(ByVal theValue As Double, Optional ByVal theIndicatorTxt As String = divPerZeroIndicatorTxt) As String
        ' checked: v2.0
        If theValue = divPerZeroIndicatorFlag Then
            Return theIndicatorTxt
        Else
            Return theValue.ToString
        End If
    End Function

    Friend Function protectedDivisionDoubleToStr(ByVal theNominator As Double, ByVal theDenominator As Double) As String
        ' checked: v2.0
        ' just a makro
        Return getStrAfterProtectedDivision(protectedDivisionDouble(theNominator, theDenominator))
    End Function

    Friend Function getPositionText(ByVal where As positionCategory) As String
        ' checked: v2.0
        Select Case where
            Case positionCategory.inRange
                Return "in_range"
            Case positionCategory.belowRange
                Return "below_range"
            Case positionCategory.aboveRange
                Return "above_range"
            Case Else
                Return ""
        End Select
    End Function

    Friend Function getShortFilename(ByVal aName As String) As String
        ' checked: v2.1
        Dim dummy As String() = Split(aName, "\")
        Return dummy(dummy.Count - 1)
    End Function

    Friend Function getShortFilenameWithoutSuffix(ByVal aName As String) As String
        ' checked: v2.1
        Dim shortName As String = getShortFilename(aName)
        Dim dummy As String()
        Dim l As Integer

        If InStr(shortName, ".") > 0 Then
            dummy = Split(shortName, ".")
            l = Len(dummy(dummy.Count - 1)) + 1
            Return Left(shortName, Len(shortName) - l)
        Else
            Return shortName
        End If
    End Function

    Friend Function getFolder(ByVal aName As String) As String
        ' checked: v2.1
        Dim dummy As String() = Split(aName, "\")
        dummy(dummy.Count - 1) = ""
        Return Join(dummy, "\")
    End Function
End Module
