Module definitions

    '=== Program Information =======================================================================
    Friend Const _
        progName = "RangeChecker", _
        progID = "RC", _
        progVersion_val = 2.21, _
        progVersion_txt = "2.2.1", _
        progFirstDate = "19.02.2005", _
        progLastDate = "18.01.2012", _
        progLastYear = "2012", _
        progCitationReference = "Fritschi J, Raddatz-Müller P, Schmid P, Wuillemin WA." & CR & LF & _
        " Patient self-management of long-term oral anticoagulation in Switzerland." & CR & LF & _
        " Swiss Med Wkly 2007;137(17-18):252-8"

    '=== generic constants =========================================================================
    Friend Const _
        CR = Chr(13), _
        LF = Chr(10), _
        TAB = Chr(9), _
        SPACE = " "

    '--- inputParser -------------------------------------------------------------------------------
    Friend Enum rcParseResult As Byte
        ' checked: v2.0
        rcNothing = 0
        rcContinue = 1
        rcError = 2
        rcFatalError = 3
    End Enum

    '--- dateParser --------------------------------------------------------------------------------
    Friend Const _
        dateFormat = "DD.MM.YYYY", _
        dateDelimiter = ".", _
        dateDayNr = 0, _
        dateMonthNr = 1, _
        dateYearNr = 2, _
        dateStdDeltaYears = 30, _
        dateMinYearForOADate = 1900

    '=== RESULT REPORTING SETTINGS =================================================================
    Friend Const _
        resultDelimiter = TAB

    Friend Const _
        divPerZeroIndicatorFlag As Double = -10000, _
        divPerZeroErrorWrongUsageFlag As Double = -20000, _
        divPerZeroIndicatorTxt As String = "DivPerZero", _
        divPerZeroErrorWrongUsageTxt As String = "DivPerZeroError", _
        notAvailableIndicatorTxt As String = "NotAvailable"

    Friend Enum printMode As Byte
        empty = 0
        labels = 1
        data = 2
    End Enum

    Friend Enum dataValidationResults As Byte
        empty = 0
        data_ok = 1
        data_error = 2
    End Enum

    '=== MODEL =====================================================================================
    Friend Enum positionCategory As SByte
        belowRange = -1
        inRange = 0
        aboveRange = 1
    End Enum

    Friend Const _
        minMin As Double = 0.0, _
        maxMax As Double = 10.0, _
        stdSafetyMin As Double = 2.0, _
        stdSafetyMax As Double = 4.5, _
        stdTimeLimit = 100, _
        minTimeLimit = 14, _
        smallDeltaValue As Double = 0.00000000000001 ' from the help system: by default, a Double value contains 15 decimal digits of precision 
    'smallDeltaValue As Double = 0.000000001 
    'smallDeltaValue As Double = Double.Epsilon
    'smallDeltaValue = 0.00001

    Friend Const _
        emptyDataSetName As String = "error_empty"

    Friend Enum workerTypes As Long
        ' checked: v2.0
        WORKER_Cancelled = -1
        WORKER_Nothing = 0

        ANALYSIS_doRun = 1
        ANALYSIS_resultOK = 2
        ANALYSIS_resultERROR = 3
    End Enum
End Module
