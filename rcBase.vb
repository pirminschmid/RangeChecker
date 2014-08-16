'---------------------------------------------------------------------------------------------------
' RangeChecker
'
' This software mainly calculates the time ratio while measured values are in a predefined range.
' A linear model is used to estimate the values between the measurements. Additional parameters
' have been added since the conception of this program (see below).
'
' RangeChecker v2.2.1 (18-Jan-2012) (c) Pirmin Schmid, mailbox@pirmin-schmid.ch,
'                                                      pirmin.schmid@gmail.com
'
' This software is Freeware but not Public Domain. An acknowledgment in publications is appreciated.
'
' See LICENSE file for additional information.
'
' current functionality:
' - import and validation of data
' - currently a linear model is implemented
' - calculation of time and # of data points in the complete time range
' - calculation for in, below and above individual range, and in, below and above safety range:
'   * # of data points, % in respect to all analyzable
'   * time, % in respect to all analyzable
' - calculation for below and above individual range, and below and above safety range:
'   * deviation from border using an AUC model
' - excluded are
'   * time ranges longer than a cutoff limit
'   * time ranges and data points with bridging
'
' current data input format: tab spaced text (exported from Microsoft Excel)
' @       @       CHECK   param   value
' @       @       SET     param   value
' #       #       UPN     min     max
' date	  INR
' date	  INR     B                       (comment: first day of bridging)
' date	  INR     B                       (comment: last day of bridging; additionals in between allowed)
' date	  INR
' #       #       UPN     min     max
' date	  INR
' date	  INR
' #       #       UPN     min     max     (comment: for empty)
' #       #       UPN     min     max
' date	  INR
' date	  INR
'
' current data output into the user's main document folder:
' log file (text)
' result file (tab spaced text, can be imported into Microsoft Excel or a statistics program)
'
' current parameters that can be set
' CHECK   PROGRAM_ID   RC
' CHECK   MIN_VERSION  2.1
' SET     SAFETY_MIN  value (preset: 2.0)
' SET     SAFETY_MAX  value (preset: 4.5)
' SET     MAX_TIME_INTERVAL  value (preset: 100)
' SET     DELTA_YEARS  value (preset: 30)
'
' history:
' - RangeChecker v1.0 (19-Feb-2005): hello world! (linear model implemented)
' - RangeChecker v1.1 (24-Feb-2005): added data controlling & comfort functions
'   to increase usability (_ratio_percent)
' - RangeChecker v1.2 (26-Feb-2005): added "counting values in range" algorithm,
'   added Tools to check data integrity before using main functions (see addon)...
' - RangeChecker v1.3 (08-Dec-2005): new data integrity check v1.1
' - RangeChecker v1.4 (07-Jan-2006): added t_lowerthanrange, t_higherthanrange,
'   dev_lowerthanrange, dev_higherthanrange, selection mechanism and workaround
'   for actual setting (all basing on linear model), integrated debug_checkrange
' - RangeChecker v1.5 (17-Jan-2007): added maxTimeInterval cutoff possibility for
'   datasets in which data entries are missing over a long time period
'
' - RangeChecker v2.0 (12-Dec-2010): moved from VBA for Excel to Visual Basic 2008 (.NET v3.5)
'   added bridging management. Main work was to implement a proper data import from tab spaced
'   spreadsheets.
' - RangeChecker v2.02 (18-Jan-2011): moved to Visual Basic 2010 (.NET v4.0)
'   test validation with data
' - RangeChecker v2.1 (06-Feb-2011 Super Bowl Edition): handle multiple files (including mainLog 
'   and individualLogs); calc mean / median of delta t between measurements
' - RangeChecker v2.1.1 (01-Mar-2011): create also a summary result file.
' - RangeChecker v2.2 (26-Oct-2011): check for empty UPN; output table legend only once in summary
'   result file; check occurence of UPN (multiples/missing)
' - RangeChecker v2.2.1 (18-Jan-2012): another quality check before running on real data of the
'   second PS-OAK study (to be published)
'---------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------
' Visual Basic 2008 / .NET 3.5 type ranges:
' - Byte         8bit   0 through 255
' - Short       16bit   -32,768 through 32,767
' - Integer     32bit   -2,147,483,648 through 2,147,483,647
' - Long        64bit   -9,223,372,036,854,775,808 through 9,223,372,036,854,775,807 (9.2...E+18)
' - Single      32bit   -3.4028235E+38 through -1.401298E-45 for negative values and from 1.401298E-45 through 3.4028235E+38 for positive values
' - Double      64bit   -1.79769313486231570E+308 through -4.94065645841246544E-324 for negative values and from 4.94065645841246544E-324 through 1.79769313486231570E+308 for positive values
'---------------------------------------------------------------------------------------------------

Imports System
Imports System.IO
Imports System.ComponentModel
Imports System.Threading

'---------------------------------------------------------------------------------------------------
'--- outputClass -----------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Friend Class outputClass
    ' checked: v2.2
    ' modified to handle multiple log files in addition to multiple result files
    ' currently: one summary log and summary result file, and individual logs and result files are created per input file
    Private isInitialized As Boolean = False
    Private mainLogStream As StreamWriter = Nothing
    Private individualLogStream As StreamWriter = Nothing
    Private mainResultStream As StreamWriter = Nothing
    Private individualResultStream As StreamWriter = Nothing

    Private resultLine As String = ""
    Private labelPrefix As String = ""
    Private theBase As rcBase = Nothing

    Private labelsMode As Boolean = False
    Private labelsPrinted As Boolean = False

    Friend Enum logSelector As Byte
        mainLog = 1
        individualLog = 2
        bothLogs = 3
    End Enum

    '--- new / finalize ----------------------------------------------------------------------------
    Friend Sub New(ByVal baseRef As rcBase)
        ' checked: v2.0
        theBase = baseRef
    End Sub

    Protected Overrides Sub Finalize()
        ' checked: v2.1.1
        cleanup()   ' just in case there was no clean exit
        MyBase.Finalize()
    End Sub

    '--- init / cleanup ----------------------------------------------------------------------------
    Friend Function initialize() As Boolean
        ' checked: v2.1.1
        If Not isInitialized Then
            Try
                mainLogStream = New StreamWriter(System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "RangeCheckerMainLog.txt"))
            Catch
                Return False
            End Try

            writeLogHeader(logSelector.mainLog)
            writeLog("Main log file", logSelector.mainLog)
            writeLog("", logSelector.mainLog)

            Try
                mainResultStream = New StreamWriter(System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "RangeCheckerMainResult.txt"))
            Catch
                writeLog("*** ERROR: could not open main result file.", logSelector.mainLog)
                Return False
            End Try
            isInitialized = True
        End If
        ' default
        Return True
    End Function

    Friend Sub cleanup()
        ' checked: v2.1.1
        If isInitialized Then
            Try
                If individualResultStream IsNot Nothing Then
                    individualResultStream.Close()
                    individualResultStream = Nothing
                End If
            Catch ex As Exception
                ' nothing yet, just don't let it break
            End Try

            Try
                If individualLogStream IsNot Nothing Then
                    individualLogStream.Close()
                    individualLogStream = Nothing
                End If
            Catch ex As Exception
                ' nothing yet, just don't let it break
            End Try

            Try
                If mainResultStream IsNot Nothing Then
                    mainResultStream.Close()
                    mainResultStream = Nothing
                End If
            Catch ex As Exception
                ' nothing yet, just don't let it break
            End Try

            Try
                If mainLogStream IsNot Nothing Then
                    mainLogStream.Close()
                    mainLogStream = Nothing
                End If
            Catch ex As Exception
                ' nothing yet, just don't let it break
            End Try
            isInitialized = False
        End If
    End Sub

    Private Sub writeLogHeader(ByVal whichLog As logSelector)
        ' checked: v2.1
        writeLog("------------------------------------------------------------------------------------", whichLog)
        writeLog(" " & progName & " v" & progVersion_txt & " (created " & progFirstDate & " / updated " & progLastDate & ")", whichLog)
        writeLog(" Freeware, (c) 2005-" & progLastYear & " Pirmin Schmid, www.pirmin-schmid.ch, pirmin.schmid@gmx.net", whichLog)
        writeLog("", whichLog)
        writeLog(" Please cite:", whichLog)
        writeLog(" " & progCitationReference, whichLog)
        writeLog("------------------------------------------------------------------------------------", whichLog)
        writeLog("", whichLog)
    End Sub

    Friend Function openResultAndLogStreams(ByVal fileName As String) As Boolean
        ' checked: v2.1.1
        If fileName = "" Then
            writeLog("*** Error in outputClass.openResultStream(): Missing fileName.", logSelector.mainLog)
            Return False
        End If

        fileName = "_" & fileName

        Try
            If individualResultStream IsNot Nothing Then
                individualResultStream.Close()
                individualResultStream = Nothing

                If individualLogStream IsNot Nothing Then
                    writeLog("*** Result and log files closed.")
                    writeLog()
                    individualLogStream.Close()
                    individualLogStream = Nothing
                Else
                    writeLog("*** Result file closed; ERROR: log file missing.", logSelector.mainLog)
                    writeLog("", logSelector.mainLog)
                End If
            End If
        Catch ex As Exception
            ' nothing
            ' just catch the situation when resultStream was not open before
        End Try

        ' open log file
        Try
            individualLogStream = New StreamWriter(System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "RangeCheckerResult" & fileName & "_log.txt"))
            writeLogHeader(logSelector.individualLog)
        Catch ex As Exception
            writeLog("*** ERROR: outputClass.openResultAndLogStreams(): Individual log file '" & System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "RangeCheckerResult" & fileName & "_log.txt") & "' could not be created (" & ex.Message & ").", logSelector.mainLog)
            Return False
        End Try

        ' open result file
        Try
            resultLine = ""
            individualResultStream = New StreamWriter(System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "RangeCheckerResult" & fileName & ".txt"))
        Catch ex As Exception
            writeLog("*** ERROR: outputClass.openResultAndLogStreams(): Result file '" & System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "RangeCheckerResult" & fileName & ".txt") & "' could not be created (" & ex.Message & ").")
            Return False
        End Try

        writeLog("*** New result output file opened '" & System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "RangeCheckerResult" & fileName & ".txt") & "'")
        writeLog("Header")
        Return True
    End Function

    Friend Sub writeLog(Optional ByVal aTxt As String = "", Optional ByVal whichLog As logSelector = logSelector.bothLogs)
        ' checked: v2.1
        If (whichLog = logSelector.bothLogs) OrElse (whichLog = logSelector.individualLog) Then
            Try
                If individualLogStream IsNot Nothing Then
                    individualLogStream.WriteLine(aTxt)
                End If
            Catch ex As Exception
                ' nothing yet, just don't let it break
                writeLog("ERROR: could not write '" & aTxt & "' into the individual log file because of " & ex.Message, logSelector.mainLog)
            End Try
        End If

        If (whichLog = logSelector.bothLogs) OrElse (whichLog = logSelector.mainLog) Then
            Try
                mainLogStream.WriteLine(aTxt)
            Catch ex As Exception
                ' nothing yet, just don't let it break
            End Try
        End If
    End Sub

    Friend Sub writeLogAndStatus(ByVal theForm As mainForm, ByVal aTxt As String, Optional ByVal isOK As Boolean = True, Optional ByVal whichLog As logSelector = logSelector.bothLogs)
        ' checked: v2.1
        theForm.updateStatus(aTxt, isOK)
        writeLog("", whichLog)
        writeLog(aTxt, whichLog)
    End Sub

    Friend Sub setResultLabelPrefix(ByVal aPrefix As String)
        ' checked: v2.0
        labelPrefix = aPrefix
    End Sub

    Friend Sub appendResultLabel(ByVal aTxt As String, Optional ByVal aDelimiter As String = resultDelimiter)
        ' checked: v2.2
        labelsMode = True
        appendResult(labelPrefix & aTxt, aDelimiter)
    End Sub

    Friend Sub appendResult(ByVal aTxt As String, Optional ByVal aDelimiter As String = resultDelimiter)
        ' checked: v2.0
        If resultLine <> "" Then
            resultLine &= aDelimiter & aTxt
        Else
            resultLine = aTxt
        End If
    End Sub

    Friend Sub newlineResult()
        ' checked: v2.2.1
        Try
            If labelsMode Then
                If Not labelsPrinted Then
                    mainResultStream.WriteLine(resultLine)
                    labelsPrinted = True
                End If
            Else
                mainResultStream.WriteLine(resultLine)
            End If
            individualResultStream.WriteLine(resultLine)
        Catch ex As Exception
            writeLog("*** ERROR: outputClass.newlineResult(): could not write into the result files (" & ex.Message & ").")
        End Try
        resultLine = ""
        labelsMode = False
    End Sub
End Class

'---------------------------------------------------------------------------------------------------
'--- simpleStatList --------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Friend Class simpleStatList
    ' checked: v2.1
    ' currently implemented: mean, median, n, sum, min, max
    Private theList As List(Of Double) = New List(Of Double)
    Private sortedList As Boolean = False

    Friend Sub New()
        ' checked: v2.1
    End Sub

    Private Sub sortList()
        ' checked: v2.1
        If Not sortedList Then
            theList.Sort()
            sortedList = True
        End If
    End Sub

    Friend Sub add(ByVal aValue As Double)
        ' checked: v2.1
        theList.Add(aValue)
        sortedList = False
    End Sub

    Friend Function mean() As Double
        ' checked: v2.1
        Dim sum As Double = 0.0

        If theList.Count > 0 Then
            For Each item As Double In theList
                sum += item
            Next
            Return sum / CDbl(theList.Count)
        Else
            Return 0.0
        End If
    End Function

    Friend Function median() As Double
        ' checked: v2.1
        Dim m As Integer = 0
        Dim n As Integer = 0

        If theList.Count > 0 Then
            sortList()
            If theList.Count Mod 2 = 0 Then
                n = theList.Count \ 2
                m = n - 1
                Return (theList(m) + theList(n)) / 2.0
            Else
                m = Math.Floor(CDbl(theList.Count) / 2.0)
                Return theList(m)
            End If
            m = CDbl(theList.Count) / 2.0
        Else
            Return 0.0
        End If
    End Function

    Friend Function n() As Integer
        ' checked: v2.1
        Return theList.Count
    End Function

    Friend Function sum() As Double
        ' checked: v2.1
        Dim theSum As Double = 0.0

        If theList.Count > 0 Then
            For Each item As Double In theList
                theSum += item
            Next
            Return theSum
        Else
            Return 0.0
        End If
    End Function

    Friend Function min() As Double
        ' checked: v2.1
        If theList.Count > 0 Then
            sortList()
            Return theList(0)
        Else
            Return 0.0
        End If
    End Function

    Friend Function max() As Double
        ' checked: v2.1
        If theList.Count > 0 Then
            sortList()
            Return theList(theList.Count - 1)
        Else
            Return 0.0
        End If
    End Function
End Class

'---------------------------------------------------------------------------------------------------
'--- date converter --------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Friend Class dateReader
    ' checked: v2.0
    ' this little reader was implemented to avoid worrying about IFormatProvider in Date.Parse()
    ' a controlled range of years was implemented additionally
    Private dDelimiter As String
    Private dDayNr As Byte
    Private dMonthNr As Byte
    Private dYearNr As Byte
    Private dMinYear As Integer
    Private dMaxYear As Integer

    Friend Sub New(ByVal aDelimiter As String, ByVal aDayNr As Byte, ByVal aMonthNr As Byte, ByVal aYearNr As Byte, ByVal aMinYear As Integer, ByVal aMaxYear As Integer)
        ' checked: v2.0
        dDelimiter = aDelimiter
        dDayNr = aDayNr
        dMonthNr = aMonthNr
        dYearNr = aYearNr
        dMinYear = aMinYear
        dMaxYear = aMaxYear
    End Sub

    Friend Function getDate(ByVal aString As String, ByRef theDate As Date) As Boolean
        ' checked: v2.0
        Dim splitString As String()
        Dim day As Integer
        Dim month As Integer
        Dim year As Integer
        Dim max_day As Integer

        Try
            splitString = Split(aString, dDelimiter)
            day = CInt(splitString(dDayNr))
            month = CInt(splitString(dMonthNr))
            year = CInt(splitString(dYearNr))

            If (year >= dMinYear) And (year <= dMaxYear) _
            And (month >= 1) And (month <= 12) Then
                Select Case month
                    Case 2
                        If Date.IsLeapYear(year) Then
                            max_day = 29
                        Else
                            max_day = 28
                        End If

                    Case 4, 6, 9, 11
                        max_day = 30

                    Case Else
                        max_day = 31
                End Select
                If (day >= 1) And (day <= max_day) Then
                    theDate = New Date(year, month, day)
                    Return True
                End If
            End If

        Catch ex As Exception
            Return False
        End Try
        ' default
        Return False
    End Function
End Class

'---------------------------------------------------------------------------------------------------
'--- data structure / analysis ---------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Friend Class dataItem
    ' checked: v2.0
    Friend theDate As Double
    Friend theValue As Double
    Friend theBridgingFlag As Boolean
    Friend theLineNr As Long

    Friend Sub New(ByVal aDate As Date, ByVal aValue As Double, ByVal aBridgingFlag As Boolean, ByVal aLineNr As Long)
        ' checked: v2.0
        theDate = aDate.ToOADate
        theValue = aValue
        theBridgingFlag = aBridgingFlag
        theLineNr = aLineNr
    End Sub
End Class

Friend Class dataResults
    ' checked: v2.1
    Private theBase As rcBase
    Private theDataItems As List(Of dataItem)

    Private prefix As String
    Private main As Boolean
    Private min As Double
    Private max As Double
    Private timeLimit As Double

    Private numberOverall As Integer = 0
    Private numberBridging As Integer = 0
    Private numberAnalyzable As Integer = 0
    Private numberIn As Integer = 0
    Private numberBelow As Integer = 0
    Private numberAbove As Integer = 0

    Private timeOverall As Double = 0.0
    Private timeBridging As Double = 0.0    ' excluded due to bridging
    Private timeClipped As Double = 0.0     ' excluded due to timeLimit
    Private timeAnalyzable As Double = 0.0
    Private timeIn As Double = 0.0
    Private timeBelow As Double = 0.0
    Private timeAbove As Double = 0.0

    Private timeMeanBetweenSampling As Double = 0.0     ' only re analyzable time
    Private timeMedianBetweenSampling As Double = 0.0   ' only re analyzable time

    Private deviationFromBorderAbove = 0.0
    Private deviationFromBorderBelow = 0.0

    Private numberOfClippedTimeperiods As Integer = 0
    Private numberOfAnalyzableTimeperiods As Integer = 0

    Friend Sub New(ByVal baseRef As rcBase, ByVal aPrefix As String, ByVal aMin As Double, ByVal aMax As Double, ByVal aTimeLimit As Integer, ByVal isMain As Boolean)
        theBase = baseRef
        main = isMain
        prefix = aPrefix
        min = aMin
        max = aMax
        timeLimit = CDbl(aTimeLimit)
    End Sub

    '=== LINEAR MODEL helper functions =============================================================
    ' linear model:
    ' y = f(x) = ax + b
    '
    ' formulas, assuming P1(x1,y1) and P2(x2,y2) and x1 <> x2
    '
    ' a = (y2 - y1) / (x2 - x1)
    ' b = y1 - (a * x1)
    ' x_cut = (y_cut - b) / a
    '===============================================================================================

    Private Function linear_getA(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
        Return ((y2 - y1) / (x2 - x1))
    End Function

    Private Function linear_getB(ByVal x As Double, ByVal y As Double, ByVal a As Double) As Double
        Return (y - a * x)
    End Function

    Private Function linear_getX_cut(ByVal y_cut As Double, ByVal a As Double, ByVal b As Double) As Double
        Return ((y_cut - b) / a)
    End Function

    '--- (A) time ----------------------------------------------------------------------------------
    Private Sub calcFullInterval()
        ' checked: v2.1
        ' old (before consideration of time clipping): calcInterval = CDbl(lastDate) - CDbl(firstDate)
        ' current:
        ' - overall (see above)
        ' - analyzable
        ' - bridging
        ' - exclusion due to clipping
        Dim deltaT As Double
        Dim deltaTList As simpleStatList = New simpleStatList()

        For i = 0 To theDataItems.Count - 2
            deltaT = theDataItems(i + 1).theDate - theDataItems(i).theDate
            If (theDataItems(i + 1).theBridgingFlag) And (theDataItems(i).theBridgingFlag) Then
                timeBridging += deltaT
            ElseIf deltaT <= timeLimit Then
                numberOfAnalyzableTimeperiods += 1
                timeAnalyzable += deltaT
                deltaTList.add(deltaT)
            Else
                numberOfClippedTimeperiods += 1
                timeClipped += deltaT
            End If
        Next

        timeOverall = theDataItems(theDataItems.Count - 1).theDate - theDataItems(0).theDate

        timeMeanBetweenSampling = deltaTList.mean()
        timeMedianBetweenSampling = deltaTList.median()

        ' just a test
        If timeOverall <> (timeAnalyzable + timeBridging + timeClipped) Then
            theBase.out.writeLog("ERROR")
        End If
    End Sub

    Private Function checkRange(ByVal val As Double, ByVal bottom As Double, ByVal top As Double) As SByte
        ' checked: v2.0
        ' returns 0 = in range, -1 = below, +1 = above
        If val > top Then
            Return positionCategory.aboveRange
        ElseIf val < bottom Then
            Return positionCategory.belowRange
        Else
            Return positionCategory.inRange
        End If
    End Function

    Private Function selectCutoff(ByVal where As SByte, ByVal lower As Double, ByVal upper As Double) As Double
        ' checked: v2.0
        ' where MUST be <> 0!

        If where > 0 Then
            Return upper
        Else
            Return lower
        End If
    End Function

    Private Function selectCutoff2(ByVal where_one As SByte, ByVal where_two As SByte, ByVal lower As Double, ByVal upper As Double) As Double
        ' checked: v2.0
        ' where_one / where_two: one of them MUST be <> 0!
        Dim dummy As SByte

        If where_two = 0 Then
            dummy = where_one
        Else
            dummy = where_two
        End If

        If dummy > 0 Then
            Return upper
        Else
            Return lower
        End If
    End Function

    Private Function calcTimeInRange(ByVal which As positionCategory) As Double
        ' checked: v2.0
        Dim topBorder As Double
        Dim bottomBorder As Double
        Dim timeSum As Double = 0.0

        ' linear model
        Dim a As Double
        Dim b As Double

        Dim x_firstDate As Double
        Dim y_firstDate As Double

        Dim x_cut_first As Double
        Dim x_cut_last As Double

        Dim x_lastDate As Double
        Dim y_lastDate As Double

        Dim where_firstDate As SByte
        Dim where_lastDate As SByte

        Dim p As SByte

        Select Case which
            Case positionCategory.inRange
                topBorder = max
                bottomBorder = min

            Case positionCategory.aboveRange
                topBorder = maxMax
                bottomBorder = max + smallDeltaValue

            Case positionCategory.belowRange
                topBorder = min - smallDeltaValue
                bottomBorder = minMin

            Case Else
                Return 0.0
        End Select

        For i = 0 To theDataItems.Count - 2
            x_firstDate = theDataItems(i).theDate
            x_lastDate = theDataItems(i + 1).theDate

            If x_firstDate < x_lastDate Then
                If (theDataItems(i + 1).theBridgingFlag) And (theDataItems(i).theBridgingFlag) Then
                    ' exclude bridged time
                    ' no calculations here, already done
                ElseIf (x_lastDate - x_firstDate) <= timeLimit Then
                    ' calculate this
                    y_firstDate = theDataItems(i).theValue
                    y_lastDate = theDataItems(i + 1).theValue
                    where_firstDate = checkRange(y_firstDate, bottomBorder, topBorder)
                    where_lastDate = checkRange(y_lastDate, bottomBorder, topBorder)
                    p = where_firstDate * where_lastDate

                    If p = 0 Then
                        If where_firstDate + where_lastDate = 0 Then
                            ' both in range --> easy case
                            timeSum += x_lastDate - x_firstDate
                        Else
                            ' there is one crossing
                            a = linear_getA(x_firstDate, y_firstDate, x_lastDate, y_lastDate)
                            b = linear_getB(x_firstDate, y_firstDate, a)
                            x_cut_first = linear_getX_cut(selectCutoff2(where_firstDate, where_lastDate, bottomBorder, topBorder), a, b)
                            If where_lastDate = 0 Then
                                timeSum += x_lastDate - x_cut_first
                            Else ' eq. where_firstDate = 0
                                timeSum += x_cut_first - x_firstDate
                            End If
                        End If
                    ElseIf p < 0 Then
                        ' ie. p=-1 ==> two crossings
                        ' (values ar out of range, but on both sides of the range)
                        a = linear_getA(x_firstDate, y_firstDate, x_lastDate, y_lastDate)
                        b = linear_getB(x_firstDate, y_firstDate, a)
                        x_cut_first = linear_getX_cut(selectCutoff(where_firstDate, bottomBorder, topBorder), a, b)
                        x_cut_last = linear_getX_cut(selectCutoff(where_lastDate, bottomBorder, topBorder), a, b)
                        timeSum += x_cut_last - x_cut_first
                    End If
                    ' the third possibility p=+1 adds no time to timeIn
                    ' (values out of range, but on the same side --> never in range)
                Else
                    ' clipped
                    ' however, no counting here. Was already counted during the calculation of complete interval
                End If
            Else
                ' this should not happen
                Throw New Exception("Error in dataResults.calcTimeInRange(): dataItems are not sorted.")
            End If
        Next
        Return timeSum
    End Function

    '--- (B) values --------------------------------------------------------------------------------
    Private Sub countValues()
        ' checked: v2.0
        numberOverall = theDataItems.Count
        For Each item As dataItem In theDataItems
            If item.theBridgingFlag Then
                numberBridging += 1
            Else
                Select Case checkRange(item.theValue, min, max)
                    Case positionCategory.inRange
                        numberIn += 1
                    Case positionCategory.belowRange
                        numberBelow += 1
                    Case positionCategory.aboveRange
                        numberAbove += 1
                End Select
            End If
        Next item
        numberAnalyzable = numberIn + numberBelow + numberAbove
    End Sub

    '--- (C) deviation from border -----------------------------------------------------------------
    Private Function calcDeviationFromBorder(ByVal which As positionCategory) As Double
        ' checked: v2.0
        ' model:
        ' - sum all AUC (area under the curve) when value is in desired range
        ' - one AUC is calculated by delta_t * values (minus border) / 2 [i.e. t * (c+d) / 2]
        ' - finally meanDeviation = aucSum / timeSum
        ' - mean deviation is related to the time a subject is out of range...
        ' - time with bridging / clipping is excluded from calculation

        Dim timeSum As Double = 0.0
        Dim aucSum As Double = 0.0

        ' linear model
        Dim a As Double
        Dim b As Double
        Dim c As Double
        Dim d As Double
        Dim t As Double

        Dim x_firstDate As Double
        Dim y_firstDate As Double

        Dim x_cut As Double

        Dim x_lastDate As Double
        Dim y_lastDate As Double

        Dim where_firstDate As SByte
        Dim where_lastDate As SByte

        For i = 0 To theDataItems.Count - 2
            x_firstDate = theDataItems(i).theDate
            x_lastDate = theDataItems(i + 1).theDate
            If x_firstDate < x_lastDate Then
                If (theDataItems(i + 1).theBridgingFlag) And (theDataItems(i).theBridgingFlag) Then
                    ' exclude bridged time
                ElseIf (x_lastDate - x_firstDate) <= timeLimit Then
                    y_firstDate = theDataItems(i).theValue
                    y_lastDate = theDataItems(i + 1).theValue
                    where_firstDate = checkRange(y_firstDate, min, max)
                    where_lastDate = checkRange(y_lastDate, min, max)

                    Select Case which
                        Case positionCategory.aboveRange
                            If where_firstDate + where_lastDate = 2 Then
                                c = y_firstDate - max
                                d = y_lastDate - max
                                t = x_lastDate - x_firstDate
                            ElseIf (where_firstDate = 1) Or (where_lastDate = 1) Then
                                ' calc one crossing
                                a = linear_getA(x_firstDate, y_firstDate, x_lastDate, y_lastDate)
                                b = linear_getB(x_firstDate, y_firstDate, a)
                                x_cut = linear_getX_cut(max, a, b)

                                If where_firstDate = 1 Then
                                    c = y_firstDate - max
                                    d = 0.0
                                    t = x_cut - x_firstDate
                                Else
                                    c = 0.0
                                    d = y_lastDate - max
                                    t = x_lastDate - x_cut
                                End If
                            Else
                                ' nothing otherwise
                                t = 0.0
                            End If

                        Case positionCategory.belowRange
                            If where_firstDate + where_lastDate = -2 Then
                                c = min - y_firstDate
                                d = min - y_lastDate
                                t = x_lastDate - x_firstDate
                            ElseIf (where_firstDate = -1) Or (where_lastDate = -1) Then
                                ' calc one crossing
                                a = linear_getA(x_firstDate, y_firstDate, x_lastDate, y_lastDate)
                                b = linear_getB(x_firstDate, y_firstDate, a)
                                x_cut = linear_getX_cut(min, a, b)

                                If where_firstDate = -1 Then
                                    c = min - y_firstDate
                                    d = 0.0
                                    t = x_cut - x_firstDate
                                Else
                                    c = 0.0
                                    d = min - y_lastDate
                                    t = x_lastDate - x_cut
                                End If
                            Else
                                ' nothing otherwise
                                t = 0.0
                            End If

                        Case Else
                            Return 0.0
                    End Select

                    If t > 0.0 Then
                        timeSum += t
                        aucSum += t * (c + d) / 2.0
                    End If
                Else
                    ' exclude clipped time
                End If
            Else
                ' this should not happen
                Throw New Exception("Error in dataResults.calcDeviationFromBorder(): dataItems are not sorted.")
            End If
        Next

        If timeSum > 0.0 Then
            Return (aucSum / timeSum)
        Else
            Return divPerZeroIndicatorFlag
        End If
    End Function

    Private Sub print(ByVal mode As printMode)
        ' checked: v2.1
        With theBase.out
            ' set prefix

            Select Case mode
                Case printMode.labels
                    If main Then
                        .appendResultLabel("values_overall")
                        .appendResultLabel("values_bridging")
                        .appendResultLabel("values_analyzable")

                        .appendResultLabel("mean_time_between_sampling")
                        .appendResultLabel("median_time_between_sampling")

                        .appendResultLabel("time_overall")
                        .appendResultLabel("time_bridging")
                        .appendResultLabel("time_clipped")
                        .appendResultLabel("number_of_clipped_periods")
                        .appendResultLabel("time_analyzable")
                        .appendResultLabel("number_of_analyzable_periods")
                    End If

                    .setResultLabelPrefix(prefix)
                    .appendResultLabel("identifier")

                    .appendResultLabel("values_inRange")
                    .appendResultLabel("percent_values_inRange")
                    .appendResultLabel("values_aboveRange")
                    .appendResultLabel("percent_values_aboveRange")
                    .appendResultLabel("values_belowRange")
                    .appendResultLabel("percent_values_belowRange")

                    .appendResultLabel("time_inRange")
                    .appendResultLabel("percent_time_inRange")

                    .appendResultLabel("time_aboveRange")
                    .appendResultLabel("percent_time_aboveRange")
                    .appendResultLabel("meanDeviationFromBorder_aboveRange")

                    .appendResultLabel("time_belowRange")
                    .appendResultLabel("percent_time_belowRange")
                    .appendResultLabel("meanDeviationFromBorder_belowRange")

                Case printMode.data
                    If main Then
                        .appendResult(numberOverall)
                        .appendResult(numberBridging)
                        .appendResult(numberAnalyzable)

                        .appendResult(timeMeanBetweenSampling)
                        .appendResult(timeMedianBetweenSampling)

                        .appendResult(timeOverall)
                        .appendResult(timeBridging)
                        .appendResult(timeClipped)
                        .appendResult(numberOfClippedTimeperiods)
                        .appendResult(timeAnalyzable)
                        .appendResult(numberOfAnalyzableTimeperiods)
                    End If

                    .appendResult(prefix)

                    .appendResult(numberIn)
                    .appendResult(protectedDivisionDoubleToStr(numberIn * 100, numberAnalyzable))
                    .appendResult(numberAbove)
                    .appendResult(protectedDivisionDoubleToStr(numberAbove * 100, numberAnalyzable))
                    .appendResult(numberBelow)
                    .appendResult(protectedDivisionDoubleToStr(numberBelow * 100, numberAnalyzable))

                    .appendResult(timeIn)
                    .appendResult(protectedDivisionDoubleToStr(timeIn * 100, timeAnalyzable))

                    .appendResult(timeAbove)
                    .appendResult(protectedDivisionDoubleToStr(timeAbove * 100, timeAnalyzable))
                    .appendResult(getStrAfterProtectedDivision(deviationFromBorderAbove, notAvailableIndicatorTxt))

                    .appendResult(timeBelow)
                    .appendResult(protectedDivisionDoubleToStr(timeBelow * 100, timeAnalyzable))
                    .appendResult(getStrAfterProtectedDivision(deviationFromBorderBelow, notAvailableIndicatorTxt))

                Case printMode.empty
                    If main Then
                        For i = 1 To 11
                            .appendResult(notAvailableIndicatorTxt)
                        Next
                    End If

                    .appendResult(prefix)
                    For i = 1 To 14
                        .appendResult(notAvailableIndicatorTxt)
                    Next
            End Select
        End With
    End Sub

    Friend Sub printLabels()
        ' checked: v2.0
        print(printMode.labels)
    End Sub

    Friend Sub printEmpty()
        ' checked: v2.0
        print(printMode.empty)
    End Sub

    Friend Function analyze(ByRef someDataItems As List(Of dataItem)) As Boolean
        ' checked: v2.0
        Try
            theDataItems = someDataItems

            calcFullInterval()
            timeIn = calcTimeInRange(positionCategory.inRange)
            timeAbove = calcTimeInRange(positionCategory.aboveRange)
            timeBelow = calcTimeInRange(positionCategory.belowRange)

            countValues()

            deviationFromBorderAbove = calcDeviationFromBorder(positionCategory.aboveRange)
            deviationFromBorderBelow = calcDeviationFromBorder(positionCategory.belowRange)

        Catch ex As Exception
            theBase.out.writeLog("Error in dataResults.analyze(): " & ex.Message)
            Return False
        End Try

        ' default
        print(printMode.data)
        Return True
    End Function
End Class

Friend Class dataSet
    ' checked: v2.0
    Private theBase As rcBase
    Private theUPN As String

    Private theIndividualMin As Double
    Private theIndividualMax As Double

    Private dataItems As List(Of dataItem) = New List(Of dataItem)
    Private results As List(Of dataResults) = New List(Of dataResults)

    Private isEmpty As Boolean = True
    Private isError As Boolean = False

    Friend Sub New(ByVal baseRef As rcBase, ByVal aUPN As String, ByVal individualMin As Double, ByVal individualMax As Double)
        ' checked: v2.2
        Dim found As Boolean = False
        Dim dummy As Long = 0

        theBase = baseRef
        theUPN = aUPN
        theIndividualMin = individualMin
        theIndividualMax = individualMax

        ' define analysis functions
        With baseRef
            results.Add(New dataResults(baseRef, "individual_", individualMin, individualMax, .timeLimit, True))
            results.Add(New dataResults(baseRef, "safety_", .safetyMin, .safetyMax, .timeLimit, False))
        End With

        If theUPN = "" Then
            theBase.out.writeLog("# DATA ERROR in line " & theBase.lineNr & ": empty UPN definition detected." & CR & LF)
            theUPN = " "    ' to keep the output table in shape
            setErrorFlag()
        ElseIf theUPN = emptyDataSetName Then
            ' nothing
        Else
            For Each knownUPN As String In theBase.upnList_Str
                If theUPN = knownUPN Then
                    theBase.out.writeLog("# DATA ERROR in line " & theBase.lineNr & ": multiple occurrences of the UPN '" & theUPN & "' detected. Check all files of this analysis." & CR & LF)
                    setErrorFlag()
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                theBase.upnList_Str.Add(theUPN)
                Try
                    dummy = CLng(theUPN)
                    theBase.upnList_Int.Add(dummy)
                Catch ex As Exception
                    theBase.upnList_allNumbers = False
                End Try
            End If
        End If

        If Not (individualMin < individualMax) Then
            theBase.out.writeLog("# DATA ERROR re UPN '" & theUPN & "': INDIVIDUAL_MIN (" & theIndividualMin & ") is not lower than INDIVIDUAL_MAX (" & theIndividualMax & ")." & CR & LF)
            setErrorFlag()
        End If
    End Sub

    Friend Sub setErrorFlag()
        ' checked: v2.0
        isError = True
    End Sub

    Friend Function getErrorFlag() As Boolean
        ' checked: v2.0
        Return isError
    End Function

    Friend Sub addDataItem(ByVal aDate As Date, ByVal aValue As Double, ByVal aBridgingFlag As Boolean)
        ' checked: v2.0
        dataItems.Add(New dataItem(aDate, aValue, aBridgingFlag, theBase.lineNr))
    End Sub

    Private Function validateDataItems() As dataValidationResults
        ' checked: v2.1.1
        Dim previousDate As Double = 0.0    ' = midnight, 30 December 1899
        Dim retVal As dataValidationResults = dataValidationResults.data_ok
        Dim j As Integer = 1
        Dim increment As Double = 0.0

        If dataItems.Count = 0 Then
            retVal = dataValidationResults.empty
        Else
            isEmpty = False
        End If

        If getErrorFlag() Then
            retVal = dataValidationResults.data_error
        End If

        ' handle multiple INR measurements at one day
        ' principle: divide 24h per # of measurements at the same day
        j = 1
        For i = 0 To dataItems.Count - 2
            Do While ((i + j) < dataItems.Count) AndAlso (dataItems(i).theDate = dataItems(i + j).theDate)
                j += 1
            Loop
            If j > 1 Then
                ' adjust the data values
                increment = 1.0 / CDbl(j)
                For k = i + 1 To i + j - 1
                    dataItems(k).theDate = dataItems(k - 1).theDate + increment
                Next

                ' report the modifications
                theBase.out.writeLog(". Note re UPN '" & theUPN & "' in lines " & dataItems(i).theLineNr & " to " & dataItems(i + j - 1).theLineNr & ": The date is identical in these " & j & " lines.")
                theBase.out.writeLog("  This may reflect multiple measurements at the same day. Time information was adjusted to use a delta-t of " & 24 / j & " hours.")
                theBase.out.writeLog("  Nevertheless, check the original file for typos of the dates in these lines." & CR & LF)

                ' adjust loop parameters
                i = i + j - 1
                j = 1
            End If
        Next

        ' check data for errors
        previousDate = 0.0
        For Each dummy As dataItem In dataItems
            With dummy
                If previousDate >= .theDate Then
                    theBase.out.writeLog("* DATA ERROR re UPN '" & theUPN & "' in line " & .theLineNr & ": Check date '" & Date.FromOADate(.theDate).ToShortDateString & "' (previous date = " & Date.FromOADate(previousDate).ToShortDateString & ") and surrounding dates/values." & CR & LF)
                    retVal = dataValidationResults.data_error
                    setErrorFlag()
                End If

                If (.theValue <= minMin) Or (.theValue >= maxMax) Then
                    theBase.out.writeLog("* DATA ERROR re UPN '" & theUPN & "' in line " & .theLineNr & ": Check value '" & .theValue & "', and surrounding dates/values." & CR & LF)
                    retVal = dataValidationResults.data_error
                    setErrorFlag()
                End If

                previousDate = .theDate
            End With
        Next dummy

        Return retVal
    End Function

    Friend Sub printLabels()
        ' checked: v2.0
        With theBase.out
            .setResultLabelPrefix("")
            .appendResultLabel("UPN")

            .appendResultLabel("individual_min")
            .appendResultLabel("individual_max")
            .appendResultLabel("safety_min")
            .appendResultLabel("safety_max")
            .appendResultLabel("timeLimit")

            For Each dummy As dataResults In results
                dummy.printLabels()
            Next

            .setResultLabelPrefix("")
            .appendResultLabel("ERROR_flag")
            .appendResultLabel("Empty_flag")
            .newlineResult()
        End With
    End Sub

    Friend Sub analyze()
        ' checked: v2.0

        With theBase.out
            .appendResult(theUPN)

            .appendResult(theIndividualMin)
            .appendResult(theIndividualMax)
            .appendResult(theBase.safetyMin)
            .appendResult(theBase.safetyMax)
            .appendResult(theBase.timeLimit)

            If validateDataItems() = dataValidationResults.data_ok Then
                For Each dummy As dataResults In results
                    If Not dummy.analyze(dataItems) Then
                        dummy.printEmpty()
                        setErrorFlag()
                    End If
                Next
            Else
                For Each dummy As dataResults In results
                    dummy.printEmpty()
                Next
            End If

            .appendResult(BooleanToString(getErrorFlag, "ERROR", "ok"))
            .appendResult(BooleanToString(isEmpty, "empty", "ok"))
            .newlineResult()
        End With
    End Sub
End Class

Friend Class inputRCparser
    ' checked: v2.0
    Private theBase As rcBase
    Private theDateReader As dateReader
    Private currentDataset As dataSet
    Private errorDataset As dataSet

    Private errorCounter As Integer = 0
    Private okCounter As Integer = 0

    Private splitLine As String()

    Private parseMode As rcParseModes = rcParseModes.rcHeader
    Private addErrorDatasetCounter As Integer = 0

    Private Enum rcParseModes As Byte
        ' checked: v2.0
        rcNothing = 0
        rcHeader = 1
        rcDatasets = 2
        rcSkipDataset_fromHeader = 3
        rcSkipDataset_fromDatasets = 4
    End Enum

    Public Sub New(ByVal baseRef As rcBase)
        ' checked: v2.0
        theBase = baseRef
        theDateReader = New dateReader(dateDelimiter, dateDayNr, dateMonthNr, dateYearNr, theBase.minYear, theBase.maxYear)
        errorDataset = New dataSet(theBase, emptyDataSetName, theBase.safetyMin, theBase.safetyMax)
        errorDataset.setErrorFlag()
    End Sub

    Private Sub incCounter(Optional ByVal isError As Boolean = False)
        ' checked: v2.0
        If isError Then
            errorCounter += 1
        Else
            okCounter += 1
        End If
    End Sub

    Friend Sub logCounter()
        ' checked: v2.1
        With theBase.out
            .writeLog()
            .writeLog("Dataset statistics")
            .writeLog("------------------")
            .writeLog(okCounter + errorCounter & " analyzed")
            .writeLog(okCounter & " OK")
            .writeLog(errorCounter & " with errors")
            .writeLog()
        End With
    End Sub

    Private Sub helper_addErrorPlaceholders()
        ' checked: v2.0
        Do While addErrorDatasetCounter > 0
            errorDataset.analyze()
            incCounter(errorDataset.getErrorFlag)
            addErrorDatasetCounter -= 1
        Loop
    End Sub

    Private Sub helper_flushBuffer()
        ' checked: v2.0
        currentDataset.analyze()
        incCounter(currentDataset.getErrorFlag)
        helper_addErrorPlaceholders()
    End Sub

    Friend Sub flushBuffer()
        ' checked: v2.0
        If (parseMode = rcParseModes.rcDatasets) Or (parseMode = rcParseModes.rcSkipDataset_fromDatasets) Then
            helper_flushBuffer()
        End If
    End Sub

    Private Function getSimpleBoolean(ByVal pos As Integer, ByVal trueString As String) As Boolean
        ' checked: v2.0
        Dim dummyBool As Boolean = False
        Try
            If splitLine.Count > pos Then
                dummyBool = (splitLine(pos) = trueString)
            End If
        Catch ex As Exception
            Return False
        End Try
        Return dummyBool
    End Function

    Private Function getDate(ByVal pos As Integer, ByRef theDate As Date) As Boolean
        ' checked: v2.1
        Try
            If theDateReader.getDate(splitLine(pos), theDate) Then
                Return True
            Else
                theBase.out.writeLog("    ERROR in line " & theBase.lineNr & ": '" & splitLine(pos) & "' is not a valid date in the format '" & dateFormat & "' and year range " & theBase.getValidYearRange() & ".")
                Return False
            End If
        Catch ex As Exception
            'theBase.out.writeLog("    ERROR in line " & theBase.lineNr & ": " & ex.Message)
            Return False
        End Try
    End Function

    Private Function getDouble(ByVal pos As Integer, ByRef theDouble As Double) As Boolean
        ' checked: v2.1
        Try
            theDouble = CDbl(splitLine(pos))
        Catch ex As Exception
            'theBase.out.writeLog("    ERROR in line " & theBase.lineNr & ": " & ex.Message)
            Return False
        End Try
        Return True
    End Function

    Private Function getInteger(ByVal pos As Integer, ByRef theInteger As Integer) As Boolean
        ' checked: v2.1
        Try
            theInteger = CInt(splitLine(pos))
        Catch ex As Exception
            'theBase.out.writeLog("    ERROR in line " & theBase.lineNr & ": " & ex.Message)
            Return False
        End Try
        Return True
    End Function

    Friend Function parseLine(ByVal currentLine As String) As rcParseResult
        ' checked: v2.0
        Dim dummyString As String
        Dim dummyDouble As Double
        Dim dummyDouble2 As Double
        Dim dummyDate As Date
        Dim dummyInt As Integer
        Dim dummyBool As Boolean
        Dim errorFlag As Boolean

        splitLine = Split(currentLine, TAB)
        Select Case splitLine(0)
            Case "@"
                ' parameters
                If (splitLine(1) = "@") And (splitLine.Count > 4) Then
                    If parseMode = rcParseModes.rcHeader Then
                        theBase.out.writeLog(currentLine)
                        Select Case splitLine(2)
                            Case "CHECK"
                                Select Case splitLine(3)
                                    Case "PROGRAM_ID"
                                        If splitLine(4) <> progID Then
                                            theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": CHECK --- this script is for another program. Requested ID: " & splitLine(4) & ". " & progName & " has the ID " & progID & ".")
                                            Return rcParseResult.rcFatalError
                                        End If

                                    Case "MIN_VERSION"
                                        errorFlag = True
                                        If getDouble(4, dummyDouble) Then
                                            If dummyDouble <= progVersion_val Then
                                                errorFlag = False
                                            End If
                                        End If
                                        If errorFlag Then
                                            theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": CHECK --- this script needs a newer version of " & progName & " (v" & progVersion_val & "), at least version " & splitLine(4) & ".")
                                            Return rcParseResult.rcFatalError
                                        End If

                                    Case Else
                                        ' unknown parameter
                                        theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": CHECK is used with an unknown parameter.")
                                        Return rcParseResult.rcFatalError
                                End Select

                            Case "SET"
                                Select Case splitLine(3)
                                    Case "SAFETY_MIN"
                                        errorFlag = True
                                        If getDouble(4, dummyDouble) Then
                                            If (dummyDouble > minMin) And (dummyDouble < maxMax) Then
                                                theBase.safetyMin = dummyDouble
                                                errorFlag = False
                                            End If
                                        End If
                                        If errorFlag Then
                                            theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": SET SAFETY_MIN cannot be set to " & splitLine(4) & ". The value needs to be in the range of " & minMin & " to " & maxMax & " (not including).")
                                            Return rcParseResult.rcFatalError
                                        End If

                                    Case "SAFETY_MAX"
                                        errorFlag = True
                                        If getDouble(4, dummyDouble) Then
                                            If (dummyDouble > minMin) And (dummyDouble < maxMax) Then
                                                theBase.safetyMax = dummyDouble
                                                errorFlag = False
                                            End If
                                        End If
                                        If errorFlag Then
                                            theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": SET SAFETY_MAX cannot be set to " & splitLine(4) & ". The value needs to be in the range of " & minMin & " to " & maxMax & " (not including).")
                                            Return rcParseResult.rcFatalError
                                        End If

                                    Case "MAX_TIME_INTERVAL"
                                        errorFlag = True
                                        If getInteger(4, dummyInt) Then
                                            If dummyInt >= minTimeLimit Then
                                                theBase.timeLimit = dummyInt
                                                errorFlag = False
                                            End If
                                        End If
                                        If errorFlag Then
                                            theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": SET MAX_TIME_INTERVAL cannot be set to " & splitLine(4) & ". The value needs minimally to be " & minTimeLimit & ".")
                                            Return rcParseResult.rcFatalError
                                        End If

                                    Case "DELTA_YEARS"
                                        errorFlag = True
                                        If getInteger(4, dummyInt) Then
                                            If theBase.setDeltaYears(dummyInt) Then
                                                errorFlag = False
                                            End If
                                        End If
                                        If errorFlag Then
                                            theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": SET DELTA_YEARS cannot be set to " & splitLine(4) & ". The value needs to be in the range of 0 to " & (theBase.maxYear - dateMinYearForOADate) & " (note: year must be >= 1900 for analysis).")
                                            Return rcParseResult.rcFatalError
                                        End If

                                    Case Else
                                        ' unknown parameter
                                        theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": SET is used with an unknown parameter.")
                                        Return rcParseResult.rcFatalError
                                End Select

                            Case Else
                                ' unknown command
                                theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": unknown command in '" & currentLine & "'.")
                                Return rcParseResult.rcFatalError
                        End Select
                    Else
                        theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": Unexpected header tag @ in '" & currentLine & "'.")
                        Return rcParseResult.rcFatalError
                    End If
                Else
                    theBase.out.writeLog("@ HEADER ERROR in line " & theBase.lineNr & ": New header definition expected. Wrong format: second @ or data missing in '" & currentLine & "'.")
                    Return rcParseResult.rcFatalError
                End If

            Case "#"
                ' new dataset
                errorFlag = True
                If (splitLine(1) = "#") And (splitLine.Count > 4) Then
                    ' get data
                    dummyString = splitLine(2)
                    If getDouble(3, dummyDouble) And getDouble(4, dummyDouble2) Then
                        errorFlag = False

                        If (parseMode = rcParseModes.rcHeader) Or (parseMode = rcParseModes.rcSkipDataset_fromHeader) Then
                            ' first time
                            theBase.logParameters()
                            If theBase.safetyMin < theBase.safetyMax Then
                                currentDataset = New dataSet(theBase, dummyString, dummyDouble, dummyDouble2)
                                currentDataset.printLabels()
                                helper_addErrorPlaceholders()
                                parseMode = rcParseModes.rcDatasets
                            Else
                                theBase.out.writeLog("@ HEADER ERROR: SAFETY_MIN is not lower than SAFETY_MAX.")
                                Return rcParseResult.rcFatalError
                            End If
                        Else
                            ' analyze first
                            helper_flushBuffer()
                            ' create a new one
                            currentDataset = New dataSet(theBase, dummyString, dummyDouble, dummyDouble2)
                            parseMode = rcParseModes.rcDatasets
                        End If
                    End If
                End If

                If errorFlag Then
                    theBase.out.writeLog("# DATA ERROR in line " & theBase.lineNr & ": New dataset expected. Wrong format; second # or data missing/wrong in '" & currentLine & "'.")
                    theBase.out.writeLog("  The next lines will be skipped up to the next new dataset definition." & CR & LF)
                    If parseMode = rcParseModes.rcHeader Then
                        parseMode = rcParseModes.rcSkipDataset_fromHeader
                    ElseIf parseMode = rcParseModes.rcDatasets Then
                        parseMode = rcParseModes.rcSkipDataset_fromDatasets
                    End If
                    addErrorDatasetCounter += 1
                    Return rcParseResult.rcError
                End If

            Case Else
                Select Case parseMode
                    Case rcParseModes.rcDatasets
                        ' parse data
                        If getDate(0, dummyDate) And getDouble(1, dummyDouble) Then
                            dummyBool = getSimpleBoolean(2, "B")
                        Else
                            theBase.out.writeLog("* DATA ERROR in line " & theBase.lineNr & ": Wrong format (expected: DATE  VALUE  [optional: B for bridging]). Check data in '" & currentLine & "'." & CR & LF)
                            currentDataset.setErrorFlag()
                            Return rcParseResult.rcError
                        End If

                        If Not currentDataset.getErrorFlag Then
                            currentDataset.addDataItem(dummyDate, dummyDouble, dummyBool)
                        End If

                    Case rcParseModes.rcHeader
                        ' assume comment or title line for the user
                        theBase.out.writeLog("- Line " & theBase.lineNr & ": comment or column header '" & currentLine & "'.")

                    Case rcParseModes.rcSkipDataset_fromDatasets, rcParseModes.rcSkipDataset_fromHeader
                        ' nothing

                    Case Else
                        ' default
                        theBase.out.writeLog("* DATA ERROR in line " & theBase.lineNr & ": Wrong format. Check data and location in '" & currentLine & "'." & CR & LF)
                        Return rcParseResult.rcFatalError
                End Select
        End Select

        ' default
        Return rcParseResult.rcContinue
    End Function
End Class

Friend Class rcBase
    ' checked: v2.2
    Friend theForm As mainForm = Nothing
    Friend safetyMin As Double = stdSafetyMin
    Friend safetyMax As Double = stdSafetyMax
    Friend timeLimit As Integer = stdTimeLimit
    Private deltaYears As Integer = dateStdDeltaYears
    Friend minYear As Integer = 0
    Friend maxYear As Integer = 0
    Friend lineNr As Long = 0

    Friend upnList_Str As List(Of String) = New List(Of String)
    Friend upnList_Int As List(Of Long) = New List(Of Long)
    Friend upnList_allNumbers As Boolean = True

    Friend out As outputClass = New outputClass(Me)

    Friend Sub New(ByVal theMainForm As mainForm)
        ' checked: v2.0
        theForm = theMainForm
        maxYear = Today.Year
        minYear = maxYear - deltaYears
        out.initialize()
    End Sub

    Friend Sub cleanup()
        ' checked: v2.0
        out.cleanup()
    End Sub

    Friend Function setDeltaYears(ByVal newDelta As Integer) As Boolean
        ' checked: v2.0
        If (newDelta >= 0) And ((maxYear - newDelta) >= dateMinYearForOADate) Then
            deltaYears = newDelta
            minYear = maxYear - deltaYears
            Return True
        Else
            Return False
        End If
    End Function

    Friend Function getValidYearRange() As String
        ' checked: v2.0
        Return minYear & " to " & maxYear
    End Function

    Friend Function analyzeFile(ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs, ByVal fileNames As List(Of String)) As workerTypes
        ' checked: v2.1
        Dim inputParser As inputRCparser = Nothing
        Dim fileStream As StreamReader = Nothing
        Dim fileLen As Long = 0
        Dim currentLine As String = ""
        Dim result As rcParseResult = rcParseResult.rcNothing

        fileNames.Sort()
        For Each fileName As String In fileNames
            If result <> rcParseResult.rcFatalError Then
                out.writeLog("*** Analysis of '" & fileName & "'", outputClass.logSelector.mainLog)
                theForm.updateFormItem(theForm.fileTxt, getShortFilename(fileName))
                worker.ReportProgress(0)

                Try
                    If out.openResultAndLogStreams(getShortFilenameWithoutSuffix(fileName)) Then
                        fileStream = New StreamReader(fileName)
                        fileLen = fileStream.BaseStream.Length
                        inputParser = New inputRCparser(Me)
                        lineNr = 0
                        Do
                            currentLine = fileStream.ReadLine
                            lineNr += 1
                            result = inputParser.parseLine(currentLine)

                            '=== report progress ===============================================================
                            If lineNr Mod 100 = 0 Then
                                worker.ReportProgress(CInt(100.0 * fileStream.BaseStream.Position / fileLen))

                                '--- handle cancel request -----------------------------------------------
                                If worker.CancellationPending Then
                                    e.Cancel = True
                                    Return workerTypes.WORKER_Cancelled
                                End If
                                '-------------------------------------------------------------------------
                            End If
                        Loop Until (fileStream.EndOfStream) Or (result = rcParseResult.rcFatalError)

                        If result <> rcParseResult.rcFatalError Then
                            inputParser.flushBuffer()
                        End If
                        worker.ReportProgress(100)
                        inputParser.logCounter()
                    Else
                        out.writeLog("*** ERROR in rcBase.analyzeFile(): Could not open the output result and log files.", outputClass.logSelector.mainLog)
                        result = rcParseResult.rcFatalError
                    End If

                Catch ex As Exception
                    out.writeLog("    ERROR: " & ex.Message)
                    result = rcParseResult.rcFatalError
                Finally
                    fileStream.Close()
                End Try

                out.writeLogAndStatus(theForm, lineNr & " lines analyzed.")
                out.writeLog()
            Else
                out.writeLog("", outputClass.logSelector.mainLog)
                out.writeLog("*** Skipped '" & fileName & "' because of prior fatal error.", outputClass.logSelector.mainLog)
            End If
        Next fileName

        ' log sorted list of analyzed UPNs
        out.writeLog("", outputClass.logSelector.mainLog)
        out.writeLog(upnList_Str.Count & " UPNs analyzed:", outputClass.logSelector.mainLog)
        If upnList_allNumbers Then
            upnList_Int.Sort()
            For Each knownUPN As Long In upnList_Int
                out.writeLog(knownUPN, outputClass.logSelector.mainLog)
            Next
        Else
            upnList_Str.Sort()
            For Each knownUPN As String In upnList_Str
                out.writeLog(knownUPN, outputClass.logSelector.mainLog)
            Next
        End If
        out.writeLog("", outputClass.logSelector.mainLog)

        If result = rcParseResult.rcFatalError Then
            out.writeLog("    ERROR: Could not analyze the file(s) (see previous error message(s) for details).")
            theForm.updateStatus("ERROR while loading/analyzing the file(s).", False)
            Return workerTypes.ANALYSIS_resultERROR
        Else
            Return workerTypes.ANALYSIS_resultOK
        End If
    End Function

    Friend Sub logParameters()
        ' checked: v2.0
        out.writeLog()
        out.writeLog("Init: used parameters for analysis:")
        out.writeLog("- safety min:  " & safetyMin & BooleanToString(safetyMin = stdSafetyMin, " [default]", ""))
        out.writeLog("- safety max:  " & safetyMax & BooleanToString(safetyMax = stdSafetyMax, " [default]", ""))
        out.writeLog("- time limit:  " & timeLimit & BooleanToString(timeLimit = stdTimeLimit, " [default]", ""))
        out.writeLog("- valid years: " & getValidYearRange() & " (defined delta: " & deltaYears & BooleanToString(deltaYears = dateStdDeltaYears, " [default]", "") & ")")
        out.writeLog()
        out.writeLog("Data")
    End Sub
End Class
