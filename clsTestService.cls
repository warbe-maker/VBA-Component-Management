VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bRegression     As Boolean
Private sDetails        As String
Private sTestResult     As String
Private sTestProcedure  As String
Private vResultExpected As Variant
Private sTestItem       As String
Private sTestedByTheWay As String

Friend Property Let TestedByTheWay(ByVal btw As String):    sTestedByTheWay = btw:              End Property
Friend Property Get Details() As String:                    Details = sDetails:                 End Property

Friend Property Let Details(ByVal s As String):             sDetails = s:                       End Property

Friend Property Get FAILED() As String:                     FAILED = "failed!":                 End Property

Friend Property Get PASSED() As String:                     PASSED = "passed!":                 End Property

Friend Property Get Regression() As Boolean:                Regression = bRegression:           End Property

Friend Property Let Regression(ByVal b As Boolean):         bRegression = b:                    End Property

Friend Property Get ResultExpected() As Variant:            ResultExpected = vResultExpected:   End Property

Friend Property Let ResultExpected(ByVal v As Variant):     vResultExpected = v:                End Property

Friend Property Get TestItem() As String:                   TestItem = sTestItem:               End Property

Friend Property Let TestItem(ByVal s As String):            sTestItem = s:                      End Property

Friend Property Get TestProcedure() As String:              TestProcedure = sTestProcedure:     End Property

Friend Property Let TestProcedure(ByVal s As String):       sTestProcedure = s:                 End Property

Friend Property Get TestResult() As String:                 TestResult = sTestResult:           End Property

Friend Property Let TestResult(ByVal s As String):          sTestResult = s:                    End Property

Public Function Evaluated(ByVal eval_result As Variant) As String
' ---------------------------------------------------------------
' Returns TRUE when the evaluation of the test result is positiv.
' ---------------------------------------------------------------
    
    If eval_result = vResultExpected Then
        sTestResult = Me.PASSED
        Evaluated = Me.PASSED
    Else
        sTestResult = Me.FAILED
        Evaluated = Me.FAILED
    End If
    If bRegression Then Me.Log

End Function

Public Sub Log()
    Debug.Print "Test procedure   : " & sTestProcedure & vbLf & _
                "Test item        : " & sTestItem
    If sTestedByTheWay <> vbNullString Then _
    Debug.Print "Tested by-the-way: " & sTestedByTheWay
    Debug.Print "Test detail      : " & sDetails & vbLf & _
                "Test result      : " & sTestResult
End Sub
