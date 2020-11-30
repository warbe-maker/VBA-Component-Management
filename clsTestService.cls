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

Public Property Let TestedByTheWay(ByVal btw As String):    sTestedByTheWay = btw:              End Property
Public Property Get Details() As String:                    Details = sDetails:                 End Property

Public Property Let Details(ByVal s As String):             sDetails = s:                       End Property

Public Property Get FAILED() As String:                     FAILED = "failed!":                 End Property

Public Property Get PASSED() As String:                     PASSED = "passed!":                 End Property

Public Property Get Regression() As Boolean:                Regression = bRegression:           End Property

Public Property Let Regression(ByVal b As Boolean):         bRegression = b:                    End Property

Public Property Get ResultExpected() As Variant:            ResultExpected = vResultExpected:   End Property

Public Property Let ResultExpected(ByVal v As Variant):     vResultExpected = v:                End Property

Public Property Get TestItem() As String:                   TestItem = sTestItem:               End Property

Public Property Let TestItem(ByVal s As String):            sTestItem = s:                      End Property

Public Property Get TestProcedure() As String:              TestProcedure = sTestProcedure:     End Property

Public Property Let TestProcedure(ByVal s As String):       sTestProcedure = s:                 End Property

Public Property Get TestResult() As String:                 TestResult = sTestResult:           End Property

Public Property Let TestResult(ByVal s As String):          sTestResult = s:                    End Property

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
