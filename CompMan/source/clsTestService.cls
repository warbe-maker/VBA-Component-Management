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

Friend Property Get details() As String:                    details = sDetails:                 End Property

Friend Property Let details(ByVal s As String):             sDetails = s:                       End Property

'Friend Property Get ResultIsFailed() As String:             ResultIsFailed = "failed!":         End Property
'
'Friend Property Get ResultIsPassed() As String:             ResultIsPassed = "passed!":         End Property
'
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

