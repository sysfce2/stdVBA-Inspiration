VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdSentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


'In numerous cases using excel we like to set standard options to optimise performance:
'
'    Dim bEnableEvents as Boolean: bEnableEvents = Application.EnableEvents
'    Dim bScreenUpdating as Boolean: bScreenUpdating = Application.ScreenUpdating
'    Application.EnableEvents = false
'    Application.ScreenUpdating = false
'    
'      ... our code ...
'    
'    Application.EnableEvents = bEnableEvents
'    Application.ScreenUpdating = bScreenUpdating
'
'Similarly with std libraries, we have to add and pop from the stack:
'
'    stdError.addStack "myMethod"
'      ... our code ...
'    stdError.popStack
'
'stdPerformance today uses the sentry pattern:
'
'    With stdPerformance.CreateMeasure("Some measurement")
'      ... our code ...
'    End With
'
'Although the two patterns have the same size (3 lines), the sentry pattern is both more beautiful and more robust.
'This class is an attempt to make the sentry pattern easier to implement in VBA.
'
'   Dim ensureScreenUpdating as stdSentry
'   set ensureScreenUpdating = stdSentry.Create( _ 
'     stdLambda.CreateMultiLine( _ 
'       Array( _ 
'         "let temp = Application.ScreenUpdating", _
'         "let Application.ScreenUpdating = $1", _
'         "Array(temp)", _
'       ) _ 
'     ), _
'     stdLambda.Create("let Application.ScreenUpdating = $1") _ 
'   )
'   With ensureScreenUpdating(False)
'     '... our code ...
'   End With
'------------------------
'   set sentry = stdSentry.Create(
'     stdCallback.CreateFromModule("", "RNOn"), _
'     stdCallback.CreateFromModule("", "RNOff") _
'   )

Implements stdICallable
Private Type TFactory
  OnInit As stdICallable
  OnDestroy As stdICallable
  passVars As Boolean
End Type
Private Type TInstance
  onDestroy As stdICallable
  DestroyArgs As Variant
  passVars As Boolean
End Type
Private Type TThis
  Factory As TFactory
  Instance As TInstance
End Type
Private This As TThis

'Create a sentry factory
'@param OnInit - Function to run when the sentry is created.
'@param OnDestroy - Function to run when the sentry is destroyed.
'@param passVars - If true, the result of OnInit is passed to OnDestroy. If false, no arguments are passed to OnDestroy.
'@returns - The sentry factory
Public Function Create(ByVal OnInit As stdICallable, ByVal OnDestroy As stdICallable, Optional ByVal passVars As Boolean = True) As stdSentry
  Set Create = New stdSentry
  Call Create.protInitFactory(OnInit, OnDestroy, passVars)
End Function

'Create a sentry which acts as an "optimiser" for code. Note, setting options doesn't always lead to performance improvements. See [vba-articles performance thread](https://sancarn.github.io/vba-articles/performance-tips.html) for details.
'@param EnableEvents   - If not Empty, sets Application.EnableEvents to this value. If Empty, Application.EnableEvents remains unchanged.
'@param ScreenUpdating - If not Empty, sets Application.ScreenUpdating to this value. If Empty, Application.ScreenUpdating remains unchanged.
'@param Calculation    - If not -1, sets Application.Calculation to this value. If -1, Application.Calculation remains unchanged.
'@returns - A sentry which will restore the original settings when destroyed.
'@requires stdLambda
'@example ```vba
'With stdSentry.CreateOptimiser(ScreenUpdating:=False, Calculation:=xlCalculationManual)
'  '... heavy range based operations ...
'End With
'```
Public Function CreateOptimiser(Optional ByVal EnableEvents As VbTriState = vbUseDefault, Optional ByVal ScreenUpdating As VbTriState = vbUseDefault, Optional ByVal Calculation As Long = -1) As stdSentry
  If Not IsObject(stdLambda) Then Err.Raise 5, "stdSentry.CreateFromObjectMethod", "stdLambda is required for this function"
  
  'Validate inputs
  If Not (Calculation = -1 Or Calculation = 2 Or Calculation = -4135 Or Calculation = -4105) Then Err.Raise 5, "stdSentry.CreateOptimiser", "Calculation must be either xlCalculationAutomatic, xlCalculationSemiautomatic, xlCalculationManual, or -1"
  
  'Convert Calculation to a variant
  Dim vCalculation As Variant: vCalculation = IIf(Calculation = -1, Empty, Calculation)

  'Create the OnInit and OnDestroy functions
  '@remark vbUseDefault == -2
  Dim onInit As stdICallable: Set onInit = stdLambda.CreateMultiline(Array( _
    "let t1 = -2", _ 
    "let t2 = -2", _ 
    "let t3 = -2", _
    "If $1 <> -2 Then let t1 = Application.EnableEvents   : let Application.EnableEvents = $1    end", _
    "If $2 <> -2 Then let t2 = Application.ScreenUpdating : let Application.ScreenUpdating = $2  end", _
    "If $3 <> -2 Then let t3 = Application.Calculation    : let Application.Calculation = $3     end", _
    "Array(t1,t2,t3)" _
  ))
  Dim onDestroy As stdICallable: Set onDestroy = stdLambda.CreateMultiline(Array( _
    "If $1 <> -2 Then let Application.EnableEvents = $1 end", _
    "If $2 <> -2 Then let Application.ScreenUpdating = $2 end", _
    "If $3 <> -2 Then let Application.Calculation = $3 end" _
  ))
  
  'Create the sentry
  Set CreateOptimiser = Create(onInit, onDestroy).Run(EnableEvents, ScreenUpdating, vCalculation)
End Function

'Create a sentry from a property of an object
'@param obj - Object to bind the property to
'@param sPropertyName - Name of the property to bind
'@param value - Value to set the property to
'@returns - The sentry
'@requires stdLambda
'@example ```vba
'With stdSentry.CreateFromObjectProperty(Application, "DisplayAlerts", false)
'  Call ThisWorkbook.SaveAs("C:\temp\test.xlsx")
'End With
'```
Public Function CreateFromObjectProperty(ByVal obj As Object, ByVal sPropertyName As String, ByVal value As Variant) As stdSentry
  If Not IsObject(stdLambda) Then Err.Raise 5, "stdSentry.CreateFromObjectMethod", "stdLambda is required for this function"
  Dim onInit As stdICallable: Set onInit = stdLambda.CreateMultiline(Array( _
    "let t1 = $1." & sPropertyName, _ 
    "let $1." & sPropertyName & " = $2", _
    "Array(t1)" _
  )).Bind(obj)
  Dim onDestroy As stdICallable: Set onDestroy = stdLambda.CreateMultiline(Array( _
    "let $1." & sPropertyName & " = $2" _
  )).Bind(obj)

  Set CreateFromObjectProperty = Create(onInit, onDestroy).Run(value)
End Function

'Create a sentry from a method of an object
'@param obj - Object to bind the method to
'@param sOnCreateName - Name of the method to call when the sentry is created
'@param sOnDestroyName - Name of the method to call when the sentry is destroyed
'@returns - The sentry
'@requires stdCallback
'@example ```vba
'With stdSentry.CreateFromObjectMethod(stdError, "AddStack", Array("MethodName"), "PopStack", Array())
'  '... our code ...
'End With
'```
Public Function CreateFromObjectMethod(ByVal obj As Object, ByVal OnInitName As String, ByVal initParams As Variant, ByVal OnDestroyName As String, ByVal destroyParams As Variant) As stdSentry
  If Not IsObject(stdCallback) Then Err.Raise 5, "stdSentry.CreateFromObjectMethod", "stdCallback is required for this function"
  Dim onInit As stdICallable: Set onInit = stdCallback.CreateFromObjectMethod(obj, OnInitName).BindEx(initParams)
  Dim onDestroy As stdICallable: Set onDestroy = stdCallback.CreateFromObjectMethod(obj, OnDestroyName).BindEx(destroyParams)
  Set CreateFromObjectMethod = Create(onInit, onDestroy, False).Run()
End Function

'Create a sentry for error stack management
'@param name - Name of the stack to add
'@returns - The sentry
'@requires stdError, stdCallback
'@example ```vba
'Public Sub MyMethod()
'  With stdSentry.CreateErrorStack("MyClass.MyMethod")
'    '... our code ...
'  End With
'End Sub
'```
Public Function CreateErrorStack(ByVal name As String) As stdSentry
  If Not IsObject(stdError) Then Err.Raise 5, "stdSentry.CreateFromObjectMethod", "stdError is required for this function"
  Set CreateErrorStack = CreateFromObjectMethod(stdError, "AddStack", Array(name), "PopStack", Array()).Run()
End Function

'Create a sentry from an AutomationSecurity setting
'@param sec - AutomationSecurity setting desired
'@returns - The sentry
'@requires stdLambda
'@example ```vba
'With stdSentry.CreateFromAutomationSecurity(msoAutomationSecurityForceDisable)
'  Call ThisWorkbook.SaveAs("C:\temp\test.xlsx")
'End With
'```
Public Function CreateFromAutomationSecurity(ByVal sec As MsoAutomationSecurity) As stdSentry
  Set CreateFromAutomationSecurity = CreateFromObjectProperty(Application, "AutomationSecurity", sec)
End Function


'@param OnInit - Function to run when the sentry is created. Should return arguments to pass to OnDestroy. If OnInit returns Empty, no arguments are passed to OnDestroy.
'@param OnDestroy - Function to run when the sentry is destroyed. Should take the arguments returned by OnInit.
Public Sub protInitFactory(ByVal OnInit As stdICallable, ByVal OnDestroy As stdICallable, ByVal passVars As Boolean)
  With This.Factory
    Set .OnInit = OnInit
    Set .OnDestroy = OnDestroy
    .passVars = passVars
  End With
End Sub
Public Sub protInitInstance(ByVal onDestroy As stdICallable, ByVal args As Variant, ByVal passVars As Boolean)
  With This.Instance
    Set .onDestroy = onDestroy
    .DestroyArgs = args
    .passVars = passVars
  End With
End Sub

'Run the sentry
'@param args - Arguments to pass to the OnInit function
'@returns - The sentry
Public Function Run(ParamArray args() As Variant) As stdSentry
Attribute Run.VB_UserMemId = 0
  Dim copy: copy = args
  Set Run = RunEx(copy)
End Function

'Run the sentry
'@param args - Arguments to pass to the OnInit function
'@returns - The sentry
Public Function RunEx(ByVal args As Variant) As stdSentry
  'Run the OnInit function and prepare destroy arguments
  Dim destroyArgs: destroyArgs = This.Factory.OnInit.RunEx(args)
  If IsEmpty(destroyArgs) Then destroyArgs = Array()

  'Return the sentry
  Set RunEx = New stdSentry
  Call RunEx.protInitInstance(This.Factory.OnDestroy, destroyArgs, This.Factory.passVars)
End Function

'Bind arguments to the sentry factory
'@param args - Arguments to bind to the sentry
'@returns - The sentry
Public Function Bind(ParamArray args() As Variant) As stdSentry
  Dim copy: copy = args
  Set Bind = BindEx(copy)
End Function

'Bind arguments to the sentry factory
'@param args as Variant<Array<Variant>> - Arguments to bind to the sentry
'@returns - The sentry
Public Function BindEx(ByVal args As Variant) As stdSentry
  Set BindEx = New stdSentry
  'TODO: 
End Function

'Making late-bound calls to `stdSentry` members. Each object which implements `stdICallable`
'will support a different set of latebound calls.
'@protected
'@param sMessage - Message to send. Standard messages include "obj" returning the object, "className" returning the class name. Other messages are implementation specific.
'@param success  - Whether the call was successful
'@param params   - Any variant, typically parameters as an array. Passed along with the message.
'@returns - Any return value.
Public Function SendMessage(ByVal sMessage As String, ByRef success As Boolean, ByVal params As Variant) As Variant
  Select Case sMessage
    Case "obj":
      success = True
      Set SendMessage = Me
    Case "className":
      success = True
      SendMessage = "stdSentry"
    Case "OnInit":
      success = True
      Set SendMessage = This.Factory.OnInit
    Case "OnDestroy":
      success = True
      Set SendMessage = This.Factory.OnDestroy
    Case Else
      success = False
  End Select
End Function


Private Sub Class_Terminate()
  With This.Instance
    If Not .onDestroy Is Nothing Then
      If This.Instance.passVars Then
        Call .onDestroy.RunEx(.DestroyArgs)
      Else
        Call .onDestroy.Run
      End If
    End If
  End With
End Sub

'Copies a variant
'@param vOut - Variant to copy to
'@param vIn - Variant to copy from
Private Sub CopyVariant(ByRef vOut, ByVal vIn)
  If IsObject(vIn) Then
    Set vOut = vIn
  Else
    vOut = vIn
  End If
End Sub

'Implement stdICallable
Private Function stdICallable_Bind(ParamArray params() As Variant) As stdICallable
  Dim v: v = params
  Call CopyVariant(stdICallable_Bind, BindEx(v))
End Function
Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
  Dim v: v = params
  Call CopyVariant(stdICallable_Run, RunEx(v))
End Function
Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
  Call CopyVariant(stdICallable_RunEx, RunEx(params))
End Function
Private Function stdICallable_SendMessage(ByVal sMessage As String, success As Boolean, ByVal params As Variant) As Variant
  Call CopyVariant(stdICallable_SendMessage, SendMessage(sMessage, success, params))
End Function

