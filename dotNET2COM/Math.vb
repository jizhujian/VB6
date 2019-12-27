''' <summary>
''' ��ѧ������
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(Math.ClassId, Math.InterfaceId, Math.EventsId)> _
Public Class Math

  ''' <summary>
  ''' COMע�����
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "078c82a5-eb10-4f8c-b299-b69fb6fe7256"
  Public Const InterfaceId As String = "88c8a0d0-1b91-45f6-a96e-9b591de6f469"
  Public Const EventsId As String = "2fbca757-bfb6-4758-a564-b454ca8e3395"

  Public Sub New()
    MyBase.New()
  End Sub

  ''' <summary>
  ''' ʮ����ת�������
  ''' </summary>
  ''' <param name="decValue"></param>
  ''' <param name="nBase"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Dec2NBase(ByVal decValue As Integer, ByVal nBase As Byte) As String

    Dim numericBaseData As String
    Dim nBaseValue As String
    Dim remainder As Integer

    numericBaseData = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".Substring(0, nBase)
    nBaseValue = ""
    Do While (decValue > 0)
      decValue = System.Math.DivRem(decValue, nBase, remainder)
      nBaseValue = numericBaseData.Substring(remainder, 1) & nBaseValue
    Loop
    If (nBaseValue = "") Then
      nBaseValue = "0"
    End If
    Return nBaseValue

  End Function

  ''' <summary>
  ''' �������תʮ����
  ''' </summary>
  ''' <param name="nBaseValue"></param>
  ''' <param name="nBase"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function NBase2Dec(ByVal nBaseValue As String, ByVal nBase As Byte) As Integer

    Dim numericBaseData As String
    Dim decValue As Integer
    Dim power As Integer
    Dim pos As Integer

    numericBaseData = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".Substring(0, nBase)
    nBaseValue = nBaseValue.ToUpper
    pos = nBaseValue.Length - 1
    power = 1
    Do While (pos >= 0)
      decValue += power * numericBaseData.IndexOf(nBaseValue.Substring(pos, 1))
      power *= nBase
      pos -= 1
    Loop
    Return decValue

  End Function

  ''' <summary>
  ''' ������ֵ���ʽֵ
  ''' </summary>
  ''' <param name="expression"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Eval(ByVal expression As String) As Double

    Dim vbCodeProvider As Microsoft.VisualBasic.VBCodeProvider
    Dim compilerResults As System.CodeDom.Compiler.CompilerResults
    Dim assembly As System.Reflection.Assembly
    Dim type As System.Type
    Dim method As System.Reflection.MethodInfo
    Dim compilerParameters As System.CodeDom.Compiler.CompilerParameters
    Dim assemblyStringBuilder As System.Text.StringBuilder
    Dim errorStringBuilder As System.Text.StringBuilder

    compilerParameters = New System.CodeDom.Compiler.CompilerParameters
    With compilerParameters
      .GenerateExecutable = False
      .GenerateInMemory = True
      .IncludeDebugInformation = False
      .ReferencedAssemblies.Add("mscorlib.dll") ' System.Math
    End With

    assemblyStringBuilder = New System.Text.StringBuilder
    With assemblyStringBuilder
      .AppendLine("Imports Microsoft.VisualBasic")
      .AppendLine("Imports System.Math")
      .AppendLine("Namespace VB_Eval_Namespace")
      .AppendLine("  Public Class VB_Eval_Class")
      .AppendLine("    Public Function VB_Eval_Method() As Double")
      .Append("      Return ").AppendLine(expression)
      .AppendLine("    End Function")
      .AppendLine("  End Class")
      .AppendLine("End Namespace")
    End With

    vbCodeProvider = New Microsoft.VisualBasic.VBCodeProvider
    compilerResults = vbCodeProvider.CompileAssemblyFromSource(compilerParameters, assemblyStringBuilder.ToString)
    If compilerResults.Errors.HasErrors Then
      errorStringBuilder = New System.Text.StringBuilder
      With compilerResults.Errors
        For i As Integer = 0 To .Count - 1
          errorStringBuilder.AppendLine(.Item(i).ToString)
        Next
      End With
      'errorStringBuilder.AppendLine().AppendLine(assemblyStringBuilder.ToString)
      Throw New System.Exception(errorStringBuilder.ToString)
    End If

    assembly = compilerResults.CompiledAssembly
    type = assembly.GetType("VB_Eval_Namespace.VB_Eval_Class", True)
    method = type.GetMethod("VB_Eval_Method")
    Eval = CDbl(method.Invoke(System.Activator.CreateInstance(type), New Object() {}))

  End Function

  ''' <summary>
  ''' ����С������������
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Truncate(ByVal value As Double) As Integer
    Return CInt(System.Math.Truncate(value))
  End Function

  ''' <summary>
  ''' ��һ�������������������ֵ��м�ֵʱ���Ὣ������Ϊ����ֵ�о���ֵ��С��ֵ
  ''' </summary>
  ''' <param name="value"></param>
  ''' <param name="decimals"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Round(ByVal value As Double, ByVal decimals As Integer) As Double
    Return System.Math.Round(value, decimals, System.MidpointRounding.AwayFromZero)
  End Function

  ''' <summary>
  ''' ����С�ڵ���ֵ�������
  ''' </summary>
  ''' <param name="value"></param>
  ''' <param name="decimals"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Floor(ByVal value As Double, Optional ByVal decimals As Integer = 0) As Double
    If decimals = 0 Then
      Return System.Math.Floor(value)
    Else
      Dim zoom As Double = CInt(System.Math.Pow(10, decimals))
      Return System.Math.Floor(value * zoom) / zoom
    End If
  End Function

  ''' <summary>
  ''' ���ش��ڵ���ֵ����С��
  ''' </summary>
  ''' <param name="value"></param>
  ''' <param name="decimals"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Ceiling(ByVal value As Double, Optional ByVal decimals As Integer = 0) As Double
    If decimals = 0 Then
      Return System.Math.Ceiling(value)
    Else
      Dim zoom As Double = CInt(System.Math.Pow(10, decimals))
      Return System.Math.Ceiling(value * zoom) / zoom
    End If
  End Function

End Class
