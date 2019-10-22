Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UnitTestProject1

<TestClass()> Public Class UnitTest2


    <TestMethod()> Public Sub 改行コードの有無チェックができる()
        Dim textFileMock As New TextFileMock()
        textFileMock.Value = "AAA" & Environment.NewLine
        Dim textFile = New TextFile(textFileMock)
        Assert.AreEqual(True, textFile.NewLineExists())
    End Sub

    <TestMethod()> Public Sub 改行コードの有無チェックができるFalse()
        Dim textFileMock As New TextFileMock()
        textFileMock.Value = "BBB"
        Dim textFile = New TextFile(textFileMock)
        Assert.AreEqual(False, textFile.NewLineExists())
    End Sub


End Class

Public Class TextFile
    Public Sub New()
    End Sub

    Private _textFile As ITextFile
    Public Sub New(textFile As ITextFile)
        _textFile = textFile
    End Sub


    Public Function NewLineExists() As Boolean
        Dim fileString = _textFile.GetData()
        Return fileString.Contains(System.Environment.NewLine)
    End Function


End Class

Public Interface ITextFile
    Function GetData() As String
End Interface

Public Class TextFileAccess
    Implements ITextFile

    Public Function GetData() As String Implements ITextFile.GetData
        Return System.IO.File.ReadAllText("C:\work\Test.txt")
    End Function
End Class

Friend Class TextFileMock
    Implements ITextFile

    Friend Property Value As String
    Public Function GetData() As String Implements ITextFile.GetData
        Return Value
    End Function
End Class