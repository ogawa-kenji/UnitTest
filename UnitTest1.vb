Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class PersonTest
    Private personName As String = "Bob"
    Private personAge As Integer = 21

    <TestMethod()> Public Sub testCreatePersonWithNameAndAge()
        Dim p = New Person(personName, personAge)
        Assert.AreEqual(personName, p.getName())
        Assert.AreEqual(personAge, p.getAge())
    End Sub

End Class

Friend Class Person
    Private personName As String
    Private personAge As Integer

    Public Sub New(personName As String, personAge As Integer)
        Me.personName = personName
        Me.personAge = personAge
    End Sub

    Public ReadOnly Property getName() As String
        Get
            Return personName
        End Get
    End Property

    Public ReadOnly Property getAge() As Integer
        Get
            Return personAge
        End Get
    End Property


End Class
