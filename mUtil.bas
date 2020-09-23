Attribute VB_Name = "mUtil"
Option Explicit


Public Sub Main()


    Dim fSrch As New fSearch
    fSrch.Init "Customer", "FirstName,LastName,DOB,Weight"
    Set fSrch = Nothing

End Sub
