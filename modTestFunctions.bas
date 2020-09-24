Attribute VB_Name = "modTestFunctions"
'------------------------------------------
'Clear the subs and insert your
'own code to speed test.
'
'Let the amount of times it repeats
'be controlled by the iterations text box
'
'Id Like to give thanks to Chris Lukas for making
'his string mapping tutorial (http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=34787&lngWId=1)
'Id been looking for ages to find both a way to performance test
'my code and speed up my strings. Nice One :)
'------------------------------------------

Public Sub Test1()
    a = Rnd()
    b = 1
    
    For i = 1 To Rnd() * 1000
        a = a + b
    Next
End Sub


Public Sub Test2()
    a = Rnd()
    b = 1
    
    For i = 1 To Rnd() * 1000
        a = a - b
    Next
End Sub


Public Sub Test3()
    a = Rnd()
    b = 1
    
    For i = 1 To Rnd() * 1000
        a = a * b
    Next
End Sub


Public Sub Test4()
    a = Rnd()
    b = 1
    
    For i = 1 To Rnd() * 1000
        a = a / b
    Next
End Sub
