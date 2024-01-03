Sub CrearArraySimple()
  Dim columnas As Variant ' Declara un array
  Dim iteracion As Variant
  Dim wS, wS1 As Worksheet, Lastrow, Lastrow1 As Long
  Lastrow = wS.UsedRange.Row - 1 + wS.UsedRange.Rows.Count

  ' Crea un array con valores del 1 al 5
  columnas = Array(1, 2, 3, 4, 5)

  For I = Lastrow To 1 Step -1
    If Not (IsEmpty(wS.Cells(I, 1))) Then Exit For
    Next I

    Lastrow = I


    ' Accede a los valores del array y muestra en la ventana de resultados
    For J = LBound(columnas) To UBound(columnas)
      For I = 2 To Lastrow
        iteracion = Sheets(1).Cells(I, columnas(J)).Value
        Debug.Print iteracion

      Next J
    Next i
End Sub

