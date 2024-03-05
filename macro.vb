Sub COMPILE()

    '%%%% LASTROW %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    Dim wS, wS1 As Worksheet, Lastrow, Lastrow1 As Long
    Dim columnas, datos_lectura, datos_escritura, repeticiones, almacenador, colores, letrasArray, filos_lectura, filos_escritura, busqueda, numericos_lectura, numericos_escritura As Variant ' Declara un array
    Dim I, J As Long
    Dim duro, suave, tipo_filo As String
    Set wS = Sheets(1)

    duro = "22X2"
    suave = "22X045"
    almacenador = 0
    colores = Sheets("Limpieza").Range("A:B").Value

    columnas = Array(7, 9, 10, 12, 13, 15, 16, 18, 19)
    numericos_validacion = Array(4, 5, 6)
    datos_lectura = Array(4, 5, 6, 20, 21)
    datos_escritura = Array(3, 4, 5, 11, 12)
    filos_lectura = Array(8, 11, 14, 17)
    filos_escritura = Array(7, 8, 9, 10)

    ReDim letrasArray(1 To 21) ' Dimensiona el array
    
    For I = 1 To 21
        letrasArray(I) = Chr(64 + I)
    Next I

    Lastrow = wS.UsedRange.Row - 1 + wS.UsedRange.Rows.Count


    For I = Lastrow To 1 Step -1
        If Not (IsEmpty(wS.Cells(I, 1))) Then Exit For
        Next I
        Lastrow = I
        For J = LBound(columnas) To UBound(columnas)
            For I = 2 To Lastrow
                Dim validador As Variant
                validador = Len(Trim(Sheets(1).Cells(I, columnas(J))))
                If validador > 1 Then
                    almacenador = validador
                End If

                If almacenador > 0 Then
                    almacenador = almacenador - 1
                    Sheets(1).Cells(I, columnas(J)).Value = "X"
                End If
            Next I
        Next J

        For J = LBound(numericos_validacion) To UBound(numericos_validacion)
            For I = 2 To Lastrow
                If Not IsNumeric(Sheets(1).Cells(I, datos_lectura(J)).Value) Then
                    MsgBox ("Error, valor no numérico en: " & letrasArray(numericos_validacion(J)) & I)
                    Sheets(1).Cells(I, numericos_validacion(J)).Select
                    Exit Sub
                End If
            Next I
        Next J
        For J = LBound(datos_lectura) To UBound(datos_lectura)
            For I = 2 To Lastrow
                Sheets(3).Cells(I + 6, datos_escritura(J)).Value = Sheets(1).Cells(I, datos_lectura(J)).Value
            Next I
        Next J

        For J = LBound(filos_lectura) To UBound(filos_lectura)
            For I = 2 To Lastrow
                If Sheets(1).Cells(I, filos_lectura(J) + 1).Value <> "" And Sheets(1).Cells(I, filos_lectura(J) + 2).Value <> "" Then
                    MsgBox ("Error, solo se puede seleccionar una opción en: " & letrasArray(filos_lectura(J)) & I)
                    Sheets(1).Cells(I, filos_lectura(J)).Select
                    Exit Sub
                End If

                busqueda = Trim(Replace(Sheets(1).Cells(I, filos_lectura(J)).Value," ","*"))
                If Sheets(1).Cells(I, filos_lectura(J) + 1).Value <> "" And Sheets(1).Cells(I, filos_lectura(J)).Value <> "" Then
                    tipo_filo = duro
                Else
                    tipo_filo = suave
                End If
                busqueda = Application.VLookup(busqueda & "*_" & tipo_filo, colores, 2,0)
                If IsError(busqueda) Then
                    MsgBox ("No se encontró el color " & Sheets(1).Cells(I, filos_lectura(J)) & " en: " & letrasArray(filos_lectura(J)) & I)
                    Sheets(1).Cells(I, filos_lectura(J)).Select
                    Exit Sub
                End If
                If Sheets(1).Cells(I, filos_lectura(J)).Value <> "" Then
                    ' Sheets(3).Cells(I + 6, filos_escritura(J)).Value = UCase(Sheets(4).Cells(busqueda, 1).Value) & "_" & tipo_filo
                    Sheets(3).Cells(I + 6, filos_escritura(J)).Value = busqueda
                End If
            Next I
        Next J

        For I = 2 To Lastrow
            If Sheets(1).Cells(I, 7).Value <> "" Then Sheets(3).Cells(I + 6, 6).Value = "S"
            If Sheets(1).Cells(I, 7).Value = "" Then Sheets(3).Cells(I + 6, 6).Value = "N"
        Next I
        Sheets(3).Activate
        MsgBox ("Completado")
End Sub

