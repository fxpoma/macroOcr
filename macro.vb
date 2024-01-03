Option Compare Text
Sub COMPILE()

    '%%%% LASTROW %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    Dim wS, wS1 As Worksheet, Lastrow, Lastrow1 As Long
    Dim columnas, repeticiones, almacenador As Variant ' Declara un array
    Dim I, J As Long
    Dim duro, suave As String
    Set wS = Sheets(1)

    duro = "22X2"
    suave = "22X045"
    almacenador = 0

    columnas = Array(6, 8, 9, 11, 12, 14, 15, 17, 18)

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

        For I = 2 To Lastrow
            c1 = 0
            c2 = 0
            c3 = 0
            c4 = 0
            Sheets(3).Cells(I + 6, 3).Value = Sheets(1).Cells(I, 3).Value
            Sheets(3).Cells(I + 6, 4).Value = Sheets(1).Cells(I, 4).Value
            Sheets(3).Cells(I + 6, 5).Value = Sheets(1).Cells(I, 5).Value
            ' Sheets(3).Cells(I + 6, 7).Value = Sheets(1).Cells(I, 7).Value ' L1
            ' Sheets(3).Cells(I + 6, 9).Value = Sheets(1).Cells(I, 10).Value ' L2
            ' Sheets(3).Cells(I + 6, 11).Value = Sheets(1).Cells(I, 13).Value ' A1
            ' Sheets(3).Cells(I + 6, 13).Value = Sheets(1).Cells(I, 16).Value ' A2
            Sheets(3).Cells(I + 6, 11).Value = Sheets(1).Cells(I, 19).Value
            Sheets(3).Cells(I + 6, 12).Value = Sheets(1).Cells(I, 20).Value

            c1 = 0
            ' '''''''''''''''' L1 '''''''''''''''''''''''''''''''''''''''
            If Sheets(1).Cells(I, 8).Value <> "" And Sheets(1).Cells(I, 7).Value <> "" Then
                Sheets(3).Cells(I + 6, 7).Value = Sheets(1).Cells(I, 7).Value & "_" & duro ' Duro
                c1 = c1 + 1
            End If

            If Sheets(1).Cells(I, 9).Value <> "" And Sheets(1).Cells(I, 7).Value <> "" Then
                Sheets(3).Cells(I + 6, 7).Value = Sheets(1).Cells(I, 7).Value & "_" & suave
                c1 = c1 + 1
            End If

            If c1 = 2 Then
                Sheets(3).Cells(I + 6, 7).Value = "ERR"
                MsgBox ("Error, ¡sólo debe seleccionar una opción! in L1")
             Exit Sub
            End If
            ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ' '''''''''''''''' L2 '''''''''''''''''''''''''''''''''''''''
            If Sheets(1).Cells(I, 11).Value <> "" And Sheets(1).Cells(I, 10).Value <> "" Then
                Sheets(3).Cells(I + 6, 8).Value = Sheets(1).Cells(I, 10).Value & "_" & duro
                c2 = c2 + 1
            End If

            If Sheets(1).Cells(I, 12).Value <> "" And Sheets(1).Cells(I, 10).Value <> "" Then
                Sheets(3).Cells(I + 6, 8).Value = Sheets(1).Cells(I, 10).Value & "_" & suave
                c2 = c2 + 1
            End If

            If c2 = 2 Then
                Sheets(3).Cells(I + 6, 8).Value = "ERR"
                MsgBox ("Error, ¡sólo debe seleccionar una opción! in L2")
             Exit Sub
            End If
            ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ' '''''''''''''''' A1 '''''''''''''''''''''''''''''''''''''''
            If Sheets(1).Cells(I, 14).Value <> "" And Sheets(1).Cells(I, 13).Value <> "" Then
                Sheets(3).Cells(I + 6, 9).Value = Sheets(1).Cells(I, 13).Value & "_" & duro
                c3 = c3 + 1
            End If

            If Sheets(1).Cells(I, 15).Value <> "" And Sheets(1).Cells(I, 13).Value <> "" Then
                Sheets(3).Cells(I + 6, 9).Value = Sheets(1).Cells(I, 13).Value & "_" & suave
                c3 = c3 + 1
            End If

            If c3 = 2 Then
                Sheets(3).Cells(I + 6, 9).Value = "ERR"
                MsgBox ("Error, ¡sólo debe seleccionar una opción! in A1")
             Exit Sub
            End If
            ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ' '''''''''''''''' A2 '''''''''''''''''''''''''''''''''''''''
            If Sheets(1).Cells(I, 17).Value <> "" And Sheets(1).Cells(I, 16).Value <> "" Then
                Sheets(3).Cells(I + 6, 10).Value = Sheets(1).Cells(I, 16).Value & "_" & duro
                c4 = c4 + 1
            End If

            If Sheets(1).Cells(I, 18).Value <> "" And Sheets(1).Cells(I, 16).Value <> "" Then
                Sheets(3).Cells(I + 6, 10).Value = Sheets(1).Cells(I, 16).Value & "_" & suave
                c4 = c4 + 1
            End If

            If c4 = 2 Then
                Sheets(3).Cells(I + 6, 10).Value = "ERR"
                MsgBox ("Error, ¡sólo debe seleccionar una opción! in A2")
             Exit Sub
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            '' VETA
            If Sheets(1).Cells(I, 6).Value <> "" Then Sheets(3).Cells(I + 6, 6).Value = "Y"
                If Sheets(1).Cells(I, 6).Value = "" Then Sheets(3).Cells(I + 6, 6).Value = "N"

                    ''''''''''

                Next I


                MsgBox ("Completado")

End Sub


