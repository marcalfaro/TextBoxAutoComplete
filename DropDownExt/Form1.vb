Option Explicit On

Public Class Form1

    Dim acColors As New AutoCompleteStringCollection()
    Dim acAnimal As New AutoCompleteStringCollection()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        acColors.AddRange(New String() {"Red", "Orange", "Red Orange", "Blue", "Green", "Blue Green", "Black", "White"})
        acAnimal.AddRange(New String() {"Zebra Fish", "Maine Coon", "Belgian Malinois", "Bull Dog", "Lazy Dog", "Goldfish", "Bull", "Cat"})

        Build_DGV()

        TextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox1.AutoCompleteCustomSource = acAnimal

    End Sub



    Private Sub Build_DGV()
        Try
            With dgv.Columns
                .Add(New DataGridViewTextBoxColumn With {.Name = "colColor", .HeaderText = "Color"})
                .Add(New DataGridViewTextBoxColumn With {.Name = "colPet", .HeaderText = "Pet"})
            End With

        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub

    Private Sub dgv_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv.EditingControlShowing
        If TypeOf e.Control Is TextBox Then
            Dim tb As TextBox = CType(e.Control, TextBox)

            If Not IsNothing(tb) Then
                tb.AutoCompleteMode = AutoCompleteMode.None
                tb.AutoCompleteSource = AutoCompleteSource.None
                tb.AutoCompleteCustomSource = Nothing

                Select Case dgv.CurrentCell.ColumnIndex
                    Case 0
                        If acColors IsNot Nothing Then
                            tb.AutoCompleteMode = AutoCompleteMode.Suggest
                            tb.AutoCompleteSource = AutoCompleteSource.CustomSource
                            tb.AutoCompleteCustomSource = acColors
                        Else
                            tb = Nothing
                        End If

                    Case 1
                        If acAnimal IsNot Nothing Then
                            tb.AutoCompleteMode = AutoCompleteMode.Suggest
                            tb.AutoCompleteSource = AutoCompleteSource.CustomSource
                            tb.AutoCompleteCustomSource = acAnimal
                        Else
                            tb = Nothing
                        End If

                End Select



            End If

        End If

    End Sub
End Class
