Class MainWindow

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        cmb_cps.ItemsSource = Enumerable.Range(1, 100)

    End Sub

End Class
