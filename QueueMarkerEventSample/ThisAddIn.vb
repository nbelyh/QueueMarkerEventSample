Imports Microsoft.Office.Interop.Visio

Public Class ThisAddIn

    Private Sub Application_MarkerEvent(app As Application, SequenceNum As Integer, ContextString As String) Handles Application.MarkerEvent
        System.Windows.Forms.MessageBox.Show(ContextString)
    End Sub

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
