# QueueMarkerEventSample
A Sample illustrating using QueueMarkerAddin in VIsio and VB.NET

Reference:
To use QueueMarkerAddin you just add a method to your app class like this (VB.NET)

```vb
Public Class ThisAddIn

  Private Sub Application_MarkerEvent(app As Application, SequenceNum As Integer, ContextString As String) Handles Application.MarkerEvent
    System.Windows.Forms.MessageBox.Show(ContextString)
  End Sub
```

In Visio you just call
![](https://i.paste.pics/613b20f0da638a4f63a2669e22cd04b7.png)

Please note that if you just want to have a context menu, you can just use ribbon context menus instead.
