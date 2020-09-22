Attribute VB_Name = "modObject"
'This is my original code

Public Sub changeSize(ByVal obj As Object, ByVal times As Integer, ByVal divide As Integer)
    obj.Visible = False
    obj.FontSize = obj.FontSize * times / divide
    obj.Width = obj.Width * times / divide
    obj.Height = obj.Height * times / divide
    obj.Left = obj.Left * times / divide
    obj.Top = obj.Top * times / divide
    obj.Visible = True
End Sub

