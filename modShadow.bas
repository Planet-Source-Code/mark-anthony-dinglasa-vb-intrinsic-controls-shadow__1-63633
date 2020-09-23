Attribute VB_Name = "modShadow"
'The only controls that this sub cannot give a shadow are
'"Line Control" and a "Timer Control"

Public Sub GiveShadow(ByRef Frm As Form, _
                       Optional ByVal iColor = &H808080, _
                       Optional ByVal iWidth = 50, _
                       Optional ByVal iHeight = 50)
    Dim Shadow As Object, Ctl As Control
        'ignores error occurence
        On Error Resume Next
            'loop through all controls on a form
            For Each Ctl In Frm.Controls
                'create s shape control
                Set Shadow = Frm.Controls.Add("VB.Shape", Ctl.Name & 0)
                    With Shadow
                        'change shape fill style to solid
                        .FillStyle = vbSolid
                        'change shape border style to transparent
                        .BorderStyle = vbTransparent
                        'Give color to the Shape control
                        .FillColor = iColor
                        'All controls created at runtime is always not visible
                        'So make the shape control visible
                        .Visible = True
                        'move shape control to the found control on the form
                        'with its width and height specified
                        .Move Ctl.Left + iWidth, Ctl.Top + iHeight, Ctl.Width, Ctl.Height
                    End With
                Set Shadow = Nothing 'Clear memory used
            Next
End Sub

'By: Mark Anthony Dinglasa :-D
'email: mark_anthony_dinglasa@yahoo.com
'       markanthonydinglasa@yahoo.com.ph
'Web: www.geocities.com/mark_anthony_dinglasa/2003

