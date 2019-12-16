Public Class MiComboBox

    Inherits System.Windows.Forms.ComboBox

    Private _textbox As TextBox        ' The embedded TextBox control that is used for the ReadOnly mode
    Private _isReadOnly As Boolean     ' True, when the ComboBox is set to ReadOnly
    Private _visible As Boolean = True ' True, when the control is visible
    Private _enabled As Boolean = True  ' True, when the control is Enabled

    Public Sub New()
        _textbox = New TextBox
    End Sub

    ' <summary>
    ' Gets or sets a value indicating whether the control is read-only.
    ' <value>
    ' <true> if the combo box is read-only; otherwise, <false>. The default is <false>.
    ' <remarks>
    ' When this property is set to <true>, the contents of the control cannot be changed 
    ' by the user at runtime. With this property set to <true>, you can still set the value
    ' in code. You can use this feature instead of disabling the control with the Enabled
    ' property to allow the contents to be copied.
    '<Browsable(True), _
    'DefaultValue(False), _
    'Category("Behavior"), _
    'Description("Controla si el valor del control combobox puede ser cambiado o no")> _
    Public Shadows Property [ReadOnly]() As Boolean
        Get
            Return _isReadOnly
        End Get
        Set(ByVal Value As Boolean)
            If Value <> _isReadOnly Then
                _isReadOnly = Value
                ShowControl()
            End If
        End Set
    End Property
    ' <summary>
    ' Gets or sets a value indicating wether the control is displayed.
    ' <value>
    ' <true> if the control is displayed; otherwise, <false>. 
    ' The default is <true>. 
    Public Shadows Property Visible() As Boolean
        Get
            Return _visible
        End Get
        Set(ByVal Value As Boolean)
            _visible = Value
            ShowControl()
        End Set
    End Property
    ' <summary>
    ' Gets or sets a value indicating wether the control is enabled.
    ' <value>
    ' <true> if the control is enabled; otherwise, <false>. 
    ' The default is <true>. 
    Public Shadows Property Enabled() As Boolean
        Get
            Return _enabled
        End Get
        Set(ByVal Value As Boolean)
            _enabled = Value
            ShowControl()
        End Set
    End Property
    ' <summary>
    ' Conceals the control from the user.
    ' <summary>
    ' <remarks>
    ' Hiding the control is equvalent to setting the <see cref="Visible"> property to <false>. 
    ' After the <Hide> method is called, the <Visible> property returns a value of 
    ' <false> until the <see cref="Show"> method is called.
    Public Shadows Sub Hide()
        Visible = False
    End Sub

    ' <summary>
    ' Displays the control to the user.
    ' <remarks>
    ' Showing the control is equivalent to setting the <see cref="Visible"> property to <true>.
    ' After the <Show> method is called, the <Visible> property returns a value of 
    ' <true> until the <see cref="Hide"> method is called.
    Public Shadows Sub Show()
        Visible = True
    End Sub


    ' <summary>
    ' Initializes the embedded TextBox with the default values from the ComboBox
    Private Sub AddTextbox()
        _textbox.ReadOnly = True
        _textbox.Location = Me.Location
        _textbox.Size = Me.Size
        _textbox.Dock = Me.Dock
        _textbox.Anchor = Me.Anchor
        _textbox.Enabled = Me.Enabled
        _textbox.Visible = Me.Visible
        _textbox.RightToLeft = Me.RightToLeft
        _textbox.Font = Me.Font
        _textbox.Text = Me.Text
        _textbox.TabStop = Me.TabStop
        _textbox.TabIndex = Me.TabIndex
    End Sub

    ' <summary>
    ' Shows either the ComboBox or the TextBox or nothing, depending on the state
    ' of the ReadOnly, Enable and Visible flags.
    Private Sub ShowControl()
        _textbox.Text = Me.Text
        If _isReadOnly Then
            _textbox.Visible = _visible
            MyBase.Visible = _visible
            _textbox.BringToFront()
        Else
            _textbox.Visible = False
            MyBase.Visible = _visible
        End If
        _textbox.Enabled = _enabled
        MyBase.Enabled = _enabled
    End Sub

    ' <summary>
    ' This member overrides <see cref="Control.OnParentChanged">
    ' <param name="e">
    Protected Overrides Sub OnParentChanged(ByVal e As EventArgs)
        MyBase.OnParentChanged(e)
        If Parent Is Nothing Then
        Else
            AddTextbox()
            _textbox.Parent = Me.Parent
        End If
    End Sub
    ' <summary>
    ' This member overrides <see cref="ReadOnlyComboBox.OnSelectedIndexChanged">.
    Protected Overrides Sub OnSelectedIndexChanged(ByVal e As EventArgs)
        MyBase.OnSelectedIndexChanged(e)
        If Me.SelectedItem Is Nothing Then
            _textbox.Clear()
        Else
            '_textbox.Text = Me.SelectedItem.ToString()
            _textbox.Text = Me.Text
        End If
    End Sub
    ' <summary>
    ' This member overrides <see cref="ReadOnlyComboBox.OnDropDownStyleChanged">.
    Protected Overrides Sub OnDropDownStyleChanged(ByVal e As EventArgs)
        MyBase.OnDropDownStyleChanged(e)
        _textbox.Text = Me.Text
    End Sub
    ' <summary>
    ' This member overrides <see cref="ReadOnlyComboBox.OnFontChanged">.
    ' <param name="e">
    Protected Overrides Sub OnFontChanged(ByVal e As EventArgs)
        MyBase.OnFontChanged(e)
        _textbox.Font = Me.Font
    End Sub
    ' <summary>
    ' This member overrides <see cref="ReadOnlyComboBox.OnResize">.
    ' <param name="e">
    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        MyBase.OnResize(e)
        _textbox.Size = Me.Size
    End Sub
    ' <summary>
    ' This member overrides <see cref="Control.OnDockChanged">.
    ' <param name="e">
    Protected Overrides Sub OnDockChanged(ByVal e As EventArgs)
        MyBase.OnDockChanged(e)
        _textbox.Dock = Me.Dock
    End Sub
    ' <summary>
    ' This member overrides <see cref="Control.OnEnabledChanged">.
    ' <param name="e">
    Protected Overrides Sub OnEnabledChanged(ByVal e As EventArgs)
        MyBase.OnEnabledChanged(e)
        ShowControl()
    End Sub
    ' <summary>
    ' This member overrides <see cref="Control.OnRightToLeftChanged">.
    ' <param name="e">
    Protected Overrides Sub OnRightToLeftChanged(ByVal e As EventArgs)
        MyBase.OnRightToLeftChanged(e)
        _textbox.RightToLeft = Me.RightToLeft
    End Sub
    ' <summary>
    ' This member overrides <see cref="Control.OnTextChanged">.
    ' <param name="e">
    Protected Overrides Sub OnTextChanged(ByVal e As EventArgs)
        MyBase.OnTextChanged(e)
        _textbox.Text = Me.Text
    End Sub
    ' <summary>
    ' This member overrides <see cref="Control.OnLocationChanged">.
    ' <param name="e">
    Protected Overrides Sub OnLocationChanged(ByVal e As EventArgs)
        MyBase.OnLocationChanged(e)
        _textbox.Location = Me.Location
    End Sub
    ' <summary>
    ' This member overrides <see cref="Control.OnTabIndexChanged">.
    ' <param name="e">
    Protected Overrides Sub OnTabIndexChanged(ByVal e As EventArgs)
        MyBase.OnTabIndexChanged(e)
        _textbox.TabIndex = Me.TabIndex
    End Sub
    ' <summary>
    ' This member overrides <see cref="Control.OnTabStopChanged">.
    ' <param name="e">
    Protected Overrides Sub OnTabStopChanged(ByVal e As EventArgs)
        MyBase.OnTabStopChanged(e)
        _textbox.TabStop = Me.TabStop
    End Sub

    Public Shadows Sub BringToFront()
        If _isReadOnly Then
            _textbox.BringToFront()
        Else
            MyBase.BringToFront()
        End If
    End Sub
    Public Shadows Sub SendToBack()
        If _isReadOnly Then
            _textbox.SendToBack()
        Else
            MyBase.SendToBack()
        End If
    End Sub

End Class
