Partial Public Class SalesForm
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

    End Sub

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TitleTextBox = New System.Windows.Forms.TextBox()
        Me.QuantityTextBox = New System.Windows.Forms.TextBox()
        Me.PriceTextBox = New System.Windows.Forms.TextBox()
        Me.ExtendedPriceTextBox = New System.Windows.Forms.TextBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CalculateSaleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SummaryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(69, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "&Title:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(50, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "&Quantity:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(65, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "&Price:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 118)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Extended Price:"
        '
        'TitleTextBox
        '
        Me.TitleTextBox.Location = New System.Drawing.Point(113, 37)
        Me.TitleTextBox.Name = "TitleTextBox"
        Me.TitleTextBox.Size = New System.Drawing.Size(361, 20)
        Me.TitleTextBox.TabIndex = 1
        '
        'QuantityTextBox
        '
        Me.QuantityTextBox.Location = New System.Drawing.Point(113, 63)
        Me.QuantityTextBox.Name = "QuantityTextBox"
        Me.QuantityTextBox.Size = New System.Drawing.Size(125, 20)
        Me.QuantityTextBox.TabIndex = 3
        '
        'PriceTextBox
        '
        Me.PriceTextBox.Location = New System.Drawing.Point(113, 89)
        Me.PriceTextBox.Name = "PriceTextBox"
        Me.PriceTextBox.Size = New System.Drawing.Size(125, 20)
        Me.PriceTextBox.TabIndex = 5
        '
        'ExtendedPriceTextBox
        '
        Me.ExtendedPriceTextBox.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExtendedPriceTextBox.Location = New System.Drawing.Point(113, 115)
        Me.ExtendedPriceTextBox.Name = "ExtendedPriceTextBox"
        Me.ExtendedPriceTextBox.ReadOnly = True
        Me.ExtendedPriceTextBox.Size = New System.Drawing.Size(125, 20)
        Me.ExtendedPriceTextBox.TabIndex = 7
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(532, 24)
        Me.MenuStrip1.TabIndex = 8
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CalculateSaleToolStripMenuItem, Me.ClearToolStripMenuItem, Me.SummaryToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'CalculateSaleToolStripMenuItem
        '
        Me.CalculateSaleToolStripMenuItem.Name = "CalculateSaleToolStripMenuItem"
        Me.CalculateSaleToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.CalculateSaleToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.CalculateSaleToolStripMenuItem.Text = "Calculate &Sale"
        '
        'ClearToolStripMenuItem
        '
        Me.ClearToolStripMenuItem.Name = "ClearToolStripMenuItem"
        Me.ClearToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F2
        Me.ClearToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.ClearToolStripMenuItem.Text = "&Clear"
        '
        'SummaryToolStripMenuItem
        '
        Me.SummaryToolStripMenuItem.Name = "SummaryToolStripMenuItem"
        Me.SummaryToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F3
        Me.SummaryToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.SummaryToolStripMenuItem.Text = "Su&mmary"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'SalesForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(532, 187)
        Me.Controls.Add(Me.ExtendedPriceTextBox)
        Me.Controls.Add(Me.PriceTextBox)
        Me.Controls.Add(Me.QuantityTextBox)
        Me.Controls.Add(Me.TitleTextBox)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "SalesForm"
        Me.Text = "R 'n R Book Sales"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TitleTextBox As System.Windows.Forms.TextBox
    Friend WithEvents QuantityTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PriceTextBox As System.Windows.Forms.TextBox
    Friend WithEvents ExtendedPriceTextBox As System.Windows.Forms.TextBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CalculateSaleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClearToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SummaryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
