<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGame
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.menuBar = New System.Windows.Forms.MenuStrip()
        Me.ResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tltTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnQuit = New System.Windows.Forms.Button()
        Me.btnNewGame = New System.Windows.Forms.Button()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.lblHint = New System.Windows.Forms.Label()
        Me.lblWordDisplay = New System.Windows.Forms.Label()
        Me.lblLettersGuessed = New System.Windows.Forms.Label()
        Me.lblAmountOfGuesses = New System.Windows.Forms.Label()
        Me.lblScore = New System.Windows.Forms.Label()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.shpRightLeg = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.shpLeftLeg = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.shpRightArm = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.shpLeftArm = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.shpBody = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.shpHead = New Microsoft.VisualBasic.PowerPacks.OvalShape()
        Me.lblEndGameSubtitle = New System.Windows.Forms.Label()
        Me.lblEndGameTitle = New System.Windows.Forms.Label()
        Me.menuBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'menuBar
        '
        Me.menuBar.BackColor = System.Drawing.SystemColors.MenuBar
        Me.menuBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ResetToolStripMenuItem, Me.HelpToolStripMenuItem, Me.QuitToolStripMenuItem})
        Me.menuBar.Location = New System.Drawing.Point(0, 0)
        Me.menuBar.Name = "menuBar"
        Me.menuBar.Size = New System.Drawing.Size(674, 24)
        Me.menuBar.TabIndex = 0
        Me.menuBar.Text = "MenuStrip1"
        '
        'ResetToolStripMenuItem
        '
        Me.ResetToolStripMenuItem.Name = "ResetToolStripMenuItem"
        Me.ResetToolStripMenuItem.Size = New System.Drawing.Size(47, 20)
        Me.ResetToolStripMenuItem.Text = "Reset"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'QuitToolStripMenuItem
        '
        Me.QuitToolStripMenuItem.Name = "QuitToolStripMenuItem"
        Me.QuitToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.QuitToolStripMenuItem.Text = "Quit"
        '
        'btnQuit
        '
        Me.btnQuit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnQuit.Location = New System.Drawing.Point(25, 509)
        Me.btnQuit.Name = "btnQuit"
        Me.btnQuit.Size = New System.Drawing.Size(138, 37)
        Me.btnQuit.TabIndex = 1
        Me.btnQuit.Text = "Quit"
        Me.btnQuit.UseVisualStyleBackColor = True
        '
        'btnNewGame
        '
        Me.btnNewGame.Location = New System.Drawing.Point(341, 395)
        Me.btnNewGame.Name = "btnNewGame"
        Me.btnNewGame.Size = New System.Drawing.Size(222, 43)
        Me.btnNewGame.TabIndex = 2
        Me.btnNewGame.Text = "New Game"
        Me.btnNewGame.UseVisualStyleBackColor = True
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(25, 466)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(138, 37)
        Me.btnReset.TabIndex = 3
        Me.btnReset.Text = "Reset"
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'lblHint
        '
        Me.lblHint.AutoSize = True
        Me.lblHint.BackColor = System.Drawing.SystemColors.Control
        Me.lblHint.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHint.Location = New System.Drawing.Point(25, 44)
        Me.lblHint.Name = "lblHint"
        Me.lblHint.Size = New System.Drawing.Size(43, 13)
        Me.lblHint.TabIndex = 4
        Me.lblHint.Text = "lblHint"
        '
        'lblWordDisplay
        '
        Me.lblWordDisplay.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblWordDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWordDisplay.Location = New System.Drawing.Point(268, 66)
        Me.lblWordDisplay.Name = "lblWordDisplay"
        Me.lblWordDisplay.Size = New System.Drawing.Size(394, 29)
        Me.lblWordDisplay.TabIndex = 5
        Me.lblWordDisplay.Text = "lblWordDisplay"
        Me.lblWordDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblLettersGuessed
        '
        Me.lblLettersGuessed.AutoSize = True
        Me.lblLettersGuessed.Location = New System.Drawing.Point(22, 210)
        Me.lblLettersGuessed.Name = "lblLettersGuessed"
        Me.lblLettersGuessed.Size = New System.Drawing.Size(91, 13)
        Me.lblLettersGuessed.TabIndex = 6
        Me.lblLettersGuessed.Text = "lblLettersGuessed"
        '
        'lblAmountOfGuesses
        '
        Me.lblAmountOfGuesses.AutoSize = True
        Me.lblAmountOfGuesses.Location = New System.Drawing.Point(22, 265)
        Me.lblAmountOfGuesses.Name = "lblAmountOfGuesses"
        Me.lblAmountOfGuesses.Size = New System.Drawing.Size(105, 13)
        Me.lblAmountOfGuesses.TabIndex = 7
        Me.lblAmountOfGuesses.Text = "lblAmountOfGuesses"
        '
        'lblScore
        '
        Me.lblScore.AutoSize = True
        Me.lblScore.Location = New System.Drawing.Point(25, 312)
        Me.lblScore.Name = "lblScore"
        Me.lblScore.Size = New System.Drawing.Size(45, 13)
        Me.lblScore.TabIndex = 8
        Me.lblScore.Text = "lblScore"
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.shpRightLeg, Me.shpLeftLeg, Me.shpRightArm, Me.shpLeftArm, Me.shpBody, Me.shpHead})
        Me.ShapeContainer1.Size = New System.Drawing.Size(674, 556)
        Me.ShapeContainer1.TabIndex = 9
        Me.ShapeContainer1.TabStop = False
        '
        'shpRightLeg
        '
        Me.shpRightLeg.BorderWidth = 2
        Me.shpRightLeg.Name = "shpRightLeg"
        Me.shpRightLeg.Visible = False
        Me.shpRightLeg.X1 = 499
        Me.shpRightLeg.X2 = 456
        Me.shpRightLeg.Y1 = 381
        Me.shpRightLeg.Y2 = 323
        '
        'shpLeftLeg
        '
        Me.shpLeftLeg.BorderWidth = 2
        Me.shpLeftLeg.Name = "shpLeftLeg"
        Me.shpLeftLeg.Visible = False
        Me.shpLeftLeg.X1 = 455
        Me.shpLeftLeg.X2 = 396
        Me.shpLeftLeg.Y1 = 322
        Me.shpLeftLeg.Y2 = 381
        '
        'shpRightArm
        '
        Me.shpRightArm.BorderWidth = 2
        Me.shpRightArm.Name = "shpRightArm"
        Me.shpRightArm.Visible = False
        Me.shpRightArm.X1 = 456
        Me.shpRightArm.X2 = 502
        Me.shpRightArm.Y1 = 264
        Me.shpRightArm.Y2 = 295
        '
        'shpLeftArm
        '
        Me.shpLeftArm.BorderWidth = 2
        Me.shpLeftArm.Name = "shpLeftArm"
        Me.shpLeftArm.Visible = False
        Me.shpLeftArm.X1 = 395
        Me.shpLeftArm.X2 = 454
        Me.shpLeftArm.Y1 = 297
        Me.shpLeftArm.Y2 = 264
        '
        'shpBody
        '
        Me.shpBody.BorderWidth = 2
        Me.shpBody.Name = "shpBody"
        Me.shpBody.Visible = False
        Me.shpBody.X1 = 454
        Me.shpBody.X2 = 455
        Me.shpBody.Y1 = 233
        Me.shpBody.Y2 = 323
        '
        'shpHead
        '
        Me.shpHead.BorderWidth = 2
        Me.shpHead.Location = New System.Drawing.Point(430, 185)
        Me.shpHead.Name = "shpHead"
        Me.shpHead.Size = New System.Drawing.Size(51, 47)
        Me.shpHead.Visible = False
        '
        'lblEndGameSubtitle
        '
        Me.lblEndGameSubtitle.AutoSize = True
        Me.lblEndGameSubtitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndGameSubtitle.Location = New System.Drawing.Point(365, 137)
        Me.lblEndGameSubtitle.Name = "lblEndGameSubtitle"
        Me.lblEndGameSubtitle.Size = New System.Drawing.Size(135, 15)
        Me.lblEndGameSubtitle.TabIndex = 10
        Me.lblEndGameSubtitle.Text = "lblEndGameSubtitle"
        Me.lblEndGameSubtitle.Visible = False
        '
        'lblEndGameTitle
        '
        Me.lblEndGameTitle.AutoSize = True
        Me.lblEndGameTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndGameTitle.Location = New System.Drawing.Point(364, 117)
        Me.lblEndGameTitle.Name = "lblEndGameTitle"
        Me.lblEndGameTitle.Size = New System.Drawing.Size(141, 20)
        Me.lblEndGameTitle.TabIndex = 11
        Me.lblEndGameTitle.Text = "lblEndGameTitle"
        Me.lblEndGameTitle.Visible = False
        '
        'frmGame
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnQuit
        Me.ClientSize = New System.Drawing.Size(674, 556)
        Me.Controls.Add(Me.lblEndGameTitle)
        Me.Controls.Add(Me.lblEndGameSubtitle)
        Me.Controls.Add(Me.lblScore)
        Me.Controls.Add(Me.lblAmountOfGuesses)
        Me.Controls.Add(Me.lblLettersGuessed)
        Me.Controls.Add(Me.lblWordDisplay)
        Me.Controls.Add(Me.lblHint)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.btnNewGame)
        Me.Controls.Add(Me.btnQuit)
        Me.Controls.Add(Me.menuBar)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MainMenuStrip = Me.menuBar
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmGame"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Hangman Game"
        Me.menuBar.ResumeLayout(False)
        Me.menuBar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents menuBar As System.Windows.Forms.MenuStrip
    Friend WithEvents ResetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QuitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tltTip As System.Windows.Forms.ToolTip
    Friend WithEvents btnQuit As System.Windows.Forms.Button
    Friend WithEvents btnNewGame As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents lblHint As System.Windows.Forms.Label
    Friend WithEvents lblWordDisplay As System.Windows.Forms.Label
    Friend WithEvents lblLettersGuessed As System.Windows.Forms.Label
    Friend WithEvents lblAmountOfGuesses As System.Windows.Forms.Label
    Friend WithEvents lblScore As System.Windows.Forms.Label
    Friend WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents shpLeftLeg As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents shpRightArm As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents shpLeftArm As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents shpBody As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents shpHead As Microsoft.VisualBasic.PowerPacks.OvalShape
    Friend WithEvents shpRightLeg As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents lblEndGameSubtitle As System.Windows.Forms.Label
    Friend WithEvents lblEndGameTitle As System.Windows.Forms.Label

End Class
