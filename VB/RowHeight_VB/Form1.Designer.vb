﻿Imports Microsoft.VisualBasic
Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Enums
Imports System.Windows.Forms

Namespace GettingStarted
	''' <summary>
	''' Summary description for Form1.
	''' </summary>
	Partial Public Class Form1
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.Container = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.sfDataGrid = New Syncfusion.WinForms.DataGrid.SfDataGrid()
			Me.button2 = New System.Windows.Forms.Button()
			CType(Me.sfDataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.SuspendLayout()
			' 
			' sfDataGrid
			' 
			Me.sfDataGrid.AccessibleName = "Table"
			Me.sfDataGrid.AllowDraggingColumns = True
			Me.sfDataGrid.AllowFiltering = True
			Me.sfDataGrid.AllowResizingColumns = True
			Me.sfDataGrid.AllowResizingHiddenColumns = True
			Me.sfDataGrid.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.sfDataGrid.AutoSizeColumnsMode = Syncfusion.WinForms.DataGrid.Enums.AutoSizeColumnsMode.Fill
			Me.sfDataGrid.BackColor = System.Drawing.SystemColors.Window
			Me.sfDataGrid.Location = New System.Drawing.Point(10, 5)
			Me.sfDataGrid.Name = "sfDataGrid"
			Me.sfDataGrid.ShowGroupDropArea = False
			Me.sfDataGrid.Size = New System.Drawing.Size(881, 576)
			Me.sfDataGrid.TabIndex = 1
			' 
			' button2
			' 
			Me.button2.Location = New System.Drawing.Point(897, 12)
			Me.button2.Name = "button2"
			Me.button2.Size = New System.Drawing.Size(84, 37)
			Me.button2.TabIndex = 3
			Me.button2.Text = "Export To Excel"
			Me.button2.UseVisualStyleBackColor = True
'			Me.button2.Click += New System.EventHandler(Me.button2_Click)
			' 
			' Form1
			' 
			Me.BackColor = System.Drawing.SystemColors.Window
			Me.ClientSize = New System.Drawing.Size(1006, 593)
			Me.Controls.Add(Me.button2)
			Me.Controls.Add(Me.sfDataGrid)
			Me.Margin = New System.Windows.Forms.Padding(5)
			Me.MinimumSize = New System.Drawing.Size(500, 400)
			Me.Name = "Form1"
			Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
			Me.Text = "Getting Started"
			CType(Me.sfDataGrid, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub

		#End Region

		#Region "API Definition"

		Private sfDataGrid As SfDataGrid

		#End Region

		Private WithEvents button2 As Button
	End Class
End Namespace

