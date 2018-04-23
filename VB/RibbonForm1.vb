Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraBars
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace BizPad
	Partial Public Class RibbonForm1
		Inherits DevExpress.XtraBars.Ribbon.RibbonForm
		Public Sub New()
			InitializeComponent()
			richEditControl1.LoadDocument("Chart.rtf")
			_ribbon.SelectedPage = Me.mailingsRibbonPage1
		End Sub

		Private Sub richEditControl1_CalculateDocumentVariable(ByVal sender As Object, ByVal e As DevExpress.XtraRichEdit.CalculateDocumentVariableEventArgs) Handles richEditControl1.CalculateDocumentVariable
			If e.VariableName = "CHART" Then
				Dim chart As New ChartImage(e.Arguments(0).Value.ToString())
				chart.Initialize()
				Dim image As DocumentImageSource = chart.CreateImage()
				Dim srv As New RichEditDocumentServer()
				srv.Document.AppendImage(image)
				e.Value = srv.Document
				e.Handled = True
			End If
		End Sub
	End Class
End Namespace