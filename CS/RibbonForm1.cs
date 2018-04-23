using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace BizPad
{
    public partial class RibbonForm1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public RibbonForm1()
        {
            InitializeComponent();
            richEditControl1.LoadDocument("Chart.rtf");
            _ribbon.SelectedPage = this.mailingsRibbonPage1;
        }

        private void richEditControl1_CalculateDocumentVariable(object sender, DevExpress.XtraRichEdit.CalculateDocumentVariableEventArgs e)
        {
            if (e.VariableName == "CHART") {
                ChartImage chart = new ChartImage(e.Arguments[0].Value.ToString());
                chart.Initialize();
                DocumentImageSource image = chart.CreateImage();
                RichEditDocumentServer srv = new RichEditDocumentServer();
                srv.Document.AppendImage(image);
                e.Value = srv.Document;
                e.Handled = true;
            }
        }
    }
}