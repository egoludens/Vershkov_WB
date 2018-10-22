using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VershkovWB
{
    public partial class WBReportProcessorForm : Form, IProgressReportObserver
    {
        public WBReportProcessorForm()
        {
            InitializeComponent();
        }

        private void sourceFileNameButton_Click(object sender, EventArgs e)
        {
            openSourceFileDialog.ShowDialog();
        }

        private void openSourceFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            sourceFileNameText.Text = openSourceFileDialog.FileName;
        }

        private void processWBReportButton_Click(object sender, EventArgs e)
        {
            progressReportText.Clear();
            WBReportProcessor reportProcessor = new WBReportProcessor(sourceFileNameText.Text);
            reportProcessor.AddProgressReportObserver(this);
            reportProcessor.Process();
        }

        public void NotifyAboutProgressReport(string progressReportUpdate)
        {
            progressReportText.AppendText(progressReportUpdate + "\r\n");
        }
    }
}
