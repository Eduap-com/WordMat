using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;


namespace MathFormForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            mprocess = new MaximaProcessClass();
        }
        private MaximaProcessClass mprocess;
        private void button1_Click(object sender, EventArgs e)
        {
            int tcount=0;
//            mprocess.OutUnits = "eV,nm";
//            mprocess.Units = 1;
            mprocess.ExecuteMaximaCommand(textBox1.Text, 0);
            do
            {
                Thread.Sleep(20);
                tcount = tcount + 1;
            } while (mprocess.Finished == 0 & tcount < 400);
//            if (mprocess.Finished==0)

            if (mprocess.Question==1)
            {
                DialogResult res;
                string svar="y";
                res = InputBox("spr", mprocess.QuestionText, ref svar);
                mprocess.AnswerQuestion(svar);
                MessageBox.Show(mprocess.LastMaximaOutput);
                   
            }
            label1.Text = mprocess.LastMaximaOutput;
            textBox2.Text = mprocess.MaximaOutput;
            label1.Refresh();

            mprocess.Reset("");
//            MessageBox.Show(mprocess.MaximaOutputArray(4));
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 60);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }

        private void button_restart_Click(object sender, EventArgs e)
        {
            mprocess.CloseProcess();
            mprocess.StartMaximaProcess();
            //mprocess.ConsoleInterrupt();
            //mprocess.Reset();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            mprocess.Units = 1;
            mprocess.OutUnits = "g,ms";
            mprocess.TurnUnitsOn("","");
            mprocess.Units = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show(mprocess.CheckForUpdate());
        }
    }
}
