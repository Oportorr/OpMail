using OscarSoft;
using System;
using System.Windows.Forms;


namespace CsFoxySmtp
{
    public partial class frmSenMail : Form
    {
        public frmSenMail()
        {
            InitializeComponent();
           
        }

  
        private void frmSenMail_Load(object sender, EventArgs e)
        {
            txtFrom.Text = txtUser.Text;
            cboPriority.SelectedIndex = 0;
            cboSsl.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpMail smtp = new OpMail
            {

                
                server    =  txtServer.Text,
                port      =  int.Parse(txtPort.Text),
                user      =  txtUser.Text,
                password  =  txtPass.Text,
                ssl       =  cboSsl.SelectedIndex == 0 ? true : false,
                priority  =  cboPriority.SelectedIndex,
                from      =  txtFrom.Text,                 
                subjet    =  txtSubject.Text,
                body      =  txtBody.Text,
                timeout   =  100000,

                notification = true            
        };

        smtp.AddTo(txtTo.Text);
        bool result =smtp.Send();

            if (result)
            {
                MessageBox.Show("Mensaje enviado");
            }
            else
            {
                MessageBox.Show(smtp.Error);
            }



        }

        private void txtUser_TextChanged(object sender, EventArgs e)
        {
            txtFrom.Text = txtUser.Text;
        }
    }
}
