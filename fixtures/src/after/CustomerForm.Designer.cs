using System.Drawing;
using System.Windows.Forms;

namespace CustomerManagement
{
    public partial class CustomerForm
    {
        private TextBox txtCusNm = new TextBox
        {
            Location = new Point(24, 24),
            Size = new Size(264, 24),
            TabIndex = 0
        };

        private TextBox txtCusNmKana = new TextBox
        {
            Location = new Point(24, 64),
            Size = new Size(264, 24),
            TabIndex = 1
        };

        private void InitializeComponent()
        {
            SuspendLayout();
            txtCusNm.MaxLength = 20;
            txtCusNmKana.MaxLength = 20;
            Controls.AddRange(new Control[] { txtCusNm, txtCusNmKana });
            ClientSize = new Size(312, 112);
            Text = "顧客情報";
            ResumeLayout(false);
            PerformLayout();
        }
    }
}
