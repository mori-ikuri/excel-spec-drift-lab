using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CustomerManagement
{
    public partial class CustomerForm : Form
    {
        public CustomerForm()
        {
            InitializeComponent();
        }

        public void ShowCustomer(IDictionary<string, object> data)
        {
            var cusNm = data[Columns.CUSTOMER_NAME].ToString();
            if (cusNm.Length > 20)
            {
                throw new ArgumentException("入力値が長すぎます。");
            }
            txtCusNm.Text = cusNm;
        }
    }
}
