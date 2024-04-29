using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CsharpADODBConnection
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ADOModules.viewAllRecord(listView1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ADOModules.addRecord(textBox1, textBox2, textBox3, listView1);

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListView.SelectedListViewItemCollection items = listView1.SelectedItems;

            foreach (ListViewItem item in items) {
                if (item.Selected == true) {
                    ADOModules.idNum = item.SubItems[0].Text;
                    textBox1.Text = item.SubItems[1].Text;
                    textBox2.Text = item.SubItems[2].Text;
                }   
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADOModules.clearFields(textBox1, textBox2, textBox3);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADOModules.updateRecord(textBox1, textBox2, textBox3, listView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADOModules.deleteRecord(textBox1, textBox2, textBox3, listView1);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            ADOModules.searchRecord(textBox3, listView1);
        }
    }
}
