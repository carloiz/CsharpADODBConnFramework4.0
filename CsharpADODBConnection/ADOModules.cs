using ADODB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CsharpADODBConnection
{
    internal class ADOModules
    {
        public static string idNum;
        public static Recordset rs;
        public static Command cmd;
         
        public static void viewAllRecord(ListView listView1)
        {
            rs = new Recordset();
            rs.Open("SELECT * FROM players ORDER BY ID DESC", ADODBConn.conn(), CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, 0);

            listView1.Items.Clear();
            while (!rs.EOF)
            {
                string[] rowData = { rs.Fields["ID"].Value.ToString(),
                                     rs.Fields["firstname"].Value.ToString(),
                                     rs.Fields["lastname"].Value.ToString() };
                ListViewItem item = new ListViewItem(rowData);
                listView1.Items.Add(item);
                rs.MoveNext();
            }

            rs.Close();
            ADODBConn.conn().Close();
        }

        public static void addRecord(TextBox text1, TextBox text2, TextBox text3, ListView listView1)
        {
            if (text1.Text == "" && text2.Text == "")
            {
                MessageBox.Show("Fill-up All Fields", "Null");
                return;
            }
            cmd = new Command();
            cmd.ActiveConnection = ADODBConn.conn();
            cmd.CommandText = "INSERT INTO players (firstname, lastname) VALUES (?, ?)";

            cmd.Parameters.Append(cmd.CreateParameter("", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255, text1.Text));
            cmd.Parameters.Append(cmd.CreateParameter("", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255, text2.Text));
            object recordsAffected;
            cmd.Execute(out recordsAffected, Type.Missing, -1);
            viewAllRecord(listView1);
            clearFields(text1, text2, text3);
        }

        public static void updateRecord(TextBox text1, TextBox text2, TextBox text3, ListView listView1)
        {
            if (idNum == "")
            {
                MessageBox.Show("Select a Record", "Null");
                return;
            }
            cmd = new Command();
            cmd.ActiveConnection = ADODBConn.conn();
            cmd.CommandText = "UPDATE players SET firstname=?, lastname=? WHERE ID=?";
            cmd.Parameters.Append(cmd.CreateParameter("", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255, text1.Text));
            cmd.Parameters.Append(cmd.CreateParameter("", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255, text2.Text));
            cmd.Parameters.Append(cmd.CreateParameter("", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255, idNum));
            object recordsAffected;
            cmd.Execute(out recordsAffected, Type.Missing, -1);
            viewAllRecord(listView1);
            clearFields(text1, text2, text3);
        }

        public static void deleteRecord(TextBox text1, TextBox text2, TextBox text3, ListView listView1)
        {
            if (idNum == "")
            {
                MessageBox.Show("Select a Record", "Null");
                return;
            }
            cmd = new Command();
            cmd.ActiveConnection = ADODBConn.conn();
            cmd.CommandText = "DELETE FROM players WHERE ID=?";
            cmd.Parameters.Append(cmd.CreateParameter("", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255, idNum));
            object recordsAffected;
            cmd.Execute(out recordsAffected, Type.Missing, -1);
            viewAllRecord(listView1);
            clearFields(text1, text2, text3);
        }

        public static void searchRecord(TextBox text3, ListView listView1)
        {
            rs = new Recordset();
            cmd = new Command();
            cmd.ActiveConnection = ADODBConn.conn();
            cmd.CommandText = "SELECT * FROM players WHERE ID LIKE ? OR firstname LIKE ? OR lastname LIKE ?";
            cmd.Parameters.Append(cmd.CreateParameter("", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255, "%" + text3.Text + "%"));
            cmd.Parameters.Append(cmd.CreateParameter("", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255, "%" + text3.Text + "%"));
            cmd.Parameters.Append(cmd.CreateParameter("", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255, "%" + text3.Text + "%"));

            rs.Open(cmd);

            listView1.Items.Clear();
            while (!rs.EOF)
            {
                string[] rowData = { rs.Fields["ID"].Value.ToString(),
                                     rs.Fields["firstname"].Value.ToString(),
                                     rs.Fields["lastname"].Value.ToString() };
                ListViewItem item = new ListViewItem(rowData);
                listView1.Items.Add(item);
                rs.MoveNext();
            }

            rs.Close();
            ADODBConn.conn().Close();
        }

        public static void clearFields(TextBox text1, TextBox text2, TextBox text3)
        {
            idNum = "";
            text1.Text = "";
            text2.Text = "";
            text3.Text = "";
        }
    }
}
