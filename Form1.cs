//Hello My name is Sankalp
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Nmms_Data_Entry_Software
{
    public partial class Form1 : Form
    {

        string Gender;//stores the gender
        string Category;//stores the category

        int MarksFlag;


        public Form1()
        {
            InitializeComponent();






            textBoxApplicationName.Text = "";
            textBoxRollNo.Text = "";
            textBoxFather.Text = "";

            textBoxMarks.Text = "";
            comboBox1.SelectedItem = null;
            comboBox2.SelectedItem = null;
            string year = dateTimePicker1.Value.ToString().Substring(6, 4);
            Console.Write(year);
            Console.WriteLine(year);
            //Real Time Year

            comboBox2.Items.Add(year);
            for (int z = 0; z < 4; z++)
            {
                int y = Int32.Parse(year);
                y--;
                year = y.ToString();
                comboBox2.Items.Add(year);
            }


            textBoxApplicationName3.Text = "";
            Gender = "";
            Category = "";
            textBoxApplicationName2.Text = "";
            textBoxFather2.Text = "";
            textBoxFather3.Text = "";

        }



        private void button6_Click(object sender, EventArgs e)//Exit Button
        {

        }
        private void Exit()//Exit Method Implementation
        {
            DialogResult iExit;
            iExit = MessageBox.Show("Are you sure you want to quit?", "Caution", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (iExit == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void button1_Click(object sender, EventArgs e)//Add Button
        {
            if ((textBoxApplicationName.Text == "") || (textBoxApplicationName3.Text == "") || (textBoxRollNo.Text == "") || (textBoxFather.Text == "") || (textBoxFather3.Text == "") || (textBoxMarks.Text == "") || (comboBox1.SelectedItem == null) || (Gender == "") || (Category == "") || (comboBox2.SelectedItem == null))
            {
                DialogResult iExit;
                iExit = MessageBox.Show("Please enter full details of the student", "Caution");
            }
            else
            {

                if (MarksFlag == -1)
                {
                    DialogResult iExit;
                    iExit = MessageBox.Show("Please Check the Marks", "Caution");
                }
                else
                {
                    DialogResult iExit;
                    iExit = MessageBox.Show("Are you sure you want to ADD THIS DATA? ", "Caution", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (iExit == DialogResult.Yes)
                    {
                        string f = dateTimePicker1.Value.ToString();
                        string date = f.Substring(0, 10);

                        //Console.WriteLine(s);
                        string roll = textBoxRollNo.Text;

                        //middle Name student
                        string m = textBoxApplicationName2.Text.Trim();
                        string word = "";
                        int count = 0;
                        for (int i = 0; i < m.Length; i++)
                        {
                            if (m[i] != ' ')
                            {
                                word = word + m[i];
                                Console.WriteLine(word);
                                count = 0;




                            }
                            if (m[i] == ' ' && count == 0)
                            {
                                count++;
                                word = word + m[i];


                            }

                        }
                        //middle Name Father

                        //string roll = textBoxRollNo.Text;
                        string mid = textBoxFather2.Text.Trim();
                        string wor = "";
                        int coun = 0;
                        for (int i = 0; i < mid.Length; i++)
                        {
                            if (mid[i] != ' ')
                            {
                                wor = wor + mid[i];

                                coun = 0;




                            }
                            if (mid[i] == ' ' && coun == 0)
                            {
                                coun++;
                                wor = wor + mid[i];


                            }

                        }














                        string applicationName = textBoxApplicationName.Text + " " + word + " " + textBoxApplicationName3.Text;









                        string fatherName = textBoxFather.Text + " " + wor + " " + textBoxFather3.Text;



                        if (textBoxApplicationName2.Text == "")
                        {

                            applicationName = textBoxApplicationName.Text + " " + textBoxApplicationName3.Text;
                        }

                        if (textBoxFather2.Text == "")
                        {

                            fatherName = textBoxFather.Text + " " + textBoxFather3.Text;
                        }




                        dataGridView1.Rows.Add(applicationName,
                                                roll,
                                                date,
                                                fatherName,
                                                Category,
                                                Gender,
                                                comboBox2.SelectedItem.ToString(),
                                                textBoxMarks.Text,
                                                textBoxState.Text,
                                                comboBox1.SelectedItem.ToString()
                                                );
                        //clearing the data
                        textBoxApplicationName.Text = "";
                        textBoxApplicationName2.Text = "";
                        textBoxApplicationName3.Text = "";
                        textBoxRollNo.Text = "";

                        textBoxFather.Text = "";
                        textBoxFather2.Text = "";
                        textBoxFather3.Text = "";

                        radioButton1.Checked = false;//reSetting the gender button
                        radioButton2.Checked = false;
                        radioButton3.Checked = false;
                        radioButton4.Checked = false;//reSetting the Category button
                        radioButton5.Checked = false;
                        radioButton6.Checked = false;
                        radioButton7.Checked = false;

                        comboBox1.SelectedItem = null;//resetting the district
                        comboBox2.SelectedItem = null;

                        textBoxMarks.Text = "";

                        // textBoxDistrict.Text = "";
                        Gender = "";
                        Category = "";

                        textBoxMarks.ReadOnly = true;

                    }
                }
            }
        }

        /*    protected override void OnFormClosing(FormClosingEventArgs e)
            {

                DialogResult iExit;
                iExit = MessageBox.Show("Are you sure you want to quit?", "Caution", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (iExit == DialogResult.Yes)
                {


                }
            }*/

        private void radioButton1_CheckedChanged(object sender, EventArgs e)//Setting up Gender radio Buttons
        {
            Gender = "Male";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)//Setting up Gender radio Buttons
        {
            Gender = "Female";
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)//Setting up Gender radio Buttons
        {
            Gender = "Others";
        }

        private void Delete()//delete method
        {
            DialogResult iExit;
            iExit = MessageBox.Show("Are you sure you want to Delete the Selected Data?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (iExit == DialogResult.Yes)
            {
                foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
                {
                    dataGridView1.Rows.RemoveAt(item.Index);
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)//Delete Buton
        {
            Delete();
        }

        private void button3_Click(object sender, EventArgs e)//Reset Button
        {
            Reset();
        }
        private void Reset()//Reset the Method
        {
            DialogResult iExit;
            iExit = MessageBox.Show("Are you sure you want to Reset the Application?", "Caution", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (iExit == DialogResult.Yes)
            {
                //Clearing all data
                textBoxRollNo.Text = "";

                textBoxApplicationName.Text = "";
                textBoxApplicationName2.Text = "";
                textBoxApplicationName3.Text = "";
                textBoxRollNo.Text = "";

                textBoxFather.Text = "";
                textBoxFather2.Text = "";
                textBoxFather3.Text = "";
                pictureBox2.Visible = false;
                pictureBox3.Visible = false;
                label11.Text = "";

                //pictureBox4.Visible = false;
                label15.Text = "";
                label12.Text = "";
                radioButton1.Checked = false;//reSetting the gender button
                radioButton2.Checked = false;
                radioButton3.Checked = false;
                Gender = "";
                textBoxApplicationName3.Text = "";
                textBoxApplicationName2.Text = "";
                radioButton4.Checked = false;//reSetting the Category button
                radioButton5.Checked = false;
                radioButton6.Checked = false;
                radioButton7.Checked = false;
                comboBox1.SelectedItem = null;
                comboBox2.SelectedItem = null;//resetting the district
                Category = "";


                textBoxMarks.Text = "";
                textBoxMarks.ReadOnly = true;


                //   textBoxDistrict.Text = "";
                //Clearing the data grid
                int numRows = dataGridView1.Rows.Count;
                for (int i = 0; i < numRows; i++)
                {
                    try
                    {
                        int max = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows.Remove(dataGridView1.Rows[max]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("All the Rows are to be deleted" + ex, "DataGridView Delete",
                        MessageBoxButtons.OK, MessageBoxIcon.Information); ;
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)//Export to Excel
        {
            try {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                app.Visible = true;

                // workbook.WritePassword = "12345";
                worksheet = workbook.Sheets["Sheet1"];

                worksheet = workbook.ActiveSheet;
                worksheet.Cells.ColumnWidth = 20;//Excel sheet

                worksheet.Name = "Nmms Report";

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (j == 2)
                        {
                            // Console.WriteLine(dataGridView1.Rows[i].Cells[j].Value);
                            string s = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            string x = s.Substring(0, 2);
                            // Console.WriteLine(int.Parse(s.Substring(2,2)));

                            //Funny Logic Lol..

                            string f = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            string date = f.Substring(3, 3) + f.Substring(0, 2) + f.Substring(5);

                            //   worksheet.Cells[i + 2, j + 1].NumberFormat = "DD-MM-YYYY";
                            worksheet.Cells[i + 2, j + 1] = date;
                            //  Console.WriteLine("1=" + date);
                        }
                        else
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            // Console.WriteLine("3=" + dataGridView1.Rows[i].Cells[j].Value.ToString());
                        }
                    }
                }
            }

            catch (Exception)
            {
                DialogResult iExit;
                iExit = MessageBox.Show("Microsoft Excel Version Not available in the system", "Caution");

            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)//GENERAL
        {
            Category = "1";
            label11.Text = "";
            pictureBox3.Visible = false;
            if (textBoxMarks.Text != "")
            {
                marksValidation();
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)//SC
        {
            Category = "2";
            label11.Text = "";
            pictureBox3.Visible = false;
            if (textBoxMarks.Text != "")
            {
                marksValidation();
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)//ST
        {
            Category = "3";
            label11.Text = "";
            pictureBox3.Visible = false;
            if (textBoxMarks.Text != "")
            {
                marksValidation();
            }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)//OBC
        {
            Category = "4";
            label11.Text = "";
            pictureBox3.Visible = false;
            if (textBoxMarks.Text != "")
            {
                marksValidation();
            }
        }



        private void textBoxMarks_KeyPress(object sender, KeyPressEventArgs e)//Restricting Year field to have only integers
        {
            if (Category == "")
            {
                label11.Text = "Choose the category first";
                label11.ForeColor = Color.Red;
                pictureBox3.Visible = true;
            }
            else
            {
                textBoxMarks.ReadOnly = false;
                char ch = e.KeyChar;
                if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
                {
                    e.Handled = true;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)//Import button
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range xlRange;

            int xlRow;
            string strfileName;

            openFileDialog1.Filter = "Excel Office | *.xls; *.xlsx";
            openFileDialog1.ShowDialog();
            strfileName = openFileDialog1.FileName;
            if (strfileName != string.Empty)
            {
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(strfileName);

                    xlWorkSheet = xlWorkbook.Worksheets["Nmms Report"];
                    xlRange = xlWorkSheet.UsedRange;
                    int i = 0;
                    for (xlRow = 2; xlRow <= xlRange.Rows.Count; xlRow++)
                    {
                        i++;
                        dataGridView1.Rows.Add(xlRange.Cells[xlRow, 1].Text, xlRange.Cells[xlRow, 2].Text, xlRange.Cells[xlRow, 3].Text, xlRange.Cells[xlRow, 4].Text, xlRange.Cells[xlRow, 5].Text, xlRange.Cells[xlRow, 6].Text, xlRange.Cells[xlRow, 7].Text, xlRange.Cells[xlRow, 8].Text, xlRange.Cells[xlRow, 9].Text, xlRange.Cells[xlRow, 10].Text);
                    }

                    xlWorkbook.Close();
                    xlApp.Quit();
                }

                catch (Exception)
                {
                    DialogResult iExit;
                    iExit = MessageBox.Show("Please Select a Valid File", "Caution");
                }

                finally
                {
                    openFileDialog1.Reset();
                }
            }
        }

        private void textBoxRollNo_KeyPress(object sender, KeyPressEventArgs e)//Restricting Roll no  field to have only integers
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }



        private void textBoxApplicationName_KeyPress(object sender, KeyPressEventArgs e)//Applicant name will get only letters
        {
            if (!char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void textBoxFather_KeyPress(object sender, KeyPressEventArgs e)//Father name will get only letters
        {
            if (!char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void textBoxMarks_TextChanged(object sender, EventArgs e)//The marks text box
        {
            marksValidation();
        }
        private void marksValidation()
        {
            try//marks should be out of 180
            {
                if (double.Parse(textBoxMarks.Text) > 180 && (Category == "" || Category == "1" || Category == "2" || Category == "3" || Category == "4"))
                {
                    label11.ForeColor = Color.Red;
                    pictureBox3.Visible = true;
                    label11.Text = "Marks Should be out of 180!!";
                    MarksFlag = -1;
                }
            }
            catch (Exception)
            {
                pictureBox3.Visible = false;
                label11.Text = "";
            }
            //validating for SC and ST
            try
            {
                if ((Category == "2" || Category == "3") && double.Parse(textBoxMarks.Text) < 57.6)
                {
                    label11.ForeColor = Color.Red;
                    pictureBox3.Visible = true;
                    label11.Text = "Marks can't be less than 32 Percent for SC/ST";
                    MarksFlag = -1;
                }
                if ((Category == "2" || Category == "3") && double.Parse(textBoxMarks.Text) >= 57.6 && float.Parse(textBoxMarks.Text) <= 180)
                {
                    pictureBox3.Visible = true;
                    label11.ForeColor = Color.Green;
                    label11.Text = "Accepted!";
                    MarksFlag = 0;
                }
            }
            catch (Exception)
            {
                pictureBox3.Visible = false;
                label11.Text = "";
            }
            //Validating for Gen and Obc
            try
            {
                if ((Category == "1" || Category == "4") && double.Parse(textBoxMarks.Text) < 72)
                {
                    label11.ForeColor = Color.Red;
                    pictureBox3.Visible = true;
                    label11.Text = "Marks can't be less than 40 Percent for Gen/OBC";
                    MarksFlag = -1;
                }
                if ((Category == "1" || Category == "4") && double.Parse(textBoxMarks.Text) >= 72 && float.Parse(textBoxMarks.Text) <= 180)
                {
                    pictureBox3.Visible = true;
                    label11.ForeColor = Color.Green;
                    label11.Text = "Accepted!";
                    MarksFlag = 0;
                }
            }
            catch (Exception)
            {
                pictureBox3.Visible = false;
                label11.Text = "";
            }
        }



        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            if (textBoxApplicationName3.Text == "")
            {
                textBoxApplicationName2.ReadOnly = true;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult iExit;
            iExit = MessageBox.Show("Are you sure you want to go back to home page?", "Caution", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (iExit == DialogResult.Yes)
            {
                this.Hide();
                Form2 f = new Form2();
                f.Show();
            }

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult iExit;
            iExit = MessageBox.Show("Are you sure you want to quit?", "Caution", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (iExit == DialogResult.Yes)
            {

                Application.ExitThread();
            }
            else
            {
                e.Cancel = true;
            }

        }

        private void textBoxApplicationName2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void textBoxApplicationName_Leave(object sender, EventArgs e)
        {
            if (textBoxApplicationName.Text == "")
            {
                pictureBox2.Visible = true;
                label20.ForeColor = Color.Red;
                label20.Text = "Cannot be Empty!!";
            }
            else
            {
                pictureBox2.Visible = false;

                label20.Text = "";

            }

        }

        private void textBoxApplicationName_Enter(object sender, EventArgs e)
        {
            pictureBox2.Visible = false;

            label20.Text = "";
        }

        private void textBoxApplicationName3_Leave(object sender, EventArgs e)
        {
            if (textBoxApplicationName3.Text == "")
            {
                pictureBox4.Visible = true;
                label21.ForeColor = Color.Red;
                label21.Text = "Cannot be Empty!!";
            }
            else
            {
                pictureBox4.Visible = false;

                label21.Text = "";

            }
        }

        private void textBoxApplicationName3_Enter(object sender, EventArgs e)
        {
            pictureBox4.Visible = false;

            label21.Text = "";

        }

        private void textBoxFather_Leave(object sender, EventArgs e)
        {
            if (textBoxFather.Text == "")
            {
                pictureBox5.Visible = true;
                label22.ForeColor = Color.Red;
                label22.Text = "Cannot be Empty!!";
            }
            else
            {
                pictureBox5.Visible = false;

                label22.Text = "";

            }
        }

        private void textBoxFather_Enter(object sender, EventArgs e)
        {
            pictureBox5.Visible = false;

            label22.Text = "";
        }

        private void textBoxFather3_Leave(object sender, EventArgs e)
        {

            if (textBoxFather3.Text == "")
            {
                pictureBox6.Visible = true;
                label23.ForeColor = Color.Red;
                label23.Text = "Cannot be Empty!!";
            }
            else
            {
                pictureBox6.Visible = false;

                label23.Text = "";

            }
        }

        private void textBoxFather3_Enter(object sender, EventArgs e)
        {
            pictureBox6.Visible = false;

            label23.Text = "";
        }

        private void textBoxRollNo_Leave(object sender, EventArgs e)
        {
            if (textBoxRollNo.Text == "")
            {
                pictureBox7.Visible = true;
                label24.ForeColor = Color.Red;
                label24.Text = "Cannot be Empty!!";
            }
            else
            {
                pictureBox7.Visible = false;

                label24.Text = "";

            }
        }

        private void textBoxRollNo_Enter(object sender, EventArgs e)
        {
            pictureBox7.Visible = false;

            label24.Text = "";
        }
    }




    //Red flag code

}

