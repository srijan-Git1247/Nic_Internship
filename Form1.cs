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
        int YearFlag;
        int MarksFlag;
        int PassFlag;

        public Form1()
        {
            InitializeComponent();
            textBoxApplicationName.Text = "";
            textBoxRollNo.Text = "";
            textBoxFather.Text = "";
            textBoxExam.Text = "";
            textBoxMarks.Text = "";
            comboBox1.SelectedItem = null;
            textBox1.Text = "";
            Gender = "";
            Category = "";
            textBox2.Text = "";
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

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
            if ((textBoxApplicationName.Text == "") || (textBoxRollNo.Text == "") || (textBoxFather.Text == "") || (textBoxExam.Text == "") || (textBoxMarks.Text == "") || (comboBox1.SelectedItem == null) || (Gender == "") || (Category == ""))
            {
                DialogResult iExit;
                iExit = MessageBox.Show("Please enter full details of the student", "Caution");
            }
            else
            {
                if (YearFlag == -1)
                {
                    DialogResult iExit;
                    iExit = MessageBox.Show("Please Enter a valid Exam date", "Caution");
                }
                else if (MarksFlag == -1)
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
                        dataGridView1.Rows.Add(textBoxApplicationName.Text,
                                                roll,
                                                date,
                                                textBoxFather.Text,
                                                Category,
                                                Gender,
                                                textBoxExam.Text,
                                                textBoxMarks.Text,
                                                textBoxState.Text,
                                                comboBox1.SelectedItem.ToString()
                                                );
                        //clearing the data
                        textBoxApplicationName.Text = "";
                        textBoxRollNo.Text = "";

                        textBoxFather.Text = "";

                        radioButton1.Checked = false;//reSetting the gender button
                        radioButton2.Checked = false;
                        radioButton3.Checked = false;
                        radioButton4.Checked = false;//reSetting the Category button
                        radioButton5.Checked = false;
                        radioButton6.Checked = false;
                        radioButton7.Checked = false;

                        comboBox1.SelectedItem = null;//resetting the district
                        textBoxExam.Text = "";
                        textBoxMarks.Text = "";

                        // textBoxDistrict.Text = "";
                        Gender = "";
                        Category = "";

                        textBoxMarks.ReadOnly = true;

                    }
                }
            }
        }

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
                textBoxApplicationName.Text = "";
                textBoxRollNo.Text = "";

                textBoxFather.Text = "";

                pictureBox3.Visible = false;
                label11.Text = "";
                pictureBox2.Visible = false;
                pictureBox4.Visible = false;
                label15.Text = "";
                label12.Text = "";
                radioButton1.Checked = false;//reSetting the gender button
                radioButton2.Checked = false;
                radioButton3.Checked = false;
                Gender = "";
                textBox1.Text = "";
                textBox2.Text = "";
                radioButton4.Checked = false;//reSetting the Category button
                radioButton5.Checked = false;
                radioButton6.Checked = false;
                radioButton7.Checked = false;
                comboBox1.SelectedItem = null;//resetting the district
                Category = "";
                textBoxExam.Text = "";

                textBoxMarks.Text = "";
                textBoxMarks.ReadOnly = true;
                textBox2.ReadOnly = true;

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
            if (textBox1.Text == "")
            {
                DialogResult iExi;
                iExi = MessageBox.Show("Please Enter a password for the excel", "Caution");
            }
            else if (PassFlag != 1)
            {
                DialogResult iExi;
                iExi = MessageBox.Show("Password not matching!!", "Caution");
            }
            else
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                app.Visible = true;

                workbook.WritePassword = textBox1.Text;
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

        private void textBoxExam_KeyPress(object sender, KeyPressEventArgs e)//Restricting Year field to have only integers
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
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

        private void textBoxExam_TextChanged(object sender, EventArgs e)//Validating Exam year between 2021 to 2100
        {
            try
            {
                if (Int32.Parse(textBoxExam.Text) < 2021 || Int32.Parse(textBoxExam.Text) > 2100)
                {
                    pictureBox2.Visible = true;
                    label12.ForeColor = Color.Red;

                    label12.Text = "Enter a Valid Year!";
                    YearFlag = -1;
                }
                else
                {
                    pictureBox2.Visible = true;
                    label12.ForeColor = Color.Green;
                    label12.Text = "Okay!";
                    YearFlag = 0;
                }
            }
            catch (Exception)
            {
                pictureBox2.Visible = false;
                label12.Text = "";
            }
        }

        private void textBoxApplicationName_KeyPress(object sender, KeyPressEventArgs e)//Applicant name will get only letters
        {
            if (!char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void textBoxFather_KeyPress(object sender, KeyPressEventArgs e)//Father name will get only letters
        {
            if (!char.IsLetter(e.KeyChar) && !char.IsControl(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar))
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (textBox1.Text == "")
            {
                pictureBox4.Visible = true;
                label15.ForeColor = Color.Red;
                label15.Text = "Set the password";
            }
            else
            {
                pictureBox4.Visible = false;
                label15.Text = "";
                textBox2.ReadOnly = false;
                Pass();
            }
        }
        private void Pass()
        {
            // Console.WriteLine(textBox1.Text);
            //Console.WriteLine(textBox2.Text);
            if (textBox1.Text.Equals(textBox2.Text) == false && textBox1.Text != "")
            {
                pictureBox4.Visible = true;
                label15.ForeColor = Color.Red;
                label15.Text = "Not Matching!!";
                PassFlag = 0;
            }
            if (textBox1.Text.Equals(textBox2.Text) && textBox1.Text != "")
            {

                pictureBox4.Visible = true;
                label15.ForeColor = Color.Green;
                label15.Text = "Password Matched!";
                PassFlag = 1;


            }
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (pictureBox4.Visible == true && label15.Text != "")
            {
                pictureBox4.Visible = false;
                label15.Text = "";
            }
            if (textBox2.Text != "")
            {
                textBox2.Text = "";
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            Pass();
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox2.ReadOnly = true;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form2 f = new Form2();
            f.ShowDialog();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }
    }
}
