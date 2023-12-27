using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq.Expressions;
using Microsoft.Office.Interop.Excel;
using Point = System.Drawing.Point;
using CheckBox = System.Windows.Forms.CheckBox;
using System.Globalization;
using Newtonsoft.Json;
using System.IO;

namespace Schedule_project
{
    public partial class MainForm : Form
    {
        List<Employee> Employees  = new List<Employee>();
        public credentials credentials = new credentials();
        Excel.Application excelApp;
        Excel.Worksheet worksheet;
        Excel.Range usedRange;
        Excel.Range currentRow;
        Excel.Workbook workbook;
        string fromAddress;
        string fromPassword;
        string smtpClient;
        Employee selectedEmployee = new Employee();
        private bool isProgrammaticEdit = false;
        public MainForm()
        {
            isProgrammaticEdit = true;
            InitializeComponent();
            MainForm_Resize();
            programStart();
            ReadTheExcel();
            dataGridView1.Columns["ElectricalTrainingSafetyColumn"].Visible = false;
            dataGridView1.Columns["Essity1YearColumn"].Visible = false;
            dataGridView1.Columns["FamiliarizedColumn"].Visible = false;
            dataGridView1.Columns["FireWorkingDateColumn"].Visible = false;
            dataGridView1.Columns["FireWorkingNumberColumn"].Visible = false;
            dataGridView1.Columns["FirstAidTrainingColumn"].Visible = false;
            dataGridView1.Columns["LiveWorkingTrainingColumn"].Visible = false;
            dataGridView1.Columns["NokianRenkat1YearColumn"].Visible = false;
            dataGridView1.Columns["NokianRenkatLOTOColumn"].Visible = false;
            dataGridView1.Columns["NvEColumn"].Visible = false;
            dataGridView1.Columns["OtherColumn"].Visible = false;
            dataGridView1.Columns["SandvikColumn"].Visible = false;
            dataGridView1.Columns["TampereenSahkolaitosColumn"].Visible = false;
            dataGridView1.Columns["TaxNumberColumn"].Visible = false;
            dataGridView1.Columns["ValttikorttiColumn"].Visible = false;
            dataGridView1.Columns["WorkSafetyTrainingColumn"].Visible = false;
            dataGridView1.Columns["WorkSafetyTrainingNumberColumn"].Visible = false;
            credentials = RestoreSettings(".\\" + "credentials.json");
            CheckForDeadlines();
            isProgrammaticEdit = false;
            dataGridView1.DefaultCellStyle.SelectionBackColor = dataGridView1.DefaultCellStyle.BackColor;
            dataGridView1.DefaultCellStyle.SelectionForeColor = dataGridView1.DefaultCellStyle.ForeColor;
        }

        Color brighterColor = Color.FromArgb(
            (int)(SystemColors.Highlight.R * 1.1f),
            (int)(SystemColors.Highlight.G * 1.1f),
            (int)(SystemColors.Highlight.B * 0.8f)
        );

        private void MainForm_Resize(object sender, EventArgs e)
        {
            //Adjusting internal controls based on form's new size
            MainForm_Resize();
        }
        private void MainForm_Resize()
        {
            //Adjusting internal controls based on form's new size
            panel1.Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 35);
            panel1.Location = new Point(10, 20);
            deleteButton.Size = new Size((int)(this.panel1.Width * 0.129), (int)(this.panel1.Height * 0.045));
            EmailButton.Size = new Size((int)(this.panel1.Width * 0.129), (int)(this.panel1.Height * 0.045));
            AddEmployeeButton.Size = new Size((int)(this.panel1.Width * 0.129), (int)(this.panel1.Height * 0.045));
            deleteButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.031));
            EmailButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.0745));
            AddEmployeeButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.118));
            FamiliarizedBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.1615));
            FireDateBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.205));
            FireNumberBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.248));
            WorkSafetyBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.292));
            WorkSafetyTrainingBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.335));
            ElectricalSafetyBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.379));
            LiveBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.422));
            FirstAidBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.466));
            ValttikorttiBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.509));
            TaxBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.553));
            NokianR1YBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.596));
            NokianRLOTOBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.63975));
            Essity1YBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.683));
            TampereenSBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.7267));
            NvEBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.7702));
            SandvikBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.814));
            OtherBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.857));
            panel2.Location = new Point((int)(this.referenceLabel.Location.X * 0.2), (int)(this.referenceLabel.Location.Y * 0.1));
            panel2.Size = new Size((int)(this.panel1.Width * 0.75), (int)(this.panel1.Height * 0.85));
            dataGridView1.Location = new Point((int)(this.referenceLabel.Location.X * 0.2), (int)(this.referenceLabel.Location.Y * 0.1));
            dataGridView1.Size = new Size((int)(this.panel1.Width * 0.7), (int)(this.panel1.Height * 0.7));
            button1.Size = new Size((int)(this.panel1.Width * 0.184), (int)(this.panel1.Height * 0.069));
            button1.Location = new Point((int)(this.referenceLabel.Location.X * 0.55), (int)(this.referenceLabel.Location.Y * 0.65));
            emailTextBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.05), (int)(this.referenceLabel.Location.Y * 0.1));
            passwordTextBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.05), (int)(this.referenceLabel.Location.Y * 0.25));
            serviceTextBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.05), (int)(this.referenceLabel.Location.Y * 0.40));
            twoFactorBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.05), (int)(this.referenceLabel.Location.Y * 0.55));
            label1.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.05));
            label2.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.2));
            label3.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.30));
            label4.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.35));
            label5.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.50));
        }

        private void programStart()
        {
            excelApp = new Excel.Application();
            // Open the Excel workbook
            workbook = excelApp.Workbooks.Open(@"/*FOLDERHERE*/");
            worksheet = workbook.Sheets[1];
            usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
        }

        private void Add_Click(object sender, EventArgs e)
        {
            isProgrammaticEdit = true;
            DialogResult resultChoice = DialogResult.Cancel;
            NewEmployeeForm MChoiceForm = new NewEmployeeForm();
            resultChoice = MChoiceForm.ShowDialog(this);

            if (resultChoice == DialogResult.OK)
            {
                Employee newEmployee = MChoiceForm.sendNewEmployee();
                Employees.Add(newEmployee);
                int currentLine = Employees.Count + 1;
                usedRange.Cells[currentLine + 2, 1].Value2 = newEmployee.EmployeeName;//
                    if (newEmployee.Familiarized == true)
                        usedRange.Cells[currentLine + 2, 2].Value2 = "OK";
                    else
                        usedRange.Cells[currentLine + 2, 2].Value2 = "";
                    usedRange.Cells[currentLine + 2, 3].Value2 = convert(newEmployee.FireWorkingDate);//
                    usedRange.Cells[currentLine + 2, 4].Value2 = newEmployee.FireWorkingNumber;//
                    usedRange.Cells[currentLine + 2, 5].Value2 = convert(newEmployee.WorkSafetyTraining);//
                    usedRange.Cells[currentLine + 2, 6].Value2 = newEmployee.WorkSafetyTrainingNumber;//
                    usedRange.Cells[currentLine + 2, 7].Value2 = convert(newEmployee.ElectricalSafetyTraining);//
                    usedRange.Cells[currentLine + 2, 8].Value2 = convert(newEmployee.LiveWorkingTraining);//
                    usedRange.Cells[currentLine + 2, 9].Value2 = convert(newEmployee.FirstAidTraining);//
                    usedRange.Cells[currentLine + 2, 10].Value2 = newEmployee.Valttikortti;//
                    usedRange.Cells[currentLine + 2, 11].Value2 = newEmployee.TaxNumber;//
                    usedRange.Cells[currentLine + 2, 12].Value2 = convert(newEmployee.NokianRenkat1Year);//
                    usedRange.Cells[currentLine + 2, 13].Value2 = convert(newEmployee.NokianRenkatLOTO);//
                    usedRange.Cells[currentLine + 2, 14].Value2 = convert(newEmployee.Essity1Year);//
                    usedRange.Cells[currentLine + 2, 15].Value2 = convert(newEmployee.TampereenSahkolaitos);//
                    usedRange.Cells[currentLine + 2, 16].Value2 = convert(newEmployee.NvE);//
                    usedRange.Cells[currentLine + 2, 17].Value2 = convert(newEmployee.Sandvik);//
                    usedRange.Cells[currentLine + 2, 18].Value2 = newEmployee.Other; //

                Excel.Range newRowRange = usedRange.Rows[currentLine + 2];

                // Set borders for the entire row
                newRowRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                newRowRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                workbook.Save();

                dataGridView1.Rows.Add(newEmployee.EmployeeName, newEmployee.Familiarized, newEmployee.FireWorkingDate, newEmployee.FireWorkingNumber, newEmployee.WorkSafetyTraining, newEmployee.WorkSafetyTrainingNumber,
                            newEmployee.ElectricalSafetyTraining, newEmployee.LiveWorkingTraining, newEmployee.FirstAidTraining, newEmployee.Valttikortti, newEmployee.TaxNumber, newEmployee.NokianRenkat1Year, newEmployee.NokianRenkatLOTO,
                            newEmployee.Essity1Year, newEmployee.TampereenSahkolaitos, newEmployee.NvE, newEmployee.Sandvik, newEmployee.Other);
            }
            isProgrammaticEdit = false;
        }

        public string convert(string input)
        {
            try
            {
                double excelDateValue = Convert.ToDouble(input);
                DateTime dateTimeValue = DateTime.FromOADate(excelDateValue);
                return dateTimeValue.ToString("MM.dd.yyyy");
            }
            catch 
            {
                return input;
            }
        }

        public void ReadTheExcel()
        {
            isProgrammaticEdit = true;
            try 
            {
                int rowCount = usedRange.Rows.Count;

                int i = 4;
                currentRow = usedRange.Rows[i];
                try
                {
                    while (currentRow.Cells[1, 1].Value2 != null) // Check if the first cell in the row is empty
                    {
                        Employee employee = new Employee();
                        int colCount = Math.Min(usedRange.Columns.Count, 18);

                        employee.EmployeeName = currentRow.Cells[1, 1].Value2.ToString();
                        if (currentRow.Cells[1, 2].Value2 != null)
                        {
                            if (currentRow.Cells[1, 2].Value2.ToString() == "OK")
                                employee.Familiarized = true;
                            else
                                employee.Familiarized = false;
                        }
                        if (currentRow.Cells[1, 3].Value2 != null)
                            employee.FireWorkingDate = convert(currentRow.Cells[1, 3].Value2.ToString());
                        if (currentRow.Cells[1, 4].Value2 != null)
                            employee.FireWorkingNumber = currentRow.Cells[1, 4].Value2.ToString();
                        if (currentRow.Cells[1, 5].Value2 != null)
                            employee.WorkSafetyTraining = convert(currentRow.Cells[1, 5].Value2.ToString());
                        if (currentRow.Cells[1, 6].Value2 != null)
                            employee.WorkSafetyTrainingNumber = currentRow.Cells[1, 6].Value2.ToString();
                        if (currentRow.Cells[1, 7].Value2 != null)
                            employee.ElectricalSafetyTraining = convert(currentRow.Cells[1, 7].Value2.ToString());
                        if (currentRow.Cells[1, 8].Value2 != null)
                            employee.LiveWorkingTraining = convert(currentRow.Cells[1, 8].Value2.ToString());
                        if (currentRow.Cells[1, 9].Value2 != null)
                            employee.FirstAidTraining = convert(currentRow.Cells[1, 9].Value2.ToString());
                        if (currentRow.Cells[1, 10].Value2 != null)
                            employee.Valttikortti = currentRow.Cells[1, 10].Value2.ToString();
                        if (currentRow.Cells[1, 11].Value2 != null)
                            employee.TaxNumber = currentRow.Cells[1, 11].Value2.ToString();
                        if (currentRow.Cells[1, 12].Value2 != null)
                            employee.NokianRenkat1Year = convert(currentRow.Cells[1, 12].Value2.ToString());
                        if (currentRow.Cells[1, 13].Value2 != null)
                            employee.NokianRenkatLOTO = convert(currentRow.Cells[1, 13].Value2.ToString());
                        if (currentRow.Cells[1, 14].Value2 != null)
                            employee.Essity1Year = convert(currentRow.Cells[1, 14].Value2.ToString());
                        if (currentRow.Cells[1, 15].Value2 != null)
                            employee.TampereenSahkolaitos = convert(currentRow.Cells[1, 15].Value2.ToString());
                        if (currentRow.Cells[1, 16].Value2 != null)
                            employee.NvE = convert(currentRow.Cells[1, 16].Value2.ToString());
                        if (currentRow.Cells[1, 17].Value2 != null)
                            employee.Sandvik = convert(currentRow.Cells[1, 17].Value2.ToString());
                        if (currentRow.Cells[1, 18].Value2 != null)
                            employee.Other = currentRow.Cells[1, 18].Value2.ToString();

                        i++;
                        currentRow = usedRange.Rows[i];
                        Employees.Add(employee);
                        dataGridView1.Rows.Add(employee.EmployeeName, employee.Familiarized, employee.FireWorkingDate, employee.FireWorkingNumber, employee.WorkSafetyTraining, employee.WorkSafetyTrainingNumber, 
                            employee.ElectricalSafetyTraining, employee.LiveWorkingTraining, employee.FirstAidTraining, employee.Valttikortti, employee.TaxNumber, employee.NokianRenkat1Year, employee.NokianRenkatLOTO, 
                            employee.Essity1Year, employee.TampereenSahkolaitos, employee.NvE, employee.Sandvik, employee.Other);
                    }
                }
                catch
                {
                } 
            }
            catch
            {
            }
            workbook.Save();
            isProgrammaticEdit = false;

        }

        private void Check_Uncheck(object sender, EventArgs e)
        {
            CheckBox senderCheck = sender as CheckBox;
            if (senderCheck.Checked)
            {
                switch (senderCheck.Name)
                {
                    case "ElectricalSafetyBox":
                        dataGridView1.Columns["ElectricalTrainingSafetyColumn"].Visible = true; 
                        break;
                    case "Essity1YBox":
                        dataGridView1.Columns["Essity1YearColumn"].Visible = true;
                        break; ;
                    case "FamiliarizedBox":
                        dataGridView1.Columns["FamiliarizedColumn"].Visible = true;
                        break;
                    case "FireDateBox":
                        dataGridView1.Columns["FireWorkingDateColumn"].Visible = true;
                        break;
                    case "FireNumberBox":
                        dataGridView1.Columns["FireWorkingNumberColumn"].Visible = true;
                        break;
                    case "FirstAidBox":
                        dataGridView1.Columns["FirstAidTrainingColumn"].Visible = true;
                        break;
                    case "LiveBox":
                        dataGridView1.Columns["LiveWorkingTrainingColumn"].Visible = true;
                        break;
                    case "NokianR1YBox":
                        dataGridView1.Columns["NokianRenkat1YearColumn"].Visible = true;
                        break;
                    case "NokianRLOTOBox":
                        dataGridView1.Columns["NokianRenkatLOTOColumn"].Visible = true;
                        break;
                    case "NvEBox":
                        dataGridView1.Columns["NvEColumn"].Visible = true;
                        break;
                    case "OtherBox":
                        dataGridView1.Columns["OtherColumn"].Visible = true;
                        break;
                    case "SandvikBox":
                        dataGridView1.Columns["SandvikColumn"].Visible = true;
                        break;
                    case "TampereenSBox":
                        dataGridView1.Columns["TampereenSahkolaitosColumn"].Visible = true;
                        break;
                    case "TaxBox":
                        dataGridView1.Columns["TaxNumberColumn"].Visible = true;
                        break;
                    case "ValttikorttiBox":
                        dataGridView1.Columns["ValttikorttiColumn"].Visible = true;
                        break;
                    case "WorkSafetyBox":
                        dataGridView1.Columns["WorkSafetyTrainingColumn"].Visible = true;
                        break;
                    case "WorkSafetyTrainingBox":
                        dataGridView1.Columns["WorkSafetyTrainingNumberColumn"].Visible = true;
                        break;
                }
            }
            if (!senderCheck.Checked)
            {
                switch (senderCheck.Name)
                {
                    case "ElectricalSafetyBox":
                        dataGridView1.Columns["ElectricalTrainingSafetyColumn"].Visible = false; 
                        break;
                    case "Essity1YBox":
                        dataGridView1.Columns["Essity1YearColumn"].Visible = false;
                        break; ;
                    case "FamiliarizedBox":
                        dataGridView1.Columns["FamiliarizedColumn"].Visible = false;
                        break;
                    case "FireDateBox":
                        dataGridView1.Columns["FireWorkingDateColumn"].Visible = false;
                        break;
                    case "FireNumberBox":
                        dataGridView1.Columns["FireWorkingNumberColumn"].Visible = false;
                        break;
                    case "FirstAidBox":
                        dataGridView1.Columns["FirstAidTrainingColumn"].Visible = false;
                        break;
                    case "LiveBox":
                        dataGridView1.Columns["LiveWorkingTrainingColumn"].Visible = false;
                        break;
                    case "NokianR1YBox":
                        dataGridView1.Columns["NokianRenkat1YearColumn"].Visible = false;
                        break;
                    case "NokianRLOTOBox":
                        dataGridView1.Columns["NokianRenkatLOTOColumn"].Visible = false;
                        break;
                    case "NvEBox":
                        dataGridView1.Columns["NvEColumn"].Visible = false;
                        break;
                    case "OtherBox":
                        dataGridView1.Columns["OtherColumn"].Visible = false;
                        break;
                    case "SandvikBox":
                        dataGridView1.Columns["SandvikColumn"].Visible = false;
                        break;
                    case "TampereenSBox":
                        dataGridView1.Columns["TampereenSahkolaitosColumn"].Visible = false;
                        break;
                    case "TaxBox":
                        dataGridView1.Columns["TaxNumberColumn"].Visible = false;
                        break;
                    case "ValttikorttiBox":
                        dataGridView1.Columns["ValttikorttiColumn"].Visible = false;
                        break;
                    case "WorkSafetyBox":
                        dataGridView1.Columns["WorkSafetyTrainingColumn"].Visible = false;
                        break;
                    case "WorkSafetyTrainingBox":
                        dataGridView1.Columns["WorkSafetyTrainingNumberColumn"].Visible = false;
                        break;
                }
            }
            dataGridView1.Update();
        }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // Clean up COM objects
            System.Runtime.InteropServices.Marshal.ReleaseComObject(currentRow);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(usedRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
        DateTime FireWorkingDate;
        public void CheckForDeadlines()
        {
            foreach (Employee employee in Employees)
            {
                if (employee.FireWorkingDate != "x" && employee.FireWorkingDate != "X" && employee.FireWorkingDate != "" && employee.FireWorkingDate != null && employee.FireWorkingDate != "f" && employee.FireWorkingDate != "s" && employee.EmployeeName == "Daniel Heikkila")
                {
                    if (DateTime.TryParseExact(employee.FireWorkingDate, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out FireWorkingDate))
                    {

                    }
                    //DateTime FireWorkingDate = DateTime.Parse(employee.FireWorkingDate.Replace(".", "/"));
                    DateTime currentDate = DateTime.Today;
                    TimeSpan difference = FireWorkingDate - currentDate;
                    int fireWorkingDaysDifference = (int)difference.TotalDays;
                    if (fireWorkingDaysDifference == 31 || fireWorkingDaysDifference == 30 || fireWorkingDaysDifference == 29 || fireWorkingDaysDifference == 28)
                    {
                        string email = employee.EmployeeName.Replace(" ", ".") + "@tp-kunnossapito.fi";
                        SendEmail(email);
                    }
                }
            }
        }
        public void SendEmail(string emailAddress) //REMEMBER TO USE THE INAPPPASSWORD LATER ON WHEN IT EXISTS
        {
            if (credentials.Login != "" && credentials.Password != "" && credentials.Client != "")
            {
                string toAddress = emailAddress;
                string subject = "Put subject here";
                string body = "Put text here";

                using (MailMessage mail = new MailMessage(credentials.Login, toAddress))
                {
                    mail.Subject = subject;
                    mail.Body = body;

                    using (SmtpClient smtp = new SmtpClient(credentials.Client, 587))
                    {
                        smtp.Credentials = new NetworkCredential(credentials.Login, credentials.Password);
                        smtp.EnableSsl = true;
                        //smtp.Send(mail);
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            fromAddress = emailTextBox.Text;
            fromPassword = passwordTextBox.Text;
            if (serviceTextBox.Text == "Outlook")
                smtpClient = "smtp.office365.com";
            else if (serviceTextBox.Text == "Gmail")
                smtpClient = "smtp.gmail.com";
            else if (serviceTextBox.Text =="Yahoo")
                smtpClient = "smtp.mail.yahoo.com";
            credentials.Login = fromAddress;
            credentials.Password = fromPassword;
            credentials.Client = smtpClient;
            if (twoFactorBox.Text != "")
            {
                credentials.InAppPassword = twoFactorBox.Text;
            }
            SaveSettings(credentials, ".\\" + "credentials.json");
            panel2.Visible = false;
        }
        public void SaveSettings(credentials credentials, string filename)
        {
            
            string json = JsonConvert.SerializeObject(credentials);
            File.WriteAllText(filename, json);
        }
        public credentials RestoreSettings(string filename)
        {
            if (File.Exists(filename))
            {
                string json = File.ReadAllText(filename);
                return JsonConvert.DeserializeObject<credentials>(json);
            }
            else
            {
                return new credentials();
            }
        }

        private void EmailButton_Click(object sender, EventArgs e)
        {
            panel2.Visible = true;
        }

        private void deleteButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell != null)//////////////////////////////////
            {
                isProgrammaticEdit = true;
                Employee employeeToDelete = null;
                foreach (Employee empl in Employees)
                {
                    if (empl.EmployeeName == selectedEmployee.EmployeeName && empl.Familiarized == selectedEmployee.Familiarized && empl.FireWorkingDate == selectedEmployee.FireWorkingDate &&
                                empl.FireWorkingNumber == selectedEmployee.FireWorkingNumber && empl.WorkSafetyTraining == selectedEmployee.WorkSafetyTraining && empl.WorkSafetyTrainingNumber == selectedEmployee.WorkSafetyTrainingNumber &&
                                empl.ElectricalSafetyTraining == selectedEmployee.ElectricalSafetyTraining && empl.LiveWorkingTraining == selectedEmployee.LiveWorkingTraining && empl.FirstAidTraining == selectedEmployee.FirstAidTraining &&
                                empl.Valttikortti == selectedEmployee.Valttikortti && empl.TaxNumber == selectedEmployee.TaxNumber && empl.NokianRenkat1Year == selectedEmployee.NokianRenkat1Year &&
                                empl.NokianRenkatLOTO == selectedEmployee.NokianRenkatLOTO && empl.Essity1Year == selectedEmployee.Essity1Year && empl.TampereenSahkolaitos == selectedEmployee.TampereenSahkolaitos &&
                                empl.NvE == selectedEmployee.NvE && empl.Sandvik == selectedEmployee.Sandvik && empl.Other == selectedEmployee.Other)
                    {
                        employeeToDelete = empl;
                        continue;
                    }
                }
                if (employeeToDelete != null)
                    Employees.Remove(employeeToDelete);

                try
                {
                    worksheet = workbook.Sheets[1];
                    usedRange = worksheet.UsedRange;
                    int rowCount = usedRange.Rows.Count;

                    int i = 4;
                    currentRow = usedRange.Rows[i];
                    try
                    {
                        while (currentRow.Cells[1, 1].Value2 != null)
                        {
                            Employee potentialEmployeeToDelete = new Employee();
                            int colCount = Math.Min(usedRange.Columns.Count, 18);

                            potentialEmployeeToDelete.EmployeeName = currentRow.Cells[1, 1].Value2.ToString();
                            if (currentRow.Cells[1, 2].Value2 != null)
                            {
                                if (currentRow.Cells[1, 2].Value2.ToString() == "OK")
                                    potentialEmployeeToDelete.Familiarized = true;
                                else
                                    potentialEmployeeToDelete.Familiarized = false;
                            }
                            if (currentRow.Cells[1, 3].Value2 != null)
                                potentialEmployeeToDelete.FireWorkingDate = convert(currentRow.Cells[1, 3].Value2.ToString());
                            if (currentRow.Cells[1, 4].Value2 != null)
                                potentialEmployeeToDelete.FireWorkingNumber = currentRow.Cells[1, 4].Value2.ToString();
                            if (currentRow.Cells[1, 5].Value2 != null)
                                potentialEmployeeToDelete.WorkSafetyTraining = convert(currentRow.Cells[1, 5].Value2.ToString());
                            if (currentRow.Cells[1, 6].Value2 != null)
                                potentialEmployeeToDelete.WorkSafetyTrainingNumber = currentRow.Cells[1, 6].Value2.ToString();
                            if (currentRow.Cells[1, 7].Value2 != null)
                                potentialEmployeeToDelete.ElectricalSafetyTraining = convert(currentRow.Cells[1, 7].Value2.ToString());
                            if (currentRow.Cells[1, 8].Value2 != null)
                                potentialEmployeeToDelete.LiveWorkingTraining = convert(currentRow.Cells[1, 8].Value2.ToString());
                            if (currentRow.Cells[1, 9].Value2 != null)
                                potentialEmployeeToDelete.FirstAidTraining = convert(currentRow.Cells[1, 9].Value2.ToString());
                            if (currentRow.Cells[1, 10].Value2 != null)
                                potentialEmployeeToDelete.Valttikortti = currentRow.Cells[1, 10].Value2.ToString();
                            if (currentRow.Cells[1, 11].Value2 != null)
                                potentialEmployeeToDelete.TaxNumber = currentRow.Cells[1, 11].Value2.ToString();
                            if (currentRow.Cells[1, 12].Value2 != null)
                                potentialEmployeeToDelete.NokianRenkat1Year = convert(currentRow.Cells[1, 12].Value2.ToString());
                            if (currentRow.Cells[1, 13].Value2 != null)
                                potentialEmployeeToDelete.NokianRenkatLOTO = convert(currentRow.Cells[1, 13].Value2.ToString());
                            if (currentRow.Cells[1, 14].Value2 != null)
                                potentialEmployeeToDelete.Essity1Year = convert(currentRow.Cells[1, 14].Value2.ToString());
                            if (currentRow.Cells[1, 15].Value2 != null)
                                potentialEmployeeToDelete.TampereenSahkolaitos = convert(currentRow.Cells[1, 15].Value2.ToString());
                            if (currentRow.Cells[1, 16].Value2 != null)
                                potentialEmployeeToDelete.NvE = convert(currentRow.Cells[1, 16].Value2.ToString());
                            if (currentRow.Cells[1, 17].Value2 != null)
                                potentialEmployeeToDelete.Sandvik = convert(currentRow.Cells[1, 17].Value2.ToString());
                            if (currentRow.Cells[1, 18].Value2 != null)
                                potentialEmployeeToDelete.Other = currentRow.Cells[1, 18].Value2.ToString();

                            if (potentialEmployeeToDelete.EmployeeName == selectedEmployee.EmployeeName && potentialEmployeeToDelete.Familiarized == selectedEmployee.Familiarized && potentialEmployeeToDelete.FireWorkingDate == selectedEmployee.FireWorkingDate &&
                                potentialEmployeeToDelete.FireWorkingNumber == selectedEmployee.FireWorkingNumber && potentialEmployeeToDelete.WorkSafetyTraining == selectedEmployee.WorkSafetyTraining && potentialEmployeeToDelete.WorkSafetyTrainingNumber == selectedEmployee.WorkSafetyTrainingNumber &&
                                potentialEmployeeToDelete.ElectricalSafetyTraining == selectedEmployee.ElectricalSafetyTraining && potentialEmployeeToDelete.LiveWorkingTraining == selectedEmployee.LiveWorkingTraining && potentialEmployeeToDelete.FirstAidTraining == selectedEmployee.FirstAidTraining &&
                                potentialEmployeeToDelete.Valttikortti == selectedEmployee.Valttikortti && potentialEmployeeToDelete.TaxNumber == selectedEmployee.TaxNumber && potentialEmployeeToDelete.NokianRenkat1Year == selectedEmployee.NokianRenkat1Year &&
                                potentialEmployeeToDelete.NokianRenkatLOTO == selectedEmployee.NokianRenkatLOTO && potentialEmployeeToDelete.Essity1Year == selectedEmployee.Essity1Year && potentialEmployeeToDelete.TampereenSahkolaitos == selectedEmployee.TampereenSahkolaitos &&
                                potentialEmployeeToDelete.NvE == selectedEmployee.NvE && potentialEmployeeToDelete.Sandvik == selectedEmployee.Sandvik && potentialEmployeeToDelete.Other == selectedEmployee.Other)
                            {
                                Excel.Range rowToDelete = worksheet.Rows[i];
                                rowToDelete.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rowToDelete);
                            }

                            i++;
                            currentRow = usedRange.Rows[i];
                        }
                    }
                    catch
                    {
                    }
                }
                catch
                {
                }

                workbook.Save();

                dataGridView1.Rows.Clear();
                foreach (Employee employee in Employees)
                {
                    dataGridView1.Rows.Add(employee.EmployeeName, employee.Familiarized, employee.FireWorkingDate, employee.FireWorkingNumber, employee.WorkSafetyTraining, employee.WorkSafetyTrainingNumber,
                                employee.ElectricalSafetyTraining, employee.LiveWorkingTraining, employee.FirstAidTraining, employee.Valttikortti, employee.TaxNumber, employee.NokianRenkat1Year, employee.NokianRenkatLOTO,
                                employee.Essity1Year, employee.TampereenSahkolaitos, employee.NvE, employee.Sandvik, employee.Other);
                }
                isProgrammaticEdit = false;
            }
        }

        private void rowSelectedUnselected(object sender, EventArgs e)
        {
            if (isProgrammaticEdit == false)
            {
                if (dataGridView1.Focused == true)
                {

                    selectedEmployee.EmployeeName = null;
                    selectedEmployee.Familiarized = false;
                    selectedEmployee.FireWorkingDate = null;
                    selectedEmployee.FireWorkingNumber = null;
                    selectedEmployee.WorkSafetyTraining = null;
                    selectedEmployee.WorkSafetyTrainingNumber = null;
                    selectedEmployee.ElectricalSafetyTraining = null;
                    selectedEmployee.LiveWorkingTraining = null;
                    selectedEmployee.FirstAidTraining = null;
                    selectedEmployee.Valttikortti = null;
                    selectedEmployee.TaxNumber = null;
                    selectedEmployee.NokianRenkat1Year = null;
                    selectedEmployee.NokianRenkatLOTO = null;
                    selectedEmployee.Essity1Year = null;
                    selectedEmployee.TampereenSahkolaitos = null;
                    selectedEmployee.NvE = null;
                    selectedEmployee.Sandvik = null;
                    selectedEmployee.Other = null;

                    try
                    {
                        selectedEmployee.EmployeeName = dataGridView1.CurrentRow.Cells["NameColumn"].Value.ToString();
                        if (dataGridView1.CurrentRow.Cells["FamiliarizedColumn"].Value.ToString() == "True")
                        {
                            selectedEmployee.Familiarized = true;
                        }
                        else
                        {
                            selectedEmployee.Familiarized = false;
                        }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["FireWorkingDateColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["FireWorkingDateColumn"].Value.ToString() != "")
                                selectedEmployee.FireWorkingDate = convert(dataGridView1.CurrentRow.Cells["FireWorkingDateColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["FireWorkingNumberColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["FireWorkingNumberColumn"].Value.ToString() != "")
                                selectedEmployee.FireWorkingNumber = dataGridView1.CurrentRow.Cells["FireWorkingNumberColumn"].Value.ToString();
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["WorkSafetyTrainingColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["WorkSafetyTrainingColumn"].Value.ToString() != "")
                                selectedEmployee.WorkSafetyTraining = convert(dataGridView1.CurrentRow.Cells["WorkSafetyTrainingColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["WorkSafetyTrainingNumberColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["WorkSafetyTrainingNumberColumn"].Value.ToString() != "")
                                selectedEmployee.WorkSafetyTrainingNumber = dataGridView1.CurrentRow.Cells["WorkSafetyTrainingNumberColumn"].Value.ToString();
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["ElectricalTrainingSafetyColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["ElectricalTrainingSafetyColumn"].Value.ToString() != "")
                                selectedEmployee.ElectricalSafetyTraining = convert(dataGridView1.CurrentRow.Cells["ElectricalTrainingSafetyColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["LiveWorkingTrainingColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["LiveWorkingTrainingColumn"].Value.ToString() != "")
                                selectedEmployee.LiveWorkingTraining = convert(dataGridView1.CurrentRow.Cells["LiveWorkingTrainingColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["FirstAidTrainingColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["FirstAidTrainingColumn"].Value.ToString() != "")
                                selectedEmployee.FirstAidTraining = convert(dataGridView1.CurrentRow.Cells["FirstAidTrainingColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["ValttikorttiColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["ValttikorttiColumn"].Value.ToString() != "")
                                selectedEmployee.Valttikortti = dataGridView1.CurrentRow.Cells["ValttikorttiColumn"].Value.ToString();
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["TaxNumberColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["TaxNumberColumn"].Value.ToString() != "")
                                selectedEmployee.TaxNumber = dataGridView1.CurrentRow.Cells["TaxNumberColumn"].Value.ToString();
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["NokianRenkat1YearColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["NokianRenkat1YearColumn"].Value.ToString() != "")
                                selectedEmployee.NokianRenkat1Year = convert(dataGridView1.CurrentRow.Cells["NokianRenkat1YearColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["NokianRenkatLOTOColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["NokianRenkatLOTOColumn"].Value.ToString() != "")
                                selectedEmployee.NokianRenkatLOTO = convert(dataGridView1.CurrentRow.Cells["NokianRenkatLOTOColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["Essity1YearColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["Essity1YearColumn"].Value.ToString() != "")
                                selectedEmployee.Essity1Year = convert(dataGridView1.CurrentRow.Cells["Essity1YearColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["TampereenSahkolaitosColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["TampereenSahkolaitosColumn"].Value.ToString() != "")
                                selectedEmployee.TampereenSahkolaitos = convert(dataGridView1.CurrentRow.Cells["TampereenSahkolaitosColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["NvEColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["NvEColumn"].Value.ToString() != "")
                                selectedEmployee.NvE = convert(dataGridView1.CurrentRow.Cells["NvEColumn"].Value.ToString());
                        }
                        catch { }
                        try
                        {
                            if (dataGridView1.CurrentRow.Cells["SandvikColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["SandvikColumn"].Value.ToString() != "")
                                selectedEmployee.Sandvik = convert(dataGridView1.CurrentRow.Cells["SandvikColumn"].Value.ToString());
                        }
                        catch { }
                        if (dataGridView1.CurrentRow.Cells["OtherColumn"].Value.ToString() != null && dataGridView1.CurrentRow.Cells["OtherColumn"].Value.ToString() != "")
                            selectedEmployee.Other = dataGridView1.CurrentRow.Cells["OtherColumn"].Value.ToString();
                    }
                    catch
                    {
                    }

                    try
                    {
                        if (selectedEmployee.EmployeeName != null && selectedEmployee.EmployeeName != "")
                        {
                            dataGridView1.DefaultCellStyle.SelectionBackColor = brighterColor;
                            dataGridView1.DefaultCellStyle.SelectionForeColor = dataGridView1.DefaultCellStyle.ForeColor;
                        }
                    }
                    catch
                    {
                        dataGridView1.DefaultCellStyle.SelectionBackColor = brighterColor;
                        dataGridView1.DefaultCellStyle.SelectionForeColor = dataGridView1.DefaultCellStyle.ForeColor;
                    }

                }
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (isProgrammaticEdit == false)
            {
                Employee employeeToEdit = null;
                string employeeName = null;
                string changedCellText = null;
                foreach (Employee empl in Employees)
                {
                    if (empl.EmployeeName == selectedEmployee.EmployeeName && empl.Familiarized == selectedEmployee.Familiarized && empl.FireWorkingDate == selectedEmployee.FireWorkingDate &&
                                empl.FireWorkingNumber == selectedEmployee.FireWorkingNumber && empl.WorkSafetyTraining == selectedEmployee.WorkSafetyTraining && empl.WorkSafetyTrainingNumber == selectedEmployee.WorkSafetyTrainingNumber &&
                                empl.ElectricalSafetyTraining == selectedEmployee.ElectricalSafetyTraining && empl.LiveWorkingTraining == selectedEmployee.LiveWorkingTraining && empl.FirstAidTraining == selectedEmployee.FirstAidTraining &&
                                empl.Valttikortti == selectedEmployee.Valttikortti && empl.TaxNumber == selectedEmployee.TaxNumber && empl.NokianRenkat1Year == selectedEmployee.NokianRenkat1Year &&
                                empl.NokianRenkatLOTO == selectedEmployee.NokianRenkatLOTO && empl.Essity1Year == selectedEmployee.Essity1Year && empl.TampereenSahkolaitos == selectedEmployee.TampereenSahkolaitos &&
                                empl.NvE == selectedEmployee.NvE && empl.Sandvik == selectedEmployee.Sandvik && empl.Other == selectedEmployee.Other)
                    {
                        employeeToEdit = empl;
                        employeeName = employeeToEdit.EmployeeName;
                        continue;
                    }
                }
                if (employeeToEdit != null)
                {
                    int columnIndex = dataGridView1.CurrentCell.ColumnIndex;//Delete row in excel, append the object and keep the datagridview
                    string columnName = dataGridView1.Columns[columnIndex].Name;
                    int rowIndex = dataGridView1.CurrentCell.RowIndex;
                    changedCellText = dataGridView1.Rows[rowIndex].Cells[columnIndex].Value?.ToString();
                    switch (columnName)
                    {
                        case "FamiliarizedColumn":
                            if (dataGridView1.Rows[rowIndex].Cells[columnIndex].Value?.ToString() == "OK")
                                employeeToEdit.Familiarized = true;
                            else
                                employeeToEdit.Familiarized = false;
                            break;
                        case "FireWorkingDateColumn":
                            employeeToEdit.FireWorkingDate = changedCellText;
                            break;
                        case "FireWorkingNumberColumn":
                            employeeToEdit.FireWorkingNumber = changedCellText;
                            break;
                        case "WorkSafetyTrainingColumn":
                            employeeToEdit.WorkSafetyTraining = changedCellText;
                            break;
                        case "WorkSafetyTrainingNumberColumn":
                            employeeToEdit.WorkSafetyTrainingNumber = changedCellText;
                            break;
                        case "ElectricalSafetyTrainingColumn":
                            employeeToEdit.ElectricalSafetyTraining = changedCellText;
                            break;
                        case "LiveWorkingTrainingColumn":
                            employeeToEdit.LiveWorkingTraining = changedCellText;
                            break;
                        case "FirstAidTrainingColumn":
                            employeeToEdit.FirstAidTraining = changedCellText;
                            break;
                        case "ValttikorttiColumn":
                            employeeToEdit.Valttikortti = changedCellText;
                            break;
                        case "TaxNumberColumn":
                            employeeToEdit.TaxNumber = changedCellText;
                            break;
                        case "NokianRenkat1YearColumn":
                            employeeToEdit.NokianRenkat1Year = changedCellText;
                            break;
                        case "NokianRenkatLOTOColumn":
                            employeeToEdit.NokianRenkatLOTO = changedCellText;
                            break;
                        case "Essity1YearColumn":
                            employeeToEdit.Essity1Year = changedCellText;
                            break;
                        case "TampereenSahkolaitosColumn":
                            employeeToEdit.TampereenSahkolaitos = changedCellText;
                            break;
                        case "NvEColumn":
                            employeeToEdit.NvE = changedCellText;
                            break;
                        case "SandvikColumn":
                            employeeToEdit.Sandvik = changedCellText;
                            break;
                        case "OtherColumn":
                            employeeToEdit.Other = changedCellText;
                            break;
                    }

                    // Open the Excel workbook
                    workbook = excelApp.Workbooks.Open(@"PATH HERE");
                    try
                    {
                        worksheet = workbook.Sheets[1];
                        usedRange = worksheet.UsedRange;
                        int rowCount = usedRange.Rows.Count;

                        if (columnName != "FamiliarizedColumn" && columnName != "FireWorkingNumberColumn" && columnName != "WorkSafetyTrainingNumberColumn" && columnName != "ValttikorttiColumn" && columnName != "TaxNumberColumn" && columnName != "OtherColumn")
                        {
                            usedRange.Cells[rowIndex + 4, columnIndex + 1].Value2 = convert(changedCellText);
                        }
                        else
                        {
                            usedRange.Cells[rowIndex + 4, columnIndex + 1].Value2 = changedCellText;
                        }
                    }
                    catch
                    {
                    }

                    workbook.Save();
                }
            }
        }

        private void Control_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return))
            {
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
        }
    }
    public class Employee
    {
        public string EmployeeName { get; set; }
        public bool Familiarized { get; set; }
        public string FireWorkingDate { get; set; }
        public string FireWorkingNumber { get; set; }
        public string WorkSafetyTraining { get; set; }
        public string WorkSafetyTrainingNumber { get; set; }
        public string ElectricalSafetyTraining { get; set; }
        public string LiveWorkingTraining { get; set; }
        public string FirstAidTraining { get; set; }
        public string Valttikortti { get; set; }
        public string TaxNumber { get; set; }
        public string NokianRenkat1Year { get; set; }
        public string NokianRenkatLOTO { get; set; }
        public string Essity1Year { get; set; }
        public string TampereenSahkolaitos { get; set; }
        public string NvE { get; set; }
        public string Sandvik { get; set; }
        public string Other { get; set; }
    }
    public class credentials
    {
        public string Login { get; set; }
        public string Password { get; set; }
        public string Client { get; set; }
        public string InAppPassword { get; set; }
    }
}
