using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CheckBox = System.Windows.Forms.CheckBox;
using Excel = Microsoft.Office.Interop.Excel;
using Point = System.Drawing.Point;

namespace Schedule_project
{
    public partial class MainForm : Form
    {
        #region Parameters

        Thread t;
        List<Employee> Employees = new List<Employee>();
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
        private Stack<CellChange> undoStack = new Stack<CellChange>();
        DateTime FireWorkingDate;
        string cellOldValue; //Old value of the cell is later sent into the undoManager
        List<string> columnsToFormat = new List<string>() { "FireWorkingDateColumn", "WorkSafetyTrainingColumn", "ElectricalTrainingSafetyColumn", "LiveWorkingTrainingColumn", "FirstAidTrainingColumn",
        "NokianRenkat1YearColumn", "NokianRenkatLOTOColumn", "Essity1YearColumn", "TampereenSahkolaitosColumn", "NvEColumn", "SandvikColumn"};//These will be turned into european format dates
        EmployeeBoxes employeeBoxesChecked = new EmployeeBoxes();
        Color brighterColor = Color.LightSkyBlue;

        #endregion

        #region Program Start

        public MainForm()
        {
            isProgrammaticEdit = true;
            InitializeComponent();
            programStart();
            t = new Thread(new ThreadStart(ReadTheExcel));
            t.Start();
            MainForm_Resize();//This first resize makes sure all the components are in correct spots when the program starts //Line 78
            ReadTheExcel();
            foreach (string columnToFormat in columnsToFormat)
            {
                try
                {
                    dataGridView1.Columns[columnToFormat].DefaultCellStyle.Format = "dd.MM.yyyy";
                }
                catch
                {
                }
            }
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
            CheckForDeadlines("Both");//Edit later to accept a parameter that selects whether the email or colouring is to be performed
            credentials = RestoreSettings(".\\" + "credentials.json");
            employeeBoxesChecked = RestoreSettings(".\\" + "boxes.json", 1);
            isProgrammaticEdit = false;
            dataGridView1.DefaultCellStyle.SelectionBackColor = dataGridView1.DefaultCellStyle.BackColor;
            dataGridView1.DefaultCellStyle.SelectionForeColor = dataGridView1.DefaultCellStyle.ForeColor;
            LanguageSelectionBox.SelectedIndex = 1;
            this.Visible = false;
        }

        private void programStart()//This function is only called when the programs starts
        {
            excelApp = new Excel.Application();
            //And it opens the Excel workbook
            workbook = excelApp.Workbooks.Open(@"C:\\Users\\danie.DANIELS_PC\\OneDrive\\Desktop\\TYÖNTEKIJÄREKISTERI JTOL.xlsx");
            worksheet = workbook.Sheets[1];
            usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
        }

        #endregion

        #region Visuals

        private void MainForm_Resize(object sender, EventArgs e)
        {
            //Adjusting internal controls based on form's new size
            Resizing();
        }

        private async void MainForm_Resize()
        {
            //Adjusting internal controls based on form's new size
            await Task.Run(() =>
            {
                Resizing();
            });
        }

        private void Resizing()
        {
            //Adjusting internal controls based on form's new size
            if (this.InvokeRequired)
            {
                // If not on the UI thread, invoke the method on the UI thread using BeginInvoke
                this.BeginInvoke(new MethodInvoker(Resizing));
                return;
            }
            panel1.Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 35);
            panel1.Location = new Point(10, 20);
            deleteButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.03));
            EmailButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.07));
            AddEmployeeButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.11));
            CheckAllBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.15));
            FamiliarizedBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.19));
            FireDateBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.23));
            FireNumberBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.27));
            WorkSafetyBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.31));
            WorkSafetyTrainingBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.35));
            ElectricalSafetyBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.39));
            LiveBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.43));
            FirstAidBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.47));
            ValttikorttiBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.51));
            TaxBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.55));
            NokianR1YBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.59));
            NokianRLOTOBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.63));
            Essity1YBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.67));
            TampereenSBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.71));
            NvEBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.75));
            SandvikBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.79));
            OtherBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.83));
            panel2.Location = new Point((int)(this.referenceLabel.Location.X * 0.2), (int)(this.referenceLabel.Location.Y * 0.1));
            panel2.Size = new Size((int)(this.panel1.Width * 0.75), (int)(this.panel1.Height * 0.85));
            dataGridView1.Location = new Point((int)(this.referenceLabel.Location.X * 0.2), (int)(this.referenceLabel.Location.Y * 0.1));
            dataGridView1.Size = new Size((int)(this.panel1.Width * 0.75), (int)(this.panel1.Height * 0.75));
            EmailConfirmButton.Size = new Size((int)(this.panel1.Width * 0.184), (int)(this.panel1.Height * 0.069));
            EmailConfirmButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.55), (int)(this.referenceLabel.Location.Y * 0.65));
            emailTextBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.05), (int)(this.referenceLabel.Location.Y * 0.1));
            passwordTextBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.05), (int)(this.referenceLabel.Location.Y * 0.25));
            serviceTextBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.05), (int)(this.referenceLabel.Location.Y * 0.40));
            twoFactorBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.05), (int)(this.referenceLabel.Location.Y * 0.55));
            EmailLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.05));
            PasswordLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.2));
            ServiceLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.30));
            label4.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.35));
            TwoFactorLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.02), (int)(this.referenceLabel.Location.Y * 0.50));
            UndoButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.2), (int)(this.referenceLabel.Location.Y * 0.05));
            LanguageSelectionBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.9), (int)(this.referenceLabel.Location.Y * 0.05));
            BackButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.65));

            IEnumerable<System.Windows.Forms.Button> buttonControls = this.Controls.OfType<System.Windows.Forms.Button>();
            foreach (System.Windows.Forms.Button button in buttonControls)
            {
                button.Size = new Size((int)(this.panel1.Width * 0.129), (int)(this.panel1.Height * 0.045));
            }

            if (panel1.Size.Width > 800)
            {
                emailTextBox.Size = new Size((int)(this.panel1.Width * 0.4), 26);
                passwordTextBox.Size = new Size((int)(this.panel1.Width * 0.4), 26);
                serviceTextBox.Size = new Size((int)(this.panel1.Width * 0.4), 26);
                twoFactorBox.Size = new Size((int)(this.panel1.Width * 0.4), 26);
            }
            else
            {
                emailTextBox.Size = new Size(320, 26);
                passwordTextBox.Size = new Size(320, 26);
                serviceTextBox.Size = new Size(320, 26);
                twoFactorBox.Size = new Size(320, 26);
            }
        }

        #endregion

        private void Add_Click(object sender, EventArgs e)//This function is for the add new employee button
        {
            isProgrammaticEdit = true;
            DialogResult resultChoice = DialogResult.Cancel;
            NewEmployeeForm MChoiceForm = new NewEmployeeForm();//Opens a new form for new employee object creation
            resultChoice = MChoiceForm.ShowDialog(this);

            //If successfully created a new employee object the following section runs
            if (resultChoice == DialogResult.OK)
            {
                Employee newEmployee = MChoiceForm.sendNewEmployee();//This retrieves the new employee from the newEmployeeForm "MChoiceForm"
                Employees.Add(newEmployee);//adds it to the list of employees

                //Then it adds the new employee to the excel
                int currentLine = Employees.Count + 1;
                usedRange.Cells[currentLine + 2, 1].Value2 = newEmployee.EmployeeName;
                if (newEmployee.Familiarized == true)
                    usedRange.Cells[currentLine + 2, 2].Value2 = "OK";
                else
                    usedRange.Cells[currentLine + 2, 2].Value2 = "";
                usedRange.Cells[currentLine + 2, 3].Value2 = convert(newEmployee.FireWorkingDate);
                usedRange.Cells[currentLine + 2, 4].Value2 = newEmployee.FireWorkingNumber;
                usedRange.Cells[currentLine + 2, 5].Value2 = convert(newEmployee.WorkSafetyTraining);
                usedRange.Cells[currentLine + 2, 6].Value2 = newEmployee.WorkSafetyTrainingNumber;
                usedRange.Cells[currentLine + 2, 7].Value2 = convert(newEmployee.ElectricalSafetyTraining);
                usedRange.Cells[currentLine + 2, 8].Value2 = convert(newEmployee.LiveWorkingTraining);
                usedRange.Cells[currentLine + 2, 9].Value2 = convert(newEmployee.FirstAidTraining);
                usedRange.Cells[currentLine + 2, 10].Value2 = newEmployee.Valttikortti;
                usedRange.Cells[currentLine + 2, 11].Value2 = newEmployee.TaxNumber;
                usedRange.Cells[currentLine + 2, 12].Value2 = convert(newEmployee.NokianRenkat1Year);
                usedRange.Cells[currentLine + 2, 13].Value2 = convert(newEmployee.NokianRenkatLOTO);
                usedRange.Cells[currentLine + 2, 14].Value2 = convert(newEmployee.Essity1Year);
                usedRange.Cells[currentLine + 2, 15].Value2 = convert(newEmployee.TampereenSahkolaitos);
                usedRange.Cells[currentLine + 2, 16].Value2 = convert(newEmployee.NvE);
                usedRange.Cells[currentLine + 2, 17].Value2 = convert(newEmployee.Sandvik);
                usedRange.Cells[currentLine + 2, 18].Value2 = newEmployee.Other;

                Excel.Range newRowRange = usedRange.Rows[currentLine + 2];

                //Sets borders for the entire row
                newRowRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                newRowRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                workbook.Save();

                //And adds the new row in the dataGridView1
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

        public async void ReadTheExcel()//This reads the excel
        {
            await Task.Run(() =>
            {
                Read();
            });
            try 
            {
                //And adds the employee to the datagridView1
                foreach (Employee employee in Employees)
                {
                    dataGridView1.Rows.Add(employee.EmployeeName, employee.Familiarized, employee.FireWorkingDate, employee.FireWorkingNumber, employee.WorkSafetyTraining, employee.WorkSafetyTrainingNumber,
                                employee.ElectricalSafetyTraining, employee.LiveWorkingTraining, employee.FirstAidTraining, employee.Valttikortti, employee.TaxNumber, employee.NokianRenkat1Year, employee.NokianRenkatLOTO,
                                employee.Essity1Year, employee.TampereenSahkolaitos, employee.NvE, employee.Sandvik, employee.Other);
                }
                CheckForDeadlines("Colour");
            }
            catch 
            { 
            }
        }

        public void Read()
        {
            isProgrammaticEdit = true;
            try
            {
                int rowCount = usedRange.Rows.Count;

                int i = 4;
                currentRow = usedRange.Rows[i];
                try
                {
                    while (currentRow.Cells[1, 1].Value2 != null) // Checks if the first cell in the row is empty
                    {
                        Employee employee = new Employee();//Creates new instance of employee and assigns values to it
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
                        Employees.Add(employee);//Adds new employee to list
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

        private void Check_Uncheck(object sender, EventArgs e)//This triggers when you check/uncheck any of the boxes on the main form
        {
            CheckBox senderCheck = sender as CheckBox;
            if (senderCheck.Checked)
            {//This switch adds the checked parameter column to the dataGridView1
                switch (senderCheck.Name)
                {
                    case "CheckAllBox":
                        foreach (Control control in panel1.Controls)
                        {
                            if (control is CheckBox checkBox)
                            {
                                checkBox.Checked = true;
                            }
                        }
                        break;
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
            {//This switch removes the checked parameter column from the dataGridView1
                switch (senderCheck.Name)
                {
                    case "CheckAllBox":
                        foreach (Control control in panel1.Controls)
                        {
                            if (control is CheckBox checkBox)
                            {
                                checkBox.Checked = false;
                            }
                        }
                        break;
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
        }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)//This function triggers after the "close" button is pressed before the program is closed
        {//It saves and closes everything related to the excel
            employeeBoxesChecked = new EmployeeBoxes();
            foreach (Control control in panel1.Controls)
            {
                if (control is CheckBox checkBox)
                {
                    CheckBox controlBox = control as CheckBox;
                    if (controlBox.Checked)
                    {
                        switch (controlBox.Name)
                        {
                            case "ElectricalSafetyBox":
                                employeeBoxesChecked.ElectricalSafetyBox = true;
                                break;
                            case "Essity1YBox":
                                employeeBoxesChecked.Essity1YBox = true;
                                break; ;
                            case "FamiliarizedBox":
                                employeeBoxesChecked.FamiliarizedBox = true;
                                break;
                            case "FireDateBox":
                                employeeBoxesChecked.FireDateBox = true;
                                break;
                            case "FireNumberBox":
                                employeeBoxesChecked.FireNumberBox = true;
                                break;
                            case "FirstAidBox":
                                employeeBoxesChecked.FirstAidBox = true;
                                break;
                            case "LiveBox":
                                employeeBoxesChecked.LiveBox = true;
                                break;
                            case "NokianR1YBox":
                                employeeBoxesChecked.NokianR1YBox = true;
                                break;
                            case "NokianRLOTOBox":
                                employeeBoxesChecked.NokianRLOTOBox = true;
                                break;
                            case "NvEBox":
                                employeeBoxesChecked.NvEBox = true;
                                break;
                            case "OtherBox":
                                employeeBoxesChecked.OtherBox = true;
                                break;
                            case "SandvikBox":
                                employeeBoxesChecked.SandvikBox = true;
                                break;
                            case "TampereenSBox":
                                employeeBoxesChecked.TampereenSBox = true;
                                break;
                            case "TaxBox":
                                employeeBoxesChecked.TaxBox = true;
                                break;
                            case "ValttikorttiBox":
                                employeeBoxesChecked.ValttikorttiBox = true;
                                break;
                            case "WorkSafetyBox":
                                employeeBoxesChecked.WorkSafetyBox = true;
                                break;
                            case "WorkSafetyTrainingBox":
                                employeeBoxesChecked.WorkSafetyTrainingBox = true;
                                break;
                        }
                    }
                    else
                    {
                        switch (controlBox.Name)
                        {
                            case "ElectricalSafetyBox":
                                employeeBoxesChecked.ElectricalSafetyBox = false;
                                break;
                            case "Essity1YBox":
                                employeeBoxesChecked.Essity1YBox = false;
                                break; ;
                            case "FamiliarizedBox":
                                employeeBoxesChecked.FamiliarizedBox = false;
                                break;
                            case "FireDateBox":
                                employeeBoxesChecked.FireDateBox = false;
                                break;
                            case "FireNumberBox":
                                employeeBoxesChecked.FireNumberBox = false;
                                break;
                            case "FirstAidBox":
                                employeeBoxesChecked.FirstAidBox = false;
                                break;
                            case "LiveBox":
                                employeeBoxesChecked.LiveBox = false;
                                break;
                            case "NokianR1YBox":
                                employeeBoxesChecked.NokianR1YBox = false;
                                break;
                            case "NokianRLOTOBox":
                                employeeBoxesChecked.NokianRLOTOBox = false;
                                break;
                            case "NvEBox":
                                employeeBoxesChecked.NvEBox = false;
                                break;
                            case "OtherBox":
                                employeeBoxesChecked.OtherBox = false;
                                break;
                            case "SandvikBox":
                                employeeBoxesChecked.SandvikBox = false;
                                break;
                            case "TampereenSBox":
                                employeeBoxesChecked.TampereenSBox = false;
                                break;
                            case "TaxBox":
                                employeeBoxesChecked.TaxBox = false;
                                break;
                            case "ValttikorttiBox":
                                employeeBoxesChecked.ValttikorttiBox = false;
                                break;
                            case "WorkSafetyBox":
                                employeeBoxesChecked.WorkSafetyBox = false;
                                break;
                            case "WorkSafetyTrainingBox":
                                employeeBoxesChecked.WorkSafetyTrainingBox = false;
                                break;
                        }
                    }
                }
            }
            SaveSettings(employeeBoxesChecked, ".\\" + "boxes.json");
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

            //Cleans up COM objects
            System.Runtime.InteropServices.Marshal.ReleaseComObject(currentRow);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(usedRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
        public void CheckForDeadlines(string option)//This function checks for approaching deadlines for renewing the employee training//Option can be Email, Colour or Both
        {
            foreach (Employee employee in Employees)//This sends reminders to the employees whose deadline is nearing
            {
                foreach (PropertyInfo propertyInfo in employee.GetType().GetProperties())
                {
                    object propertyValue = propertyInfo.GetValue(employee);
                    if (option == "Colour" || option == "Both")
                        if (propertyValue is string stringValue && DateTime.TryParse(stringValue, out DateTime date))
                        {

                            DateTime currentDate = DateTime.Today;
                            TimeSpan difference = FireWorkingDate - currentDate;
                            int fireWorkingDaysDifference = (int)difference.TotalDays;
                            if (fireWorkingDaysDifference == 31 || fireWorkingDaysDifference == 30 || fireWorkingDaysDifference == 29 || fireWorkingDaysDifference == 28)
                            {
                                //string email = employee.EmployeeName.Replace(" ", ".") + "@tp-kunnossapito.fi";
                                //SendEmail(email);
                            }
                        }
                }
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)//This colours the dates according to time left until the deadline
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    try
                    {
                        if (option == "Colour" && cell.Value != null || option == "Both" && cell.Value != null)
                            if (DateTime.TryParse(cell.Value.ToString(), out DateTime date))//Here it checks if the date is under 1 month/1 week and colours it the cell accordingly
                            {
                                DateTime currentDate = DateTime.Today;
                                TimeSpan difference = date - currentDate;
                                int DaysDifference = (int)difference.TotalDays;

                                if (DaysDifference > 31)
                                {
                                    cell.Style.BackColor = Color.LightGreen;
                                }
                                else if (DaysDifference < 31 && DaysDifference > 7)
                                {
                                    cell.Style.BackColor = Color.Khaki;
                                }
                                else if (DaysDifference < 8)
                                {
                                    cell.Style.BackColor = Color.Salmon;
                                }
                            }
                    }
                    catch
                    {
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

        public void SaveSettings(credentials credentials, string filename)//This is the saving part of the memory of email credentials
        {
            string json = JsonConvert.SerializeObject(credentials);
            File.WriteAllText(filename, json);
        }
        public credentials RestoreSettings(string filename)//This is the retrieving part of the memory of the email credentials
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

        public void SaveSettings(EmployeeBoxes employeeBoxes, string filename)//This is the saving part of the memory of email credentials
        {
            string json = JsonConvert.SerializeObject(employeeBoxes);
            File.WriteAllText(filename, json);
        }
        public EmployeeBoxes RestoreSettings(string filename, int x)//This is the retrieving part of the memory of the email credentials
        {
            if (File.Exists(filename))
            {
                string json = File.ReadAllText(filename);
                employeeBoxesChecked = JsonConvert.DeserializeObject<EmployeeBoxes>(json);
                foreach (Control control in panel1.Controls)
                {
                    if (control is CheckBox checkBox && json != "null")
                    {
                        CheckBox controlBox = control as CheckBox;
                        switch (controlBox.Name)
                        {
                            case "ElectricalSafetyBox":
                                controlBox.Checked = employeeBoxesChecked.ElectricalSafetyBox;
                                break;
                            case "Essity1YBox":
                                controlBox.Checked = employeeBoxesChecked.Essity1YBox;
                                break; ;
                            case "FamiliarizedBox":
                                controlBox.Checked = employeeBoxesChecked.FamiliarizedBox;
                                break;
                            case "FireDateBox":
                                controlBox.Checked = employeeBoxesChecked.FireDateBox;
                                break;
                            case "FireNumberBox":
                                controlBox.Checked = employeeBoxesChecked.FireNumberBox;
                                break;
                            case "FirstAidBox":
                                controlBox.Checked = employeeBoxesChecked.FirstAidBox;
                                break;
                            case "LiveBox":
                                controlBox.Checked = employeeBoxesChecked.LiveBox;
                                break;
                            case "NokianR1YBox":
                                controlBox.Checked = employeeBoxesChecked.NokianR1YBox;
                                break;
                            case "NokianRLOTOBox":
                                controlBox.Checked = employeeBoxesChecked.NokianRLOTOBox;
                                break;
                            case "NvEBox":
                                controlBox.Checked = employeeBoxesChecked.NvEBox;
                                break;
                            case "OtherBox":
                                controlBox.Checked = employeeBoxesChecked.OtherBox;
                                break;
                            case "SandvikBox":
                                controlBox.Checked = employeeBoxesChecked.SandvikBox;
                                break;
                            case "TampereenSBox":
                                controlBox.Checked = employeeBoxesChecked.TampereenSBox;
                                break;
                            case "TaxBox":
                                controlBox.Checked = employeeBoxesChecked.TaxBox;
                                break;
                            case "ValttikorttiBox":
                                controlBox.Checked = employeeBoxesChecked.ValttikorttiBox;
                                break;
                            case "WorkSafetyBox":
                                controlBox.Checked = employeeBoxesChecked.WorkSafetyBox;
                                break;
                            case "WorkSafetyTrainingBox":
                                controlBox.Checked = employeeBoxesChecked.WorkSafetyTrainingBox;
                                break;
                        }
                    }
                }
                return employeeBoxesChecked;
            }
            else
            {
                return new EmployeeBoxes();
            }
        }

        private void EmailButton_Click(object sender, EventArgs e)//This makes the email credentials panel visible when tehe EmailButton is pressed
        {
            panel2.Visible = true;
        }

        private async void deleteButton_Click(object sender, EventArgs e)//This is activated when the delete button is pressed
        {
            if (dataGridView1.CurrentCell != null)
            {
                //Adjusting internal controls based on form's new size
                await Task.Run(() =>
                {
                    deletion();
                });

                dataGridView1.Rows.Clear();//Then it clears the dataGridView1
                foreach (Employee employee in Employees)
                {//And writes the updated version
                    dataGridView1.Rows.Add(employee.EmployeeName, employee.Familiarized, employee.FireWorkingDate, employee.FireWorkingNumber, employee.WorkSafetyTraining, employee.WorkSafetyTrainingNumber,
                                employee.ElectricalSafetyTraining, employee.LiveWorkingTraining, employee.FirstAidTraining, employee.Valttikortti, employee.TaxNumber, employee.NokianRenkat1Year, employee.NokianRenkatLOTO,
                                employee.Essity1Year, employee.TampereenSahkolaitos, employee.NvE, employee.Sandvik, employee.Other);
                }
                isProgrammaticEdit = false;
                CheckForDeadlines("Colour");
            }
        }
        public void deletion()
        {
            isProgrammaticEdit = true;
            Employee employeeToDelete = null;
            foreach (Employee empl in Employees)
            {//This finds the employee object that correlates with the selected row in datagridview1
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
                Employees.Remove(employeeToDelete);//Removes the selected employee from the list

            //Then it records the change for the "Undo" functionality with the Undo, CanUndo and RecordChange functions
            try
            {
                RecordChange(-99, -99, null, null, employeeToDelete);
            }
            catch
            {
            }

            try
            {
                worksheet = workbook.Sheets[1];
                usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;

                int i = 4;
                currentRow = usedRange.Rows[i];
                try
                {
                    while (currentRow.Cells[1, 1].Value2 != null)//This finds the selected employee in the excel file
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
                        {//When the correct employee is found
                            Excel.Range rowToDelete = worksheet.Rows[i];
                            rowToDelete.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);//The excel row is deleted with the rest of the list shifting up
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
        }

        private void rowSelectedUnselected()//This forgets previously selected row and reads the new one
        {
            if (isProgrammaticEdit == false)
            {
                if (dataGridView1.Focused == true)
                {//It creates assigns values from the selected row to a selectedEmployee object

                    //First it clears previous parameters
                    cellOldValue = null;
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
                    {//And then it assigns the new parameters
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
                    {//This is just visual editing
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

        private void writeNewLineExcel(Employee employeeToEdit, string changedCellText)//This function gets called after a cell's value gets changed and it writes the new value to the excel file
        {
            if (employeeToEdit != null)
            {
                int columnIndex = dataGridView1.CurrentCell.ColumnIndex;
                string columnName = dataGridView1.Columns[columnIndex].Name;
                int rowIndex = dataGridView1.CurrentCell.RowIndex;
                changedCellText = dataGridView1.Rows[rowIndex].Cells[columnIndex].Value?.ToString();

                switch (columnName)
                {//Here it edits the parameter that was changed in the cell to the employee object it correlates to
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

                //Opens the Excel workbook
                workbook = excelApp.Workbooks.Open(@"C:\Users\danie.DANIELS_PC\OneDrive\Desktop\TYÖNTEKIJÄREKISTERI JTOL.xlsx");
                try
                {
                    worksheet = workbook.Sheets[1];
                    usedRange = worksheet.UsedRange;
                    int rowCount = usedRange.Rows.Count;

                    //And writes the new value to excel
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

                //This is just formatting
                foreach (string columnToFormat in columnsToFormat)
                {
                    try
                    {
                        dataGridView1.Columns[columnToFormat].DefaultCellStyle.Format = "dd.MM.yyyy";
                    }
                    catch
                    {
                    }
                }

                workbook.Save();
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)//This activates on end of cell edit
        {
            if (isProgrammaticEdit == false)
            {//Here it determines the employee whose cell was edited and assigns it to the employeeToEdit object
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

                //Then it passes the employee to edit and the new cell value to writeNewLineExcel function
                writeNewLineExcel(employeeToEdit, changedCellText);

                //Then it records the change for the "Undo" functionality with the Undo, CanUndo and RecordChange functions
                try
                {
                    RecordChange(dataGridView1.CurrentCell.RowIndex, dataGridView1.CurrentCell.ColumnIndex, cellOldValue, dataGridView1.CurrentCell.Value.ToString(), null);
                }
                catch
                {
                }

            }
        }

        private void Control_KeyUp(object sender, KeyEventArgs e)//This just allows the user to navigate textboxes with the press of "Enter"
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return))
            {
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)//This triggers on begin of cell edit
        {
            rowSelectedUnselected(sender, e);

            //It assigns the current (old) cell value to the cellOldValue 
            try
            {
                cellOldValue = dataGridView1.CurrentCell.Value?.ToString();
            }
            catch
            {
            }
        }

        public void RecordChange(int rowIndex, int columnIndex, object oldValue, object newValue, Employee employee)//This records the change for the "Undo" functionality with the Undo, CanUndo and RecordChange functions
        {
            if (employee == null)
            {
                undoStack.Push(new CellChange(rowIndex, columnIndex, oldValue, newValue, null));
            }
            else if (rowIndex == -99 && columnIndex == -99 && oldValue == null && newValue == null && employee != null)
            {
                undoStack.Push(new CellChange(-99, -99, null, null, employee));
            }
        }

        public bool CanUndo()//This checks if the undoStack is empty or not
        {
            return undoStack.Count > 0;
        }

        public void Undo()//This undoes an action and deletes it from the undo stack (memory) 
        {
            isProgrammaticEdit = true;
            if (undoStack.Count > 0)
            {
                CellChange change = undoStack.Pop();
                if (change.Employee == null)
                {
                    dataGridView1.Rows[change.RowIndex].Cells[change.ColumnIndex].Value = change.OldValue;
                    dataGridView1.CurrentCell = dataGridView1.Rows[change.RowIndex].Cells[change.ColumnIndex];
                    Employee employeeToEdit = new Employee();
                    employeeToEdit.EmployeeName = dataGridView1.Rows[change.RowIndex].Cells["NameColumn"].ToString();

                    //And records the undo to the Excel file
                    writeNewLineExcel(employeeToEdit, change.OldValue.ToString());
                }
                else
                {
                    Employee newEmployee = change.Employee;
                    Employees.Add(newEmployee);//adds it to the list of employees

                    //Then it adds the new employee to the excel
                    int currentLine = Employees.Count + 1;
                    usedRange.Cells[currentLine + 2, 1].Value2 = newEmployee.EmployeeName;
                    if (newEmployee.Familiarized == true)
                        usedRange.Cells[currentLine + 2, 2].Value2 = "OK";
                    else
                        usedRange.Cells[currentLine + 2, 2].Value2 = "";
                    usedRange.Cells[currentLine + 2, 3].Value2 = convert(newEmployee.FireWorkingDate);
                    usedRange.Cells[currentLine + 2, 4].Value2 = newEmployee.FireWorkingNumber;
                    usedRange.Cells[currentLine + 2, 5].Value2 = convert(newEmployee.WorkSafetyTraining);
                    usedRange.Cells[currentLine + 2, 6].Value2 = newEmployee.WorkSafetyTrainingNumber;
                    usedRange.Cells[currentLine + 2, 7].Value2 = convert(newEmployee.ElectricalSafetyTraining);
                    usedRange.Cells[currentLine + 2, 8].Value2 = convert(newEmployee.LiveWorkingTraining);
                    usedRange.Cells[currentLine + 2, 9].Value2 = convert(newEmployee.FirstAidTraining);
                    usedRange.Cells[currentLine + 2, 10].Value2 = newEmployee.Valttikortti;
                    usedRange.Cells[currentLine + 2, 11].Value2 = newEmployee.TaxNumber;
                    usedRange.Cells[currentLine + 2, 12].Value2 = convert(newEmployee.NokianRenkat1Year);
                    usedRange.Cells[currentLine + 2, 13].Value2 = convert(newEmployee.NokianRenkatLOTO);
                    usedRange.Cells[currentLine + 2, 14].Value2 = convert(newEmployee.Essity1Year);
                    usedRange.Cells[currentLine + 2, 15].Value2 = convert(newEmployee.TampereenSahkolaitos);
                    usedRange.Cells[currentLine + 2, 16].Value2 = convert(newEmployee.NvE);
                    usedRange.Cells[currentLine + 2, 17].Value2 = convert(newEmployee.Sandvik);
                    usedRange.Cells[currentLine + 2, 18].Value2 = newEmployee.Other;

                    Excel.Range newRowRange = usedRange.Rows[currentLine + 2];

                    //Sets borders for the entire row
                    newRowRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    newRowRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    workbook.Save();

                    //And adds the new row in the dataGridView1
                    dataGridView1.Rows.Add(newEmployee.EmployeeName, newEmployee.Familiarized, newEmployee.FireWorkingDate, newEmployee.FireWorkingNumber, newEmployee.WorkSafetyTraining, newEmployee.WorkSafetyTrainingNumber,
                                newEmployee.ElectricalSafetyTraining, newEmployee.LiveWorkingTraining, newEmployee.FirstAidTraining, newEmployee.Valttikortti, newEmployee.TaxNumber, newEmployee.NokianRenkat1Year, newEmployee.NokianRenkatLOTO,
                                newEmployee.Essity1Year, newEmployee.TampereenSahkolaitos, newEmployee.NvE, newEmployee.Sandvik, newEmployee.Other);
                }
            }
            isProgrammaticEdit = false;
        }

        private class CellChange
        {
            public int RowIndex { get; }
            public int ColumnIndex { get; }
            public object OldValue { get; }
            public object NewValue { get; }
            public Employee Employee { get; set; }

            public CellChange(int rowIndex, int columnIndex, object oldValue, object newValue, Employee employee)
            {
                RowIndex = rowIndex;
                ColumnIndex = columnIndex;
                OldValue = oldValue;
                NewValue = newValue;
                Employee = employee;
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)//This is triggered on the changing of any cell value
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    try
                    {
                        if (cell.Value != null)
                        {
                            if (DateTime.TryParse(cell.Value.ToString(), out DateTime date))
                            //Here it colours the cell if the value entered is a date
                            {
                                DateTime currentDate = DateTime.Today;
                                TimeSpan difference = date - currentDate;
                                int DaysDifference = (int)difference.TotalDays;

                                if (DaysDifference > 31)
                                {
                                    cell.Style.BackColor = Color.LightGreen;
                                }
                                else if (DaysDifference < 31 && DaysDifference > 7)
                                {
                                    cell.Style.BackColor = Color.Khaki;
                                }
                                else if (DaysDifference < 8)
                                {
                                    cell.Style.BackColor = Color.Salmon;
                                }
                            }
                        }
                    }
                    catch
                    {
                    }
                }
            }
        }

        private void EmailConfirmButton_Click(object sender, EventArgs e)//This function is called when the EmailConfirmButton is clicked
        {
            //Here it assigns values necessary for the SaveSettings function
            fromAddress = emailTextBox.Text;
            fromPassword = passwordTextBox.Text;
            if (serviceTextBox.Text == "Outlook")
                smtpClient = "smtp.office365.com";
            else if (serviceTextBox.Text == "Gmail")
                smtpClient = "smtp.gmail.com";
            else if (serviceTextBox.Text == "Yahoo")
                smtpClient = "smtp.mail.yahoo.com";
            credentials.Login = fromAddress;
            credentials.Password = fromPassword;
            credentials.Client = smtpClient;
            if (twoFactorBox.Text != "")
            {
                credentials.InAppPassword = twoFactorBox.Text;
            }

            //And then it calls said function
            SaveSettings(credentials, "credentials.json");
            panel2.Visible = false;
        }

        private void LanguageSelectionBox1_SelectedValueChanged(object sender, EventArgs e)//This is plain language selection
        {
            if (LanguageSelectionBox.Text == "EN")
            {
                MainLanguageEnglish language = Language.mainLanguageSelectEng();
                deleteButton.Text = language.deleteButton;
                EmailButton.Text = language.EmailButton;
                AddEmployeeButton.Text = language.AddEmployeeButton;
                UndoButton.Text = language.UndoButton;
                EmailConfirmButton.Text = language.EmailConfirmButton;
                EmailLabel.Text = language.EmailLabel;
                PasswordLabel.Text = language.PasswordLabel;
                ServiceLabel.Text = language.ServiceLabel;
                TwoFactorLabel.Text = language.TwoFactorLabel;
                BackButton.Text = language.BackButton;

                CheckAllBox.Text = language.CheckAllBox;
                FamiliarizedBox.Text = language.FamiliarizedBox;
                FireDateBox.Text = language.FireDateBox;
                FireNumberBox.Text = language.FireNumberBox;
                WorkSafetyBox.Text = language.WorkSafetyBox;
                WorkSafetyTrainingBox.Text = language.WorkSafetyTrainingBox;
                ElectricalSafetyBox.Text = language.ElectricalSafetyBox;
                LiveBox.Text = language.LiveBox;
                FirstAidBox.Text = language.FirstAidBox;
                TaxBox.Text = language.TaxBox;
                NokianR1YBox.Text = language.NokianR1YBox;
                Essity1YBox.Text = language.Essity1YBox;
                OtherBox.Text = language.OtherBox;

                dataGridView1.Columns["NameColumn"].HeaderText = language.NameColumn;
                dataGridView1.Columns["ElectricalTrainingSafetyColumn"].HeaderText = language.ElectricalTrainingSafetyColumn;
                dataGridView1.Columns["Essity1YearColumn"].HeaderText = language.Essity1YearColumn;
                dataGridView1.Columns["FamiliarizedColumn"].HeaderText = language.FamiliarizedColumn;
                dataGridView1.Columns["FireWorkingDateColumn"].HeaderText = language.FireWorkingDateColumn;
                dataGridView1.Columns["FireWorkingNumberColumn"].HeaderText = language.FireWorkingNumberColumn;
                dataGridView1.Columns["FirstAidTrainingColumn"].HeaderText = language.FirstAidTrainingColumn;
                dataGridView1.Columns["LiveWorkingTrainingColumn"].HeaderText = language.LiveWorkingTrainingColumn;
                dataGridView1.Columns["NokianRenkat1YearColumn"].HeaderText = language.NokianRenkat1YearColumn;
                dataGridView1.Columns["NokianRenkatLOTOColumn"].HeaderText = language.NokianRenkatLOTOColumn;
                dataGridView1.Columns["OtherColumn"].HeaderText = language.OtherColumn;
                dataGridView1.Columns["TaxNumberColumn"].HeaderText = language.TaxNumberColumn;
                dataGridView1.Columns["WorkSafetyTrainingColumn"].HeaderText = language.WorkSafetyTrainingColumn;
                dataGridView1.Columns["WorkSafetyTrainingNumberColumn"].HeaderText = language.WorkSafetyTrainingNumberColumn;
            }
            else
            {
                MainLanguageFinnish language = Language.mainLanguageSelectFin();
                deleteButton.Text = language.deleteButton;
                EmailButton.Text = language.EmailButton;
                AddEmployeeButton.Text = language.AddEmployeeButton;
                UndoButton.Text = language.UndoButton;
                EmailConfirmButton.Text = language.EmailConfirmButton;
                EmailLabel.Text = language.EmailLabel;
                PasswordLabel.Text = language.PasswordLabel;
                ServiceLabel.Text = language.ServiceLabel;
                TwoFactorLabel.Text = language.TwoFactorLabel;
                BackButton.Text = language.BackButton;

                CheckAllBox.Text = language.CheckAllBox;
                FamiliarizedBox.Text = language.FamiliarizedBox;
                FireDateBox.Text = language.FireDateBox;
                FireNumberBox.Text = language.FireNumberBox;
                WorkSafetyBox.Text = language.WorkSafetyBox;
                WorkSafetyTrainingBox.Text = language.WorkSafetyTrainingBox;
                ElectricalSafetyBox.Text = language.ElectricalSafetyBox;
                LiveBox.Text = language.LiveBox;
                FirstAidBox.Text = language.FirstAidBox;
                TaxBox.Text = language.TaxBox;
                NokianR1YBox.Text = language.NokianR1YBox;
                Essity1YBox.Text = language.Essity1YBox;
                OtherBox.Text = language.OtherBox;

                dataGridView1.Columns["NameColumn"].HeaderText = language.NameColumn;
                dataGridView1.Columns["ElectricalTrainingSafetyColumn"].HeaderText = language.ElectricalTrainingSafetyColumn;
                dataGridView1.Columns["Essity1YearColumn"].HeaderText = language.Essity1YearColumn;
                dataGridView1.Columns["FamiliarizedColumn"].HeaderText = language.FamiliarizedColumn;
                dataGridView1.Columns["FireWorkingDateColumn"].HeaderText = language.FireWorkingDateColumn;
                dataGridView1.Columns["FireWorkingNumberColumn"].HeaderText = language.FireWorkingNumberColumn;
                dataGridView1.Columns["FirstAidTrainingColumn"].HeaderText = language.FirstAidTrainingColumn;
                dataGridView1.Columns["LiveWorkingTrainingColumn"].HeaderText = language.LiveWorkingTrainingColumn;
                dataGridView1.Columns["NokianRenkat1YearColumn"].HeaderText = language.NokianRenkat1YearColumn;
                dataGridView1.Columns["NokianRenkatLOTOColumn"].HeaderText = language.NokianRenkatLOTOColumn;
                dataGridView1.Columns["OtherColumn"].HeaderText = language.OtherColumn;
                dataGridView1.Columns["TaxNumberColumn"].HeaderText = language.TaxNumberColumn;
                dataGridView1.Columns["WorkSafetyTrainingColumn"].HeaderText = language.WorkSafetyTrainingColumn;
                dataGridView1.Columns["WorkSafetyTrainingNumberColumn"].HeaderText = language.WorkSafetyTrainingNumberColumn;
            }
        }

        private void UndoButton_Click(object sender, EventArgs e)//This is triggered by clicking the Undo button
        {
            if (CanUndo())
            {
                Undo();
            }
        }

        private void BackButton_Click(object sender, EventArgs e)//This backs out of the email credentials change panel
        {
            panel2.Visible = false;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowSelectedUnselected();
        }
        private void rowSelectedUnselected(object sender, EventArgs e)
        {
            rowSelectedUnselected();
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
    public class EmployeeBoxes
    {
        public bool FamiliarizedBox { get; set; }
        public bool FireDateBox { get; set; }
        public bool FireNumberBox { get; set; }
        public bool WorkSafetyBox { get; set; }
        public bool WorkSafetyTrainingBox { get; set; }
        public bool ElectricalSafetyBox { get; set; }
        public bool LiveBox { get; set; }
        public bool FirstAidBox { get; set; }
        public bool ValttikorttiBox { get; set; }
        public bool TaxBox { get; set; }
        public bool NokianR1YBox { get; set; }
        public bool NokianRLOTOBox { get; set; }
        public bool Essity1YBox { get; set; }
        public bool TampereenSBox { get; set; }
        public bool NvEBox { get; set; }
        public bool SandvikBox { get; set; }
        public bool OtherBox { get; set; }
    }

    public class credentials
    {
        public string Login { get; set; }
        public string Password { get; set; }
        public string Client { get; set; }
        public string InAppPassword { get; set; }
    }
}
