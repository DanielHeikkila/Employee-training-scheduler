using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ScrollBar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Schedule_project
{
    public partial class NewEmployeeForm : Form
    {
        public Employee employeeCreate = new Employee();
        public NewEmployeeForm()
        {
            InitializeComponent();
            panel1.Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 35);
            DialogResult = DialogResult.Cancel;
            NewEmployeeForm_Resize();
        }
        private void NewEmployeeForm_Resize(object sender, EventArgs e)
        {
            NewEmployeeForm_Resize();
        }
        private void NewEmployeeForm_Resize()
        {
            //Adjusting internal controls based on form's new size
            panel1.Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 35);
            panel1.Location = new Point(10, 20);
            FamiliarizedCreationCheck.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.829));
            NameCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.085));
            CreateNameLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.050));
            FireWorkingDateCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.172));
            label1.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.138));
            FireWorkingNumberCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.259));
            label2.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.225));
            WorkSafetyTrainingCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.346));
            label5.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.312));
            WorkSafetyTrainingNumberCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.433));
            label4.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.399));
            ElectricalSafetyTrainingCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.51975));
            label3.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.486));
            LiveWorkingTrainingCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.6067));
            label15.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.573));
            FirstAidTrainingCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.694));
            label14.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.6602));
            ValttikorttiCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.085));
            label11.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.050));
            TaxNumberCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.172));
            label10.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.138));
            NokianRenkat1YearCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.259));
            label9.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.225));
            NokianRenkatLOTOCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.346));
            label8.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.312));
            Esstity1YearCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.433));
            label7.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.399));
            TampereenSahkolaitosCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.51975));
            label6.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.486));
            NvECreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.6067));
            label13.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.573));
            SandvikCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.694));
            label12.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.6602));
            OtherCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.829));
            label16.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.794));
            CreateCreationButton.Size = new Size((int)(this.panel1.Width * 0.184), (int)(this.panel1.Height * 0.069));
            CreateCreationButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.6), (int)(this.referenceLabel.Location.Y * 0.829));
        }

        public void CreateCreationButton_Click(object sender, EventArgs e)
        {
            if (NameCreationBox.Text != "")
            {
                employeeCreate.EmployeeName = NameCreationBox.Text;
                if (FamiliarizedCreationCheck.Checked == true)
                    employeeCreate.Familiarized = true;
                else
                    employeeCreate.Familiarized = false;
                employeeCreate.FireWorkingDate = FireWorkingDateCreationBox.Text;
                employeeCreate.FireWorkingNumber = FireWorkingNumberCreationBox.Text;
                employeeCreate.WorkSafetyTraining = WorkSafetyTrainingCreationBox.Text;
                employeeCreate.WorkSafetyTrainingNumber = WorkSafetyTrainingNumberCreationBox.Text;
                employeeCreate.ElectricalSafetyTraining = ElectricalSafetyTrainingCreationBox.Text;
                employeeCreate.LiveWorkingTraining = LiveWorkingTrainingCreationBox.Text;
                employeeCreate.FirstAidTraining = FirstAidTrainingCreationBox.Text;
                employeeCreate.Valttikortti = ValttikorttiCreationBox.Text;
                employeeCreate.TaxNumber = TaxNumberCreationBox.Text;
                employeeCreate.NokianRenkat1Year = NokianRenkat1YearCreationBox.Text;
                employeeCreate.NokianRenkatLOTO = NokianRenkatLOTOCreationBox.Text;
                employeeCreate.Essity1Year = Esstity1YearCreationBox.Text;
                employeeCreate.TampereenSahkolaitos = TampereenSahkolaitosCreationBox.Text;
                employeeCreate.NvE = NvECreationBox.Text;
                employeeCreate.Sandvik = SandvikCreationBox.Text;
                employeeCreate.Other = OtherCreationBox.Text;
                DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                NameCreationBox.BackColor = Color.Pink;
            }
        }

        public Employee sendNewEmployee()
        {
            return employeeCreate;
        }

        private void Control_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return))
            {
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
        }
    }
}
