using System;
using System.Drawing;
using System.Windows.Forms;

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
            LanguageSelectionBox.SelectedIndex = 1;
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
            FireWorkingDateCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.138));
            FireWorkingNumberCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.259));
            FireWorkingNumberCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.225));
            WorkSafetyTrainingCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.346));
            WorkSafetyTrainingCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.312));
            WorkSafetyTrainingNumberCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.433));
            WorkSafetyTrainingNumberCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.399));
            ElectricalSafetyTrainingCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.51975));
            ElectricalSafetyTrainingCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.486));
            LiveWorkingTrainingCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.6067));
            LiveWorkingTrainingCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.573));
            FirstAidTrainingCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.694));
            FirstAidTrainingCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.6602));
            ValttikorttiCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.085));
            ValttikorttiCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.050));
            TaxNumberCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.172));
            TaxNumberCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.138));
            NokianRenkat1YearCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.259));
            NokianRenkat1YearCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.225));
            NokianRenkatLOTOCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.346));
            NokianRenkatLOTOCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.312));
            Esstity1YearCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.433));
            Esstity1YearCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.399));
            TampereenSahkolaitosCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.51975));
            TampereenSahkolaitosCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.486));
            NvECreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.6067));
            NvECreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.573));
            SandvikCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.694));
            SandvikCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.4), (int)(this.referenceLabel.Location.Y * 0.6602));
            OtherCreationBox.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.829));
            OtherCreationLabel.Location = new Point((int)(this.referenceLabel.Location.X * 0.0276), (int)(this.referenceLabel.Location.Y * 0.794));
            CreateCreationButton.Size = new Size((int)(this.panel1.Width * 0.184), (int)(this.panel1.Height * 0.069));
            CreateCreationButton.Location = new Point((int)(this.referenceLabel.Location.X * 0.6), (int)(this.referenceLabel.Location.Y * 0.829));
            if (panel1.Size.Width > 800)
            {
                NameCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                FireWorkingDateCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                FireWorkingNumberCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                WorkSafetyTrainingCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                WorkSafetyTrainingNumberCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                ElectricalSafetyTrainingCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                LiveWorkingTrainingCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                FirstAidTrainingCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                ValttikorttiCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                TaxNumberCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                NokianRenkat1YearCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                NokianRenkatLOTOCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                Esstity1YearCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                TampereenSahkolaitosCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                NvECreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                SandvikCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
                OtherCreationBox.Size = new Size((int)(this.panel1.Width * 0.295), 26);
            }
            else
            {
                NameCreationBox.Size = new Size(240, 26);
                FireWorkingDateCreationBox.Size = new Size(240, 26);
                FireWorkingNumberCreationBox.Size = new Size(240, 26);
                WorkSafetyTrainingCreationBox.Size = new Size(240, 26);
                WorkSafetyTrainingNumberCreationBox.Size = new Size(240, 26);
                ElectricalSafetyTrainingCreationBox.Size = new Size(240, 26);
                LiveWorkingTrainingCreationBox.Size = new Size(240, 26);
                FirstAidTrainingCreationBox.Size = new Size(240, 26);
                ValttikorttiCreationBox.Size = new Size(240, 26);
                TaxNumberCreationBox.Size = new Size(240, 26);
                NokianRenkat1YearCreationBox.Size = new Size(240, 26);
                NokianRenkatLOTOCreationBox.Size = new Size(240, 26);
                Esstity1YearCreationBox.Size = new Size(240, 26);
                TampereenSahkolaitosCreationBox.Size = new Size(240, 26);
                NvECreationBox.Size = new Size(240, 26);
                SandvikCreationBox.Size = new Size(240, 26);
                OtherCreationBox.Size = new Size(240, 26);
            }
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

        private void LanguageSelectionBox_SelectedValueChanged(object sender, EventArgs e)
        {
            if (LanguageSelectionBox.Text == "EN")
            {
                NewEmployeeLanguageEnglish language = Language.NewEmployeeLanguageSelectEng();
                CreateNameLabel.Text = language.CreateNameLabel;
                FireWorkingDateCreationLabel.Text = language.FireWorkingDateCreationLabel;
                FireWorkingNumberCreationLabel.Text = language.FireWorkingNumberCreationLabel;
                WorkSafetyTrainingCreationLabel.Text = language.WorkSafetyTrainingCreationLabel;
                WorkSafetyTrainingNumberCreationLabel.Text = language.WorkSafetyTrainingNumberCreationLabel;
                ElectricalSafetyTrainingCreationLabel.Text = language.ElectricalSafetyTrainingCreationLabel;
                LiveWorkingTrainingCreationLabel.Text = language.LiveWorkingTrainingCreationLabel;
                FirstAidTrainingCreationLabel.Text = language.FirstAidTrainingCreationLabel;
                TaxNumberCreationLabel.Text = language.TaxNumberCreationLabel;
                NokianRenkat1YearCreationLabel.Text = language.NokianRenkat1YearCreationLabel;
                Esstity1YearCreationLabel.Text = language.Esstity1YearCreationLabel;
                OtherCreationLabel.Text = language.OtherCreationLabel;
                FamiliarizedCreationCheck.Text = language.FamiliarizedCreationCheck;
                CreateCreationButton.Text = language.CreateCreationButton;
            }
            else
            {
                NewEmployeeLanguageFinnish language = Language.NewEmployeeLanguageSelectFin();
                CreateNameLabel.Text = language.CreateNameLabel;
                FireWorkingDateCreationLabel.Text = language.FireWorkingDateCreationLabel;
                FireWorkingNumberCreationLabel.Text = language.FireWorkingNumberCreationLabel;
                WorkSafetyTrainingCreationLabel.Text = language.WorkSafetyTrainingCreationLabel;
                WorkSafetyTrainingNumberCreationLabel.Text = language.WorkSafetyTrainingNumberCreationLabel;
                ElectricalSafetyTrainingCreationLabel.Text = language.ElectricalSafetyTrainingCreationLabel;
                LiveWorkingTrainingCreationLabel.Text = language.LiveWorkingTrainingCreationLabel;
                FirstAidTrainingCreationLabel.Text = language.FirstAidTrainingCreationLabel;
                TaxNumberCreationLabel.Text = language.TaxNumberCreationLabel;
                NokianRenkat1YearCreationLabel.Text = language.NokianRenkat1YearCreationLabel;
                Esstity1YearCreationLabel.Text = language.Esstity1YearCreationLabel;
                OtherCreationLabel.Text = language.OtherCreationLabel;
                FamiliarizedCreationCheck.Text = language.FamiliarizedCreationCheck;
                CreateCreationButton.Text = language.CreateCreationButton;
            }
        }
    }
}
