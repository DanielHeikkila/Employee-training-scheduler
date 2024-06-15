namespace Schedule_project
{
    public class Language
    {
        public static MainLanguageEnglish mainLanguageSelectEng()
        {
            MainLanguageEnglish text = new MainLanguageEnglish();
            return text;
        }
        public static MainLanguageFinnish mainLanguageSelectFin()
        {
            MainLanguageFinnish text = new MainLanguageFinnish();
            return text;
        }
        public static NewEmployeeLanguageEnglish NewEmployeeLanguageSelectEng()
        {
            NewEmployeeLanguageEnglish text = new NewEmployeeLanguageEnglish();
            return text;
        }
        public static NewEmployeeLanguageFinnish NewEmployeeLanguageSelectFin()
        {
            NewEmployeeLanguageFinnish text = new NewEmployeeLanguageFinnish();
            return text;
        }
    }
    public class MainLanguageEnglish
    {
        public string deleteButton = "Delete selected";
        public string EmailButton = "Change email";
        public string AddEmployeeButton = "Add new employee";
        public string UndoButton = "Undo";
        public string EmailConfirmButton = "Confirm Email credentials";
        public string EmailLabel = "Email Username";
        public string PasswordLabel = "Email Password";
        public string ServiceLabel = "Email Service";
        public string TwoFactorLabel = "App specific password (two-factor authentication)";
        public string BackButton = "Back";

        public string CheckAllBox = "All";
        public string FamiliarizedBox = "Familiarized";
        public string FireDateBox = "Fire Working Date";
        public string FireNumberBox = "Fire Working Number";
        public string WorkSafetyBox = "Work Safety Training";
        public string WorkSafetyTrainingBox = "Work Safety Training Number";
        public string ElectricalSafetyBox = "Electrical Safety Training";
        public string LiveBox = "Live Working Training";
        public string FirstAidBox = "First Aid Training";
        public string TaxBox = "Tax Number";
        public string NokianR1YBox = "Nokian Renkat 1 Year";
        public string Essity1YBox = "Essity 1 Year";
        public string OtherBox = "Other";

        public string NameColumn = "Name";
        public string ElectricalTrainingSafetyColumn = "Electrical Safety Training";
        public string Essity1YearColumn = "Essity 1 Year";
        public string FamiliarizedColumn = "Familiarized";
        public string FireWorkingDateColumn = "Fire Working Date";
        public string FireWorkingNumberColumn = "Fire Working Number";
        public string FirstAidTrainingColumn = "First Aid Training";
        public string LiveWorkingTrainingColumn = "Live Working Training";
        public string NokianRenkat1YearColumn = "Nokian Renkat 1 Year";
        public string NokianRenkatLOTOColumn = "Nokian Renkat LOTO";
        public string OtherColumn = "Other";
        public string TaxNumberColumn = "Tax Number";
        public string WorkSafetyTrainingColumn = "Work Safety Training";
        public string WorkSafetyTrainingNumberColumn = "Work Safety Training Number";
    }
    public class MainLanguageFinnish
    {
        public string deleteButton = "Poista valittu";
        public string EmailButton = "Muuta sähköpostiosoitetta";
        public string AddEmployeeButton = "Lisää uusi työntekijä";
        public string UndoButton = "Peruuta";
        public string EmailConfirmButton = "Vahvista Sähköpostitiedot";
        public string EmailLabel = "Email Käyttäjätunnus";
        public string PasswordLabel = "Email Salasana";
        public string ServiceLabel = "Email Palvelu";
        public string TwoFactorLabel = "Sovelluskohtainen salasana (kaksitekijätodennus)";
        public string BackButton = "Pois";

        public string CheckAllBox = "Kaikki";
        public string FamiliarizedBox = "TP-kunnossapito perehdytys";
        public string FireDateBox = "Tulityökortti";
        public string FireNumberBox = "Tulityökortin nro";
        public string WorkSafetyBox = "Työturvallisuus-kortti";
        public string WorkSafetyTrainingBox = "Työturv.kortin nro";
        public string ElectricalSafetyBox = "Sähkötyöturvallisuus-kortti";
        public string LiveBox = "Alle 1000V jännitetyökortti";
        public string FirstAidBox = "Hätäensiapu /EA";
        public string TaxBox = "Veronro";
        public string NokianR1YBox = "Nokian Renkat 1 Vuoti";
        public string Essity1YBox = "Essity 1 Vuoti";
        public string OtherBox = "Huomioita";

        public string NameColumn = "Nimi";
        public string ElectricalTrainingSafetyColumn = "Sähkötyöturvallisuus-kortti";
        public string Essity1YearColumn = "Essity 1 Vuoti";
        public string FamiliarizedColumn = "TP-kunnossapito perehdytys";
        public string FireWorkingDateColumn = "Tulityökortti";
        public string FireWorkingNumberColumn = "Tulityökortin nro";
        public string FirstAidTrainingColumn = "Hätäensiapu /EA";
        public string LiveWorkingTrainingColumn = "Alle 1000V jännitetyökortti";
        public string NokianRenkat1YearColumn = "Nokian Renkat 1 Vuoti";
        public string NokianRenkatLOTOColumn = "Nokian Renkat LOTO";
        public string OtherColumn = "Huomioita";
        public string TaxNumberColumn = "Veronro";
        public string WorkSafetyTrainingColumn = "Työturvallisuus-kortti";
        public string WorkSafetyTrainingNumberColumn = "Työturv.kortin nro";
    }
    public class NewEmployeeLanguageEnglish
    {
        public string CreateNameLabel = "Name";
        public string FireWorkingDateCreationLabel = "Fire Working Date";
        public string FireWorkingNumberCreationLabel = "Fire Working Number";
        public string WorkSafetyTrainingCreationLabel = "Work Safety Training";
        public string WorkSafetyTrainingNumberCreationLabel = "Work Safety Training Number";
        public string ElectricalSafetyTrainingCreationLabel = "Electrical Safety Training";
        public string LiveWorkingTrainingCreationLabel = "Live Working Training";
        public string FirstAidTrainingCreationLabel = "First Aid Training";
        public string TaxNumberCreationLabel = "Tax Number";
        public string NokianRenkat1YearCreationLabel = "Nokian Renkat 1 Year";
        public string Esstity1YearCreationLabel = "Essity 1 Year";
        public string OtherCreationLabel = "Other";
        public string FamiliarizedCreationCheck = "Familiarized";
        public string CreateCreationButton = "Create";
    }
    public class NewEmployeeLanguageFinnish
    {
        public string CreateNameLabel = "Nimi";
        public string FireWorkingDateCreationLabel = "Tulityökortti";
        public string FireWorkingNumberCreationLabel = "Tulityökortin nro";
        public string WorkSafetyTrainingCreationLabel = "Työturvallisuus-kortti";
        public string WorkSafetyTrainingNumberCreationLabel = "Työturv.kortin nro";
        public string ElectricalSafetyTrainingCreationLabel = "Sähkötyöturvallisuus-kortti";
        public string LiveWorkingTrainingCreationLabel = "Alle 1000V jännitetyökortti";
        public string FirstAidTrainingCreationLabel = "Hätäensiapu /EA";
        public string TaxNumberCreationLabel = "Veronro";
        public string NokianRenkat1YearCreationLabel = "Nokian Renkat 1 Vuoti";
        public string Esstity1YearCreationLabel = "Essity 1 Vuoti";
        public string OtherCreationLabel = "Huomioita";
        public string FamiliarizedCreationCheck = "TP-kunnossapito perehdytys";
        public string CreateCreationButton = "Luo";
    }
}
