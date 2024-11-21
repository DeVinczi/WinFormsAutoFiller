namespace WinFormsAutoFiller.Models.PytaniaDoWnioskuEntity
{
    public class CompanyFormData
    {
        public string NIP { get; set; }
        public string REGON { get; set; }
        public string KRS { get; set; }
        public string PhoneNumber { get; set; }
        public string Fax { get; set; }
        public string AdditionalAddresses { get; set; }
        public ContactPerson PUPContact { get; set; } = new ContactPerson();
        public SigningPerson SigningPerson { get; set; } = new SigningPerson();
        public string PKD { get; set; }

        // Priority Questions
        public string PlannedInvestments { get; set; }

        // Financial Questions
        public string FinancialReportingMethod { get; set; }
        public string AdditionalBusinessAddresses { get; set; }
        public BankAccount BankAccountDetails { get; set; } = new BankAccount();

        // HR Questions
        public EmploymentDetails EmploymentData { get; set; } = new EmploymentDetails();
    }

    public record ContactPerson
    {
        public string Name { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string EmailPassword { get; set; }
        public string Position { get; set; }
        public bool HasElectronicSignature { get; set; }
    }

    public record SigningPerson
    {
        public string Name { get; set; }
        public string HasElectronicSignIn { get; set; }
    }

    public record BankAccount
    {
        public string AccountNumber { get; set; }
        public string BankName { get; set; }
    }

    public record EmploymentDetails
    {
        public int FullTimeEmployees { get; set; }
        public int PartTimeEmployees { get; set; }
        public int ContractEmployees { get; set; }
    }
}
