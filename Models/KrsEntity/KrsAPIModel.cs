namespace WinFormsAutoFiller.Models.KrsEntity
{
    public class Odpis
    {
        public string Rodzaj { get; set; }
        public NaglowekA NaglowekA { get; set; }
        public Dane Dane { get; set; }
    }

    public class NaglowekA
    {
        public string Rejestr { get; set; }
        public string NumerKRS { get; set; }
        public string DataCzasOdpisu { get; set; }
        public string StanZDnia { get; set; }
        public string DataRejestracjiWKRS { get; set; }
        public int NumerOstatniegoWpisu { get; set; }
        public string DataOstatniegoWpisu { get; set; }
        public string SygnaturaAktSprawyDotyczacejOstatniegoWpisu { get; set; }
        public string OznaczenieSaduDokonujacegoOstatniegoWpisu { get; set; }
        public int StanPozycji { get; set; }
    }

    public class Dane
    {
        public Dzial1 Dzial1 { get; set; }
        public Dzial2 Dzial2 { get; set; }
        public Dzial3 Dzial3 { get; set; }
    }

    public class Dzial1
    {
        public DanePodmiotu DanePodmiotu { get; set; }
        public SiedzibaIAdres SiedzibaIAdres { get; set; }
        public UmowaStatut UmowaStatut { get; set; }
        public PozostaleInformacje PozostaleInformacje { get; set; }
        public List<Wspolnik> WspolnicySpzoo { get; set; }
        public Kapital Kapital { get; set; }
    }

    public class DanePodmiotu
    {
        public string FormaPrawna { get; set; }
        public Identyfikatory Identyfikatory { get; set; }
        public string Nazwa { get; set; }
        public object DaneOWczesniejszejRejestracji { get; set; }
        public bool CzyProwadziDzialalnoscZInnymiPodmiotami { get; set; }
        public bool CzyPosiadaStatusOPP { get; set; }
    }

    public class Identyfikatory
    {
        public string Regon { get; set; }
        public string Nip { get; set; }
    }

    public class SiedzibaIAdres
    {
        public Siedziba Siedziba { get; set; }
        public Adres Adres { get; set; }
        public string AdresPocztyElektronicznej { get; set; }
        public string AdresStronyInternetowej { get; set; }
    }

    public class Siedziba
    {
        public string Kraj { get; set; }
        public string Wojewodztwo { get; set; }
        public string Powiat { get; set; }
        public string Gmina { get; set; }
        public string Miejscowosc { get; set; }
    }

    public class Adres
    {
        public string Ulica { get; set; }
        public string NrDomu { get; set; }
        public string NrLokalu { get; set; }
        public string Miejscowosc { get; set; }
        public string KodPocztowy { get; set; }
        public string Poczta { get; set; }
        public string Kraj { get; set; }
    }

    public class UmowaStatut
    {
        public List<InformacjaOZawarciuZmianieUmowyStatutu> InformacjaOZawarciuZmianieUmowyStatutu { get; set; }
    }

    public class InformacjaOZawarciuZmianieUmowyStatutu
    {
        public string ZawarcieZmianaUmowyStatutu { get; set; }
    }

    public class PozostaleInformacje
    {
        public string CzasNaJakiUtworzonyZostalPodmiot { get; set; }
        public string InformacjaOLiczbieUdzialow { get; set; }
    }

    public class Wspolnik
    {
        public Nazwisko Nazwisko { get; set; }
        public Imiona Imiona { get; set; }
        public Identyfikator Identyfikator { get; set; }
        public string PosiadaneUdzialy { get; set; }
        public bool CzyPosiadaCaloscUdzialow { get; set; }
    }

    public class Nazwisko
    {
        public string NazwiskoICzlon { get; set; }
    }

    public class Imiona
    {
        public string Imie { get; set; }
        public string ImieDrugie { get; set; }
    }

    public class Identyfikator
    {
        public string Pesel { get; set; }
    }

    public class Kapital
    {
        public WysokoscKapitaluZakladowego WysokoscKapitaluZakladowego { get; set; }
        public object WniesioneAporty { get; set; }
    }

    public class WysokoscKapitaluZakladowego
    {
        public string Wartosc { get; set; }
        public string Waluta { get; set; }
    }

    public class Dzial2
    {
        public Reprezentacja Reprezentacja { get; set; }
    }

    public class Reprezentacja
    {
        public string NazwaOrganu { get; set; }
        public string SposobReprezentacji { get; set; }
        public List<Sklad> Sklad { get; set; }
    }

    public class Sklad
    {
        public Nazwisko Nazwisko { get; set; }
        public Imiona Imiona { get; set; }
        public Identyfikator Identyfikator { get; set; }
        public string FunkcjaWOrganie { get; set; }
        public bool CzyZawieszona { get; set; }
    }

    public class Dzial3
    {
        public PrzedmiotDzialalnosci PrzedmiotDzialalnosci { get; set; }
        public WzmiankiOZlozonychDokumentach WzmiankiOZlozonychDokumentach { get; set; }
    }

    public class PrzedmiotDzialalnosci
    {
        public List<Przedmiot> PrzedmiotPrzewazajacejDzialalnosci { get; set; }
        public List<Przedmiot> PrzedmiotPozostalejDzialalnosci { get; set; }
    }

    public class Przedmiot
    {
        public string Opis { get; set; }
        public string KodDzial { get; set; }
        public string KodKlasa { get; set; }
        public string KodPodklasa { get; set; }
    }

    public class WzmiankiOZlozonychDokumentach
    {
        public List<WzmiankaOZlozeniuRocznegoSprawozdaniaFinansowego> WzmiankaOZlozeniuRocznegoSprawozdaniaFinansowego { get; set; }
    }

    public class WzmiankaOZlozeniuRocznegoSprawozdaniaFinansowego
    {
        public string DataZlozenia { get; set; }
        public string ZaOkresOdDo { get; set; }
    }
}
