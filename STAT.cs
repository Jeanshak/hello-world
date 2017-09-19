using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using System.ComponentModel;
using System.Drawing;
using DevExpress.XtraEditors.Repository;
using System.Security.Cryptography;
using System.Net.Mail;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Globalization;
using System.Diagnostics;


namespace OrderPro
{
  static class STAT
  {
    public static string login_user = "";
    public static string login_pass = "";
    public static bool login_user_admin = false;

    public const string VersionAndDate = "Version 1.296 Datum 15.08.2016";
    public const string Version = "1.296";
    public static int MAXINZNR = -1;
    public static int MAXINZNR_ART = -1;
    public static int MAXINZNR_KND = -1;
    public static frm__HAUPT _Haup;
    public static bool FIRSTTIME = true;
    public static bool FIRSTTIME_MONITOR = true;
    public static bool ArtPosNeuZugefuegt = false;

    public static bool PosAenderungJa = false;
    public static bool PosMengenAenderungJa = false;

    //  public static frmArtikelVergleich _frmArtikelVergleich;
    //public static frm_Stapelverarbeitung _frm_Stapelverarbeitung;
    public static DevExpress.XtraEditors.StyleController styleController;
    //public static _HAUPTSUCHE Auswahl_HAUPTSUCHE = null;
    //public static frm_TABELLE_STAMM_SUCHE MasterTabelle = null;

    // public static frm_TABELLE_PBSEASY_SUCHE PBSTabelle = null;
    // public static frm_ARTIKELANLAGE ArtikelAnlage = null;
    // public static frmStuecklisten Stückliste = null;

    public static string AUFTRAGSNUMMER  = "";
    public static string BESTELLNUMMER = "";
    public static string LS_KNR = "";
    //WSLOG
    public static frm_SWLog swLog = null;
    public static List<WSTask> WSTaskList = new List<WSTask>();
    public static int TaskID = 999;
    public static int update_cnt = 0;

    
    public static List<AuftragArtikelKunden_SQL> KugArtikelList = new List<AuftragArtikelKunden_SQL>();
    public static DataTable LastSearchResult = null;
    public static List<cInstanz> InstanzListe = new List<cInstanz>(); // liste mit allen offenen Instanzen
    public static string dbname = "";
    public static string dbuser = "";
    public static string dbpass = "";
    public static string dbDisplayUser = "";

    public static List<Last10> Last10 = new List<Last10>();
    public static List<string> Kunde_Preislisten = new List<string>();
    public static string LastSuchString = string.Empty;
    public static bool sofortBeenden = false;
    public static string filepath = System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\dma\\OrderPro\\";
    public static System.Diagnostics.Stopwatch Stopuhr = new System.Diagnostics.Stopwatch();
    public static List<string> delFileAtEnd = new List<string>();
      
    // JRO20130513
    public static Int32 Timer_frm_Auftrag_Details_Kalk_Refresh = 1000 * 3 ; // 5 Sek /  1 Min =1000*60*5
    public static bool Aend_Flag_Auftrag_Details_Kalk = false;
    public static frm_Einstellungen _frm_Einstellungen;
    
    public static frm_Auftragerfassen _frm_Auftragerfassen;
    public static frm_Bestellungen _frm_Bestellungen;
    public static frm_Wareneingang _frm_Wareneingang;
    public static frm_AlleAuftraege_Seelensuppe _frm_AlleAuftraege_Seelensuppe;

    // MTW Pfade/Adressen usw.
    public static string MTW_AuftragImportPfad = "";
    public static string MTW_DownloadPfad = "";
    public static string MTW_UploadPfad = "";
    public static string MTW_FTP_ADRESSE = "";
    public static string MTW_FTP_USER = "";
    public static string MTW_FTP_PWD = "";
    public static string MTW_SICHER_PFAD = ""; //MTW_SICHERUNGEN
    public static string MTW_FTP_VON_MTW = "";
    public static string MTW_FTP_AN_MTW = "";
    public static bool Prozess_Aufruf = false;
    public static string ALERT_LOG_XE = "";
    public static string SCAN_PFAD_LIEFERSCHEIN_QUELLE = "";
    public static string SCAN_PFAD_LIEFERSCHEIN_ZIEL = "";
    public static string COMBO_VERPACKUNGSARTIKEL = "";
    public static string COMBO_VERPACKUNGSARTIKEL_WERT_SETZEN = "";
    public static string PFADE_EXPORT_224 = "";

    // Allgemeines 
    public static string ErrorLogPfad = "";
    public static string ErrorLogDatei = "";
    public static Int32 FehlerZaehler = 0;
    public static string PFAD_PDF_ABLAGE = "";
    public static string PFAD_INVENTUR_ABLAGE = "";
    public static string PFAD_FOTOS = "";
    public static string PFAD_LOGOS_KUNDEN = "";
    public static Int32 StationsNr = 0;

    public static string MAIL_BESTELLEN = "";
    public static string MAIL_UMLAGERN	 = "";
    public static string MAIL_SAMMELBOX = ""; // MAIL-SAMMELBOX    
    public static bool bSendungsavisPerMailSenden = false;

    public static string gLieferterminGeaendert;
    public static string gMengeGeaendert;
    // Suchen Rückgaben
    public static string SUCHE_ERGEBNIS1 = "";
    public static string SUCHE_ERGEBNIS2 = "";
    public static string SUCHE_ERGEBNIS3 = "";
    public static string SUCHE_ERGEBNIS4 = "";

    public static string BESTELL_NR_KUNDE = "";
    // JRO20130710
    //public static frm_Artikel _frm_Artikel;
    public static DataTable ArtikelSucheErgebnis = null;
    

    public static fxmrThreading.ThreadFrame threadFrameUpdate;
    public static string picturePath = @"\\fileserv\Artikelbilder\ARTNR_Bilder\";

    public static fxmr_Sql.LoginDaten login = null;
    public static fxmr_Functions.ApplicationStatistics statistics = null;
    public static fxmrThreading.ThreadFrame frm_Bild_Thread = null;
    public static fxmrThreading.ThreadFrame frm_Bild_Update_Thread = null;
    public static fxmrThreading.ThreadFrame ThreadLoadArtikel = null;

    public static System.IO.Pipes.NamedPipeServerStream npsInput = null; 

    public static Color ColorTabsBG = Color.FromArgb(191, 219, 255);
    public static Color ColorTabsFG = Color.FromArgb(21, 66, 139);

    public static Color ColorRowFocus = Color.SteelBlue;
    public static Color ColorRowActive = Color.LightSteelBlue;
    public static Color ColorRowOther = Color.White;
    public static Color ColorCellFocus = Color.OrangeRed;

    public static Color ColorNachlieferung = Color.MediumOrchid;
    public static Color ColorGeloeschterArtikel = Color.OrangeRed;
    public static Color ColorGeliefertArtikel = Color.GreenYellow;
    public static Color ColorVersandStatus = Color.GreenYellow;
    public static Color ColorKommissionierStatus = Color.Aquamarine;
    public static Color ColorPackStatus= Color.Yellow;
    public static Color ColorLieferterminUeberschritten = Color.Red;

    public static string ProgrammPfad = Environment.CurrentDirectory.ToString();
    //System.IO.Directory.SetCurrentDirectory(STAT.ProgrammPfad);    

    //Dürfen die Fenstereinstellungen verändert und gespeichert werden?
    static  public bool dockingAdmin = true; // false; JRO soll frei sein !
    public static int UpdateProgressMax = 1;
    public static int UpdateProgresCurrent = 0;
    public  delegate void DEL_Refresh();
    public static DEL_Refresh DEL_refresh;

    public static Font defaultFont = new Font(new FontFamily("Tahoma"),
                         11,
                         FontStyle.Regular,
                         GraphicsUnit.Pixel);


		
		public static fxmr_Sql.SQL getDBConnection()
    {

      return new fxmr_Sql.SQL(login); 
    }

    public static int KW(DateTime Datum)
    {
      CultureInfo CUI = CultureInfo.CurrentCulture;
      return CUI.Calendar.GetWeekOfYear(Datum, CUI.DateTimeFormat.CalendarWeekRule, CUI.DateTimeFormat.FirstDayOfWeek);
    }


    public static DateTime StartDay(int KW)
    {
      try
      {
        int _tempweek = 0;
        if (0 < KW && KW < 53)
        {
          // auslesen der eingestellten Systemsprache
          CultureInfo CUI = CultureInfo.CreateSpecificCulture("de-DE");

          // festlegen von FirstDate auf den 01.01. des aktuellen Jahres
          DateTime FirstDate = new DateTime(CUI.DateTimeFormat.Calendar.GetYear(DateTime.Now), 1, 1);
          FirstDate = FirstDate.AddDays(KW * 7);
          // solange bis _tempweek = der eingestellten Kalenderwoche ist
          //while (_tempweek < KW)
          //{
          //  FirstDate = FirstDate.AddDays(1);
          //  _tempweek = CUI.Calendar.GetWeekOfYear(FirstDate,
          //                                         CUI.DateTimeFormat.CalendarWeekRule,
          //                                         CUI.DateTimeFormat.FirstDayOfWeek);
          //}
          DayOfWeek FirstDay = CUI.Calendar.GetDayOfWeek(FirstDate);
          return FirstDate;
        }
        else
        {
          throw new ArgumentOutOfRangeException("KW");
        }
      }
      catch (Exception e)
      {

        throw new Exception(String.Format("Eingabe ausserhalb der 0 < Eingabe < 53", e.Message));
      }

    }

    public static void OpenExplorer(string path)
    {
      if (Directory.Exists(path))
        Process.Start("explorer.exe", path);
    }

    public static void ExitApp()
    {
      if (System.Windows.Forms.Application.MessageLoop)
      {
        // Use this since we are a WinForms app
        System.Windows.Forms.Application.Exit();
      }
      else
      {
        // Use this since we are a console app
        System.Environment.Exit(1);
      }
    }
    // Alle Einstellungen aus der Datenbank JRO20140223 
    public static void Vorlauf()
    {
      string cWarnung = "ACHTUNG! WICHTIGE TRANSFERPFADE/SYSTEMPFADE SIND NICHT IM ZUGRIFF - SYSTEMFEHLER - BITTE EDV-LEITUNG INFORMIEREN";
      fxmr_Sql.SQL orac = STAT.getDBConnection();
      string it="SELECT wert FROM Einstellungen where lfnr=1";
      DataTable dt_Einstellungen = orac.SelectTable(it).Tables[0];
      //foreach (DataRow row in dt_Einstellungen.Rows)
      //{
      //  STAT.MTW_AuftragImportPfad = row["wert"].ToString().Trim();
      //  if (!Directory.Exists(STAT.MTW_AuftragImportPfad))
      //  {
      //    try
      //    {
      //      Directory.CreateDirectory(STAT.MTW_AuftragImportPfad);
      //    }
      //    catch (Exception ex)
      //    {
      //      MessageBox.Show("Fehler beim Erzeugen des Ordners '" + STAT.MTW_AuftragImportPfad + "': " + ex.Message, cWarnung);
      //      ExitApp();
      //    }
      //  }
      //}
      // MTW ff.
      it = @"SELECT wert FROM Einstellungen where DETAIL='FTP_DOWN_MTW' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      //foreach (DataRow row in dt_Einstellungen.Rows)
      //{
      //  STAT.MTW_DownloadPfad = row["wert"].ToString().Trim();
      //  if (!Directory.Exists(STAT.MTW_DownloadPfad))
      //  {
      //    try
      //    {
      //      Directory.CreateDirectory(STAT.MTW_DownloadPfad);
      //    }
      //    catch (Exception ex)
      //    {
      //      MessageBox.Show("Fehler beim Erzeugen des Ordners '" + STAT.MTW_DownloadPfad + "': " + ex.Message, cWarnung);
      //      ExitApp();
      //    }
      //  }
      //}
      it = @"SELECT wert FROM Einstellungen where DETAIL='FTP_UP_MTW' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      //foreach (DataRow row in dt_Einstellungen.Rows)
      //{
      //  STAT.MTW_UploadPfad = row["wert"].ToString().Trim();
      //  if (!Directory.Exists(STAT.MTW_UploadPfad))
      //  { try
      //  {
      //    Directory.CreateDirectory(STAT.MTW_UploadPfad);
      //    }
      //    catch (Exception ex)
      //  {
      //    MessageBox.Show("Fehler beim Erzeugen des Ordners '" + STAT.MTW_UploadPfad + "': " + ex.Message, cWarnung);
      //    ExitApp();
      //    }
      //  }
      //}
      it = @"SELECT wert,wert2,wert3 FROM Einstellungen where DETAIL='ADRESSE_USER_PWD' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.MTW_FTP_ADRESSE = row["wert"].ToString().Trim();
        STAT.MTW_FTP_USER = row["wert2"].ToString().Trim();
        STAT.MTW_FTP_PWD = row["wert3"].ToString().Trim();
      }
      it = @"SELECT wert,wert2 FROM Einstellungen where DETAIL='FTP_PFAD_UP_DOWN' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.MTW_FTP_AN_MTW= row["wert"].ToString().Trim();
        STAT.MTW_FTP_VON_MTW = row["wert2"].ToString().Trim();
      }
      it = @"SELECT wert FROM Einstellungen where DETAIL='MTW_SICHERUNGEN' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      //{
      //  STAT.MTW_SICHER_PFAD = row["wert"].ToString().Trim();
      //  if (!Directory.Exists(STAT.MTW_SICHER_PFAD))
      //  {
      //    try
      //    {
      //      Directory.CreateDirectory(STAT.MTW_SICHER_PFAD);
      //    }
      //    catch (Exception ex)
      //    {
      //      MessageBox.Show("Fehler beim Erzeugen des Ordners '" + STAT.MTW_SICHER_PFAD + "': " + ex.Message, cWarnung);
      //      ExitApp();
      //    }
      //  }
      //}
      
      // MTW Ende 
      it = @"SELECT wert, wert2 FROM Einstellungen where DETAIL='ERROR_LOG' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      //foreach (DataRow row in dt_Einstellungen.Rows)
      //{
      //  STAT.ErrorLogPfad  = row["wert"].ToString().Trim();
      //  if (!Directory.Exists(STAT.ErrorLogPfad))
      //  {
      //    try
      //    {
      //      Directory.CreateDirectory(STAT.ErrorLogPfad);
      //    }
      //    catch (Exception ex)
      //    {
      //      MessageBox.Show("Fehler beim Erzeugen des Ordners '" + STAT.ErrorLogPfad + "': " + ex.Message, cWarnung);
      //      ExitApp();
      //    }
      //  }
      //  STAT.ErrorLogDatei = row["wert2"].ToString().Trim();
      //}
      // PFAD_PDF_ABLAGE
      it = @"SELECT wert FROM Einstellungen where DETAIL='PFAD_PDF_ABLAGE' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      //foreach (DataRow row in dt_Einstellungen.Rows)
      //{
      //  STAT.PFAD_PDF_ABLAGE = row["wert"].ToString().Trim();
      //  if (!Directory.Exists(STAT.PFAD_PDF_ABLAGE))
      //  {
      //    try
      //    {
      //      Directory.CreateDirectory(STAT.PFAD_PDF_ABLAGE);
      //    }
      //    catch (Exception ex)
      //    {
      //      if (Program.ProgrammArgument != "Seriennummern_Erfassung".ToUpper()) // hier wird der Pfad nicht benötigt //
      //      {
      //        MessageBox.Show("Fehler beim Erzeugen des Ordners '" + STAT.PFAD_PDF_ABLAGE + "': " + ex.Message, cWarnung);
      //        ExitApp();
      //      }
      //    }
      //  }
      //}

      // PFAD_INVENTUR_ABLAGE
      it = @"SELECT wert FROM Einstellungen where DETAIL='PFAD_INVENTUR_ABLAGE' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      //foreach (DataRow row in dt_Einstellungen.Rows)
      //{
      //  STAT.PFAD_INVENTUR_ABLAGE = row["wert"].ToString().Trim();
      //  if (STAT.PFAD_INVENTUR_ABLAGE != "")
      //  {
      //    string cYear = DateTime.Now.Year.ToString();
      //    string monat = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month);
      //    STAT.PFAD_INVENTUR_ABLAGE = STAT.PFAD_INVENTUR_ABLAGE + cYear + @"\Inventur\Monatsinventuren " + cYear + @"\" + monat + @"\";

      //  }
      //  if (!Directory.Exists(STAT.PFAD_INVENTUR_ABLAGE))
      //  {
      //    try
      //    {
      //      Directory.CreateDirectory(STAT.PFAD_INVENTUR_ABLAGE);
      //    }
      //    catch (Exception ex)
      //    {
      //      if (Program.ProgrammArgument != "PFAD_INVENTUR_ABLAGE".ToUpper()) // hier wird der Pfad nicht benötigt //
      //      {
      //        MessageBox.Show("Fehler beim Erzeugen des Ordners '" + STAT.PFAD_INVENTUR_ABLAGE + "': " + ex.Message, cWarnung);
      //        ExitApp();
      //      }
      //    }
      //  }
      //  break;
      //}

      //// STAT.PFADE_EXPORT_224 = LIEBHERR SPEDITION "EDI-STARTER" 
      //it = @"SELECT wert FROM Einstellungen where BEREICH='PFADE' AND DETAIL='EXPORT_224' ";
      //dt_Einstellungen = orac.SelectTable(it).Tables[0];
      //foreach (DataRow row in dt_Einstellungen.Rows)
      //{
      //  STAT.PFADE_EXPORT_224 = row["wert"].ToString().Trim();
      //  if (!Directory.Exists(STAT.PFADE_EXPORT_224))
      //  {
      //    try
      //    {
      //      Directory.CreateDirectory(STAT.PFADE_EXPORT_224);
      //    }
      //    catch (Exception ex)
      //    {
      //      //if (Program.ProgrammArgument != "PFADE_EXPORT_224".ToUpper()) // hier wird der Pfad nicht benötigt //
      //      //{
      //      //  MessageBox.Show("Fehler beim Erzeugen des Ordners '" + STAT.PFAD_INVENTUR_ABLAGE + "': " + ex.Message, cWarnung);
      //      //  ExitApp();
      //      //}
      //    }
      //  }
      //  break;
      //}
      
      // PFAD_FOTOS 
      it = @"SELECT wert FROM Einstellungen where BEREICH='PFADBILDER' AND DETAIL='FOTOS' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.PFAD_FOTOS = row["wert"].ToString().Trim();
      }

      // PFAD_LOGOS_KUNDEN / JRO20150918 
      it = @"SELECT wert FROM Einstellungen where BEREICH='PFAD_LOGOS_KUNDEN' AND DETAIL='PFAD_LOGOS_KUNDEN' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.PFAD_LOGOS_KUNDEN = row["wert"].ToString().Trim();
      }

      // PFAD_ALERT_LOG_XE >> Datenbank log Alert - Senden
      it = @"SELECT wert FROM Einstellungen where BEREICH='ALERT_LOG_XE' AND DETAIL='ALERT_LOG_XE' ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.ALERT_LOG_XE  = row["wert"].ToString().Trim();
      }
      // SCAN_PFAD_LIEFERSCHEIN
      it = @"SELECT wert,wert2 FROM Einstellungen where BEREICH='SCAN_PFAD_LIEFERSCHEIN'  ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.SCAN_PFAD_LIEFERSCHEIN_QUELLE = row["wert"].ToString().Trim();
        STAT.SCAN_PFAD_LIEFERSCHEIN_ZIEL = row["wert2"].ToString().Trim();
        break;
      }
      //VERPACKUNGSART_COMBO
      STAT.COMBO_VERPACKUNGSARTIKEL = "";
      STAT.COMBO_VERPACKUNGSARTIKEL_WERT_SETZEN = "";
      it = @"SELECT wert,wert2 FROM Einstellungen where BEREICH='VERPACKUNGSART_COMBO'  ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.COMBO_VERPACKUNGSARTIKEL = " " + row["wert"].ToString().Trim() + " ";
        STAT.COMBO_VERPACKUNGSARTIKEL_WERT_SETZEN  = row["wert2"].ToString().Trim();
        break;
      }

      //MAIL-BESTELLEN	Mail für Bestellungen, Minderbestand	Niels.Benne@dma-hoelling.de
      it = @"SELECT wert FROM Einstellungen where BEREICH='MAIL-BESTELLEN'  ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.MAIL_BESTELLEN = row["wert"].ToString().Trim();
        break;
      }
      //MAIL-UMLAGERN	Mail für Umlagerungen	Lager@dma-hoelling.de
      it = @"SELECT wert FROM Einstellungen where BEREICH='MAIL-UMLAGERN'  ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.MAIL_UMLAGERN = row["wert"].ToString().Trim();
        break;
      }
      //MAIL-UMLAGERN	Mail für Umlagerungen	Lager@dma-hoelling.de
      it = @"SELECT wert FROM Einstellungen where BEREICH='MAIL-SAMMELBOX'  ";
      dt_Einstellungen = orac.SelectTable(it).Tables[0];
      foreach (DataRow row in dt_Einstellungen.Rows)
      {
        STAT.MAIL_SAMMELBOX = row["wert"].ToString().Trim();
        break;
      }
      if (STAT.MAIL_SAMMELBOX == "")
      {
        MessageBox.Show("STAT.MAIL_SAMMELBOX  ist leer.", " Kein Mailversand an Mailspeicher für DMAS-Ausgehende Mails");
      }
      // ENDE Vorlauf
    }      
      
    public static Dictionary<int, List<string>> ArtListZwischenspeicher = new Dictionary<int, List<string>>();


//Insert into EINSTELLUNGEN (LFNR,BEREICH,DETAIL,WERT,ANLAGEDAT,BENUTZER,WERT2,WERT3,WERTNUM) values ('501','STAMMDATEN','BETRIEBE','S',to_date('27.11.15','DD.MM.RR'),'Rogge','Speditionen',null,null);
//Insert into EINSTELLUNGEN (LFNR,BEREICH,DETAIL,WERT,ANLAGEDAT,BENUTZER,WERT2,WERT3,WERTNUM) values ('502','STAMMDATEN','BETRIEBE','K',to_date('27.11.15','DD.MM.RR'),'Rogge','Kunden',null,null);
//Insert into EINSTELLUNGEN (LFNR,BEREICH,DETAIL,WERT,ANLAGEDAT,BENUTZER,WERT2,WERT3,WERTNUM) values ('503','STAMMDATEN','BETRIEBE','F',to_date('27.11.15','DD.MM.RR'),'Rogge','Eigene Firmen DMA/Parts',null,null);


    // Combo-Box -..- die Kunden aus der Datenbank laden und in schreiben
    public static void Loadcb_Speditionen(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Speditionen = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt = postgres.SelectTable("SELECT KNR,NAME,NAME2,KURZNAME,ORT FROM KUNDEN WHERE KENNUNG = 'S' ORDER BY KURZNAME").Tables[0];
      if (!boxlist_Speditionen.Contains(box))
      {
        boxlist_Speditionen.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Speditionen)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(400, 600);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt;
        item.Properties.ValueMember = "KNR";
        item.Properties.DisplayMember = "KURZNAME";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Firmentypen aus der Datenbank laden und in cb_xxxx schreiben
    public static void Loadcb_Lieferarten(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Lieferarten = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt_Liefart = postgres.SelectTable("SELECT WERT FROM EINSTELLUNGEN WHERE BEREICH = 'LIEFART' AND DETAIL='LIEFART' ").Tables[0];
      if (!boxlist_Lieferarten.Contains(box))
      {
        boxlist_Lieferarten.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Lieferarten)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(200, 200);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Liefart;
        item.Properties.ValueMember = "WERT";
        item.Properties.DisplayMember = "WERT";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    public static string gConncectionStringSQLServer(string user, string password)
    {
      string constr = "";
      if (user == "" && password == "")
      { constr = @"Data Source=HVH-87-SQL;Initial Catalog=WindreamConnect;Persist Security Info=True;User ID=Rogge;Password=SQHallo14927"; }
      else
      { constr = @"Data Source=HVH-87-SQL;Initial Catalog=WindreamConnect;Persist Security Info=True;User ID=" + user + ";Password=" + password + " "; }

      if (Environment.MachineName == "VM-W81-64") // local JRO
      {
        constr = @"Data Source=VM-W81-64\SQLEXPRESS;Initial Catalog=WindreamConnect;Persist Security Info=True;User ID=sa;Password=hallo";
      }
      return constr;
    }


    public static string gBZPPfad()
    {
      string pfad = @"U:\BZP\BZP_DAT\";
      if (Environment.MachineName == "VM-W81-64") // local JRO
      {
        pfad = @"c:\temp\bzp\bzp_dat\";
      }
      return pfad;
    }

    public static string gBDSDatenPfad()
    {
      string pfad = @"G:\PROGRAMME\BDS2002\bds_daten.dbc";
      if (Environment.MachineName == "VM-W81-64") // local JRO
      {
        pfad = @"c:\temp\bzp\bzp_dat\bds_daten.dbc";
      }
      return pfad;
    }

    public static string gBDSStammPfad()
    {
      string pfad = @"G:\PROGRAMME\BDS2002\bds_Stamm.dbc";
      if (Environment.MachineName == "VM-W81-64") // local JRO
      {
        pfad = @"c:\temp\bzp\bzp_dat\bds_daten.dbc";
      }
      return pfad;
    }

    // Combo-Box -..- die Firmentypen aus der Datenbank laden und in cb_xxxx schreiben
    public static void Loadcb_Firmentypen(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Firmentypen = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt_Lager = postgres.SelectTable("SELECT WERT, WERT2 FROM EINSTELLUNGEN WHERE BEREICH = 'STAMMDATEN' AND DETAIL='BETRIEBE' ").Tables[0];
      if (!boxlist_Firmentypen.Contains(box))
      {
        boxlist_Firmentypen.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Firmentypen)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(600, 400);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Lager;
        item.Properties.ValueMember = "WERT";
        item.Properties.DisplayMember = "WERT2";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }


    // Combo-Box -..- die Versandarten aus der Datenbank laden und in cb_xxxx schreiben
    public static void Loadcb_Versandarten(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Versandarten = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt_Lager = postgres.SelectTable("SELECT BEZEICHNUNG, VNR FROM VERSANDARTEN ").Tables[0];
      if (!boxlist_Versandarten.Contains(box))
      {
        boxlist_Versandarten.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Versandarten)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(600, 400);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Lager;
        item.Properties.ValueMember = "VNR";
        item.Properties.DisplayMember = "BEZEICHNUNG";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Banken aus der Datenbank laden und in cb_xxxx schreiben
    public static void Loadcb_Banken(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Banken = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt_Lager = postgres.SelectTable("SELECT NAME, BNR, IBAN, BIC, BLZ FROM BANKEN ").Tables[0];
      if (!boxlist_Banken.Contains(box))
      {
        boxlist_Banken.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Banken)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(600, 400);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Lager;
        item.Properties.ValueMember = "BNR";
        item.Properties.DisplayMember = "NAME";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Verpackungs-Artikel laden 
    public static void Loadcb_Verpackungs_Artikel(DevExpress.XtraEditors.SearchLookUpEdit box, string Verpackungsartikel)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Verpackungs_Artikel = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      string iiit = "select ARTNR, BEZ, WARENGRUPPE FROM ARTIKEL WHERE ARTNR in " + Verpackungsartikel + " ";
      DataTable dt_Verpackungs_Artikel = postgres.SelectTable(iiit).Tables[0];
      if (!boxlist_Verpackungs_Artikel.Contains(box))
      {
        boxlist_Verpackungs_Artikel.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Verpackungs_Artikel)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(600, 400);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Verpackungs_Artikel;
        item.Properties.ValueMember = "ARTNR";
        item.Properties.DisplayMember = "ARTNR";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Artikel-Lagerorte aus der Datenbank laden und in cb_Lagerort schreiben
    public static void Loadcb_Artikel_Lagerort(DevExpress.XtraEditors.SearchLookUpEdit box, string ARTNR)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Artikel_Lager = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      //DataTable dt_Lager = postgres.SelectTable("select concat(concat(LAGERORT,' - '), BEREICH) as LAGERORT, BEREICH from LAGERORTE ORDER BY LAGERORT, BEREICH").Tables[0];
      DataTable dt_Lager = postgres.SelectTable("select LAGERORT, MENGE, PRIO, MINDESTMENGE, ANLAGE FROM LAGER_ARTIKEL WHERE ARTNR='" + ARTNR + "' ORDER BY PRIO, LAGERORT").Tables[0];
      if (!boxlist_Artikel_Lager.Contains(box))
      {
        boxlist_Artikel_Lager.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Artikel_Lager)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(600, 400);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Lager;
        item.Properties.ValueMember = "LAGERORT";
        item.Properties.DisplayMember = "LAGERORT";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Lagerorte aus der Datenbank laden und in cb_Lagerort schreiben
    public static void Loadcb_Lagerort(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Lager = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      //DataTable dt_Lager = postgres.SelectTable("select concat(concat(LAGERORT,' - '), BEREICH) as LAGERORT, BEREICH from LAGERORTE ORDER BY LAGERORT, BEREICH").Tables[0];
      DataTable dt_Lager = postgres.SelectTable("select LAGERORT, BEREICH, PRIORITAET from LAGERORTE ORDER BY LAGERORT, BEREICH").Tables[0];
      if (!boxlist_Lager.Contains(box))
      {
        boxlist_Lager.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Lager)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(400, 400);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Lager;
        item.Properties.ValueMember = "LAGERORT";
        item.Properties.DisplayMember = "LAGERORT";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Artikel aus der Datenbank laden und in schreiben
    public static void Loadcb_Artikel(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Artikel = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      //DataTable dt_Lager = postgres.SelectTable("select concat(concat(LAGERORT,' - '), BEREICH) as LAGERORT, BEREICH from LAGERORTE ORDER BY LAGERORT, BEREICH").Tables[0];
      DataTable dt_Lager = postgres.SelectTable("select ARTNR, BEZ, WARENGRUPPE from ARTIKEL ORDER BY ARTNR").Tables[0];
      if (!boxlist_Artikel.Contains(box))
      {
        boxlist_Artikel.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Artikel)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(400, 600);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt_Lager;
        item.Properties.ValueMember = "ARTNR";
        item.Properties.DisplayMember = "ARTNR";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Lieferanten aus der Datenbank laden und in schreiben
    public static void Loadcb_Lieferanten(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Lieferanten = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt = postgres.SelectTable("SELECT LFNR, FIRMA, ORT FROM Lieferanten ORDER BY FIRMA").Tables[0];
      if (!boxlist_Lieferanten.Contains(box))
      {
        boxlist_Lieferanten.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Lieferanten)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(400, 600);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt;
        item.Properties.ValueMember = "LFNR";
        item.Properties.DisplayMember = "FIRMA";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }


    // Combo-Box -..- die FIRMEN (DMA & PARTS) aus der Datenbank laden und in schreiben
    public static void Loadcb_Firmen(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Firmen = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt = postgres.SelectTable("SELECT KNR,KURZNAME,NAME2,ORT FROM KUNDEN WHERE KENNUNG = 'F' ORDER BY KURZNAME").Tables[0];
      if (!boxlist_Firmen.Contains(box))
      {
        boxlist_Firmen.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Firmen)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(400, 600);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt;
        item.Properties.ValueMember = "KNR";
        item.Properties.DisplayMember = "KURZNAME";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Kunden aus der Datenbank laden und in schreiben
    public static void Loadcb_Kunden(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Kunden = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt = postgres.SelectTable("SELECT KNR,KURZNAME,NAME2,ORT FROM KUNDEN WHERE KENNUNG = 'K' ORDER BY KURZNAME").Tables[0];
      if (!boxlist_Kunden.Contains(box))
      {
        boxlist_Kunden.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Kunden)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(400, 600);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt;
        item.Properties.ValueMember = "KNR";
        item.Properties.DisplayMember = "KURZNAME";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Abteilungen für Mitarbeiter aus der Datenbank laden und in Box schreiben
    public static void Loadcb_ABTEILUNG(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_ABTEILUNG = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt_ABTEILUNG = postgres.SelectTable("SELECT NR, BEZ FROM ABTEILUNG ORDER BY BEZ  ").Tables[0];
      if (!boxlist_ABTEILUNG.Contains(box))
      {
        boxlist_ABTEILUNG.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_ABTEILUNG)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(300, 350);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_ABTEILUNG;
        item.Properties.ValueMember = "NR";
        item.Properties.DisplayMember = "BEZ";
        item.Properties.PopulateViewColumns();
      }
    }
    // Combo-Box -..- die Länder aus der Datenbank laden und in Box schreiben
    public static void Loadcb_Laender(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Laender = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      //DataTable dt_Lager = postgres.SelectTable("select concat(concat(LAGERORT,' - '), BEREICH) as LAGERORT, BEREICH from LAGERORTE ORDER BY LAGERORT, BEREICH").Tables[0];
      DataTable dt_Laender = postgres.SelectTable("select BEZ_DEU,LAND,BEZ_INT from LAENDER ORDER BY BEZ_DEU ").Tables[0];
      if (!boxlist_Laender.Contains(box))
      {
        boxlist_Laender.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Laender)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(300, 350);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Laender;
        item.Properties.ValueMember = "LAND";
        item.Properties.DisplayMember = "BEZ_DEU";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Combo-Box -..- die Langetexte aus der Datenbank laden und in Objekt schreiben
    public static void Loadcb_Lieferbedingungen(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Lieferbedingung = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      // 100	TEXTBAUSTEIN	Lieferbedingung	ODER Zahlungsbedingungen ODER Preisstellung
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      string it = "select LFNR, WERT3 from EINSTELLUNGEN WHERE BEREICH= 'TEXTBAUSTEIN' AND UPPER(DETAIL)= UPPER('Lieferbedingung') ORDER BY WERT3 ";
      DataTable dt_Texte = postgres.SelectTable(it).Tables[0];
      if (!boxlist_Lieferbedingung.Contains(box))
      {
        boxlist_Lieferbedingung.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Lieferbedingung)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(700, 200);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt_Texte;
        item.Properties.ValueMember = "LFNR";
        item.Properties.DisplayMember = "WERT3";
        item.Properties.PopulateViewColumns();
        item.Properties.View.Columns["LFNR"].Visible = false;
      }
    }

    // Combo-Box -..- die Langetexte aus der Datenbank laden und in Objekt schreiben
    public static void Loadcb_Zahlungsbedingungen(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Zahlungsbedingung = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      // 100	TEXTBAUSTEIN	Lieferbedingung	ODER Zahlungsbedingungen ODER Preisstellung
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      string it = "select LFNR, WERT3 from EINSTELLUNGEN WHERE BEREICH= 'TEXTBAUSTEIN' AND UPPER(DETAIL)= UPPER('Zahlungsbedingungen') ORDER BY WERT3 ";
      DataTable dt_Texte = postgres.SelectTable(it).Tables[0];
      if (!boxlist_Zahlungsbedingung.Contains(box))
      {
        boxlist_Zahlungsbedingung.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Zahlungsbedingung)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(700, 300);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt_Texte;
        item.Properties.ValueMember = "LFNR";
        item.Properties.DisplayMember = "WERT3";
        item.Properties.PopulateViewColumns();
        item.Properties.View.Columns["LFNR"].Visible = false;
      }
    }

    // Combo-Box -..- die Langetexte aus der Datenbank laden und in Objekt schreiben
    public static void Loadcb_Preisstellung(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Preisstellung = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      // 100	TEXTBAUSTEIN	Lieferbedingung	ODER Zahlungsbedingungen ODER Preisstellung
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      string it = "select LFNR, WERT3 from EINSTELLUNGEN WHERE BEREICH= 'TEXTBAUSTEIN' AND UPPER(DETAIL)= UPPER('Preisstellung') ORDER BY WERT3 ";
      DataTable dt_Texte = postgres.SelectTable(it).Tables[0];
      if (!boxlist_Preisstellung.Contains(box))
      {
        boxlist_Preisstellung.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Preisstellung)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(700, 300);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt_Texte;
        item.Properties.ValueMember = "LFNR";
        item.Properties.DisplayMember = "WERT3";
        item.Properties.PopulateViewColumns();
        item.Properties.View.Columns["LFNR"].Visible = false;
      }
    }



    // Combo-Box -..- die Kunde-ArtNrn aus der Datenbank laden und in Objekt schreiben
    public static void Loadcb_KundeArtNr(DevExpress.XtraEditors.SearchLookUpEdit box, string KNR, string Artnr)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_KundeArtNr = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      string it = @"select KUNDEARTNR, ARTNR, VKPREIS from KUNDEN_ARTIKEL where KUNDEN_ARTIKEL.KNR = " + KNR + " AND KUNDEN_ARTIKEL.ARTNR = '" + Artnr + "' ORDER BY KUNDEARTNR ";
      DataTable dt_Texte = postgres.SelectTable(it).Tables[0];
      if (!boxlist_KundeArtNr.Contains(box))
      {
        boxlist_KundeArtNr.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_KundeArtNr)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(400, 500);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt_Texte;
        item.Properties.ValueMember = "KUNDEARTNR";
        item.Properties.DisplayMember = "KUNDEARTNR";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["ARTNR"].Visible = false;
      }
    }


    // Combo-Box -..- Druckwiederholung EINLAGERUNGSLISTE
    public static void Loadcb_Druckwiederholung(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Einlagerungsliste = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      string it = @"select DISTINCT DRUCK_OK from EINLAGERUNGSLISTE ORDER BY DRUCK_OK ";
      DataTable dt_Texte = postgres.SelectTable(it).Tables[0];
      if (!boxlist_Einlagerungsliste.Contains(box))
      {
        boxlist_Einlagerungsliste.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Einlagerungsliste)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(400, 500);
        item.Properties.ShowClearButton = true;
        item.Properties.ShowFooter = true;
        item.Properties.DataSource = dt_Texte;
        item.Properties.ValueMember = "DRUCK_OK";
        item.Properties.DisplayMember = "DRUCK_OK";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["ARTNR"].Visible = false;
      }
    }


    // Speziell für Sendungsavis Vor-Form und Frachtbrief ****  ****  ****  ****  ****  ****  ****  ****  ****  ****  ****  **** 
    // Combo-Box -..- Versand-Avis-Zeiten ab ... Uhr
    public static void Loadcb_Versand_Avis_Zeiten(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Versand_Avis_Zeiten = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt_Table = postgres.SelectTable("SELECT WERT FROM EINSTELLUNGEN WHERE BEREICH = 'TEXTBAUSTEIN' AND DETAIL='VERSAND' ORDER BY WERT ").Tables[0];
      if (!boxlist_Versand_Avis_Zeiten.Contains(box))
      {
        boxlist_Versand_Avis_Zeiten.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Versand_Avis_Zeiten)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(200,200);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Table;
        item.Properties.ValueMember = "WERT";
        item.Properties.DisplayMember = "WERT";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Speziell für Sendungsavis Vor-Form und Frachtbrief ****  ****  ****  ****  ****  ****  ****  ****  ****  ****  ****  **** 
    // Combo-Box -..- Ladezeiten Mo-Fr 10-14 Uhr usw.
    public static void Loadcb_Ladezeiten(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Ladezeiten = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt_Table = postgres.SelectTable("SELECT WERT FROM EINSTELLUNGEN WHERE BEREICH = 'TEXTBAUSTEIN' AND DETAIL='LADEZEITEN' ORDER BY WERT ").Tables[0];
      if (!boxlist_Ladezeiten.Contains(box))
      {
        boxlist_Ladezeiten.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Ladezeiten)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(200, 200);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Table;
        item.Properties.ValueMember = "WERT";
        item.Properties.DisplayMember = "WERT";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    // Speziell für Sendungsavis Vor-Form und Frachtbrief ****  ****  ****  ****  ****  ****  ****  ****  ****  ****  ****  **** 
    public static void Loadcb_Express(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_Ladezeiten = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
      fxmr_Sql.SQL postgres = STAT.getDBConnection();
      DataTable dt_Table = postgres.SelectTable("SELECT WERT FROM EINSTELLUNGEN WHERE BEREICH = 'TEXTBAUSTEIN' AND DETAIL='EXPRESS' ORDER BY WERT ").Tables[0];
      if (!boxlist_Ladezeiten.Contains(box))
      {
        boxlist_Ladezeiten.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_Ladezeiten)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(200, 200);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_Table;
        item.Properties.ValueMember = "WERT";
        item.Properties.DisplayMember = "WERT";
        item.Properties.PopulateViewColumns();
        //item.Properties.View.Columns["LAGERORT"].Visible = false;
      }
    }

    public static void lockGUI( WeifenLuo.WinFormsUI.Docking.DockContent dock)
    {
        if (!dockingAdmin)
        {
          // Fenster schließbar ja oder nein // closable 
          dock.DockAreas = WeifenLuo.WinFormsUI.Docking.DockAreas.Document;
          //JRO20140326 dock.CloseButtonVisible = false;
          // Wenn Sie die Module freigeben (dock.AllowDrop = true & dock.AllowEndUserDocking = true) können sie verschoben werden und 
          // auch die Positionen wird beim Schließen/Öffnen gespeichert bzw. geladen.
          dock.CloseButtonVisible = true;
          dock.AllowDrop = true;            // false;
          dock.AllowEndUserDocking = true;  // false;   
        }
        else
        {
            dock.DockAreas = WeifenLuo.WinFormsUI.Docking.DockAreas.Document | WeifenLuo.WinFormsUI.Docking.DockAreas.Float
                | WeifenLuo.WinFormsUI.Docking.DockAreas.DockLeft | WeifenLuo.WinFormsUI.Docking.DockAreas.DockRight 
                | WeifenLuo.WinFormsUI.Docking.DockAreas.DockBottom | WeifenLuo.WinFormsUI.Docking.DockAreas.DockTop;
            dock.CloseButtonVisible = true;
            dock.AllowDrop = true;
            dock.AllowEndUserDocking = true;
     
        }
        dock.Update();
    }


    public static Dictionary<string, bool> SuchergebnisColumns = new Dictionary<string, bool>();
    public static DataTable WringEans(DataTable source, string EANfeld)
    {
        DataTable dt = source;
        DataTable dtneu = dt.Clone();
        dtneu.Rows.Clear();
        foreach (DataRow row in dt.Rows)
        {
            string ean = row[EANfeld].ToString();
            //if (ValidateEAN(ean) == false) //falsche eans in neue tabelle
            //{
            //    dtneu.Rows.Add(row);
            //}
        }
        return dtneu;
    }


    public static bool stringCheckIsNumber(string str)
    {
        char[] StrArr = str.ToCharArray(); ;
        // string strNeu = string.Empty;
        bool fehler = false;
        foreach (char c in StrArr)
        {
            int i = 0;
            if (!int.TryParse(c.ToString(), out i))
            {
                //strNeu += c;
                fehler = true;
            }

        }
        return fehler;
    }

    public static bool stringCheckIsArtNr(string str)
    {
        bool fehler = false;
        if (str.Length != 10)
        {
            fehler = true;
        }
        else
        {
            fehler = stringCheckIsNumber(str);
        }
        return fehler;
    }

    public static bool stringCheckOracle(string str)
    {

        return true;
    }

    public static void Fehleranzeige(string str)
    {
      if (str != "")
      {
        STAT._Haup.ClearInfoBox(); // 
        STAT._Haup.AddInfoBox("ERR:" + str, "", 3);
      }

    }

    public static void saveDGVSettings(DevExpress.XtraGrid.GridControl grid, string filename)
    {
        string dgvsfile = STAT.filepath + filename + ".xml";
        try
        {
            DevExpress.XtraGrid.Views.Grid.GridView gv = ((DevExpress.XtraGrid.Views.Grid.GridView)(grid.MainView));
            gv.SaveLayoutToXml(dgvsfile);
            //fxmr_Einstellungen.XMLSET set = new fxmr_Einstellungen.XMLSET();
            //foreach (DataGridViewColumn item in dgv.Columns)
            //{
            //    set.createOrAddOrChange_Int(dgvsfile, dgv.Name + "W" + item.HeaderText, item.Width);
            //    set.createOrAddOrChange_Int(dgvsfile, dgv.Name + "I" + item.HeaderText, item.DisplayIndex);
            //}
            //if (dgv.SortedColumn == null)
            //{
            //    set.createOrAddOrChange_String(dgvsfile, dgv.Name + "SC", "");
            //    set.createOrAddOrChange_String(dgvsfile, dgv.Name + "SO", "");
            //}
            //else
            //{
            //    set.createOrAddOrChange_String(dgvsfile, dgv.Name + "SC", dgv.SortedColumn.Name.ToString());
            //    set.createOrAddOrChange_String(dgvsfile, dgv.Name + "SO", dgv.SortOrder.ToString());
            //}
        }
        catch (Exception)
        { }

    }

    public static void loadDGVSettings(DevExpress.XtraGrid.GridControl grid, string filename)
    {

        string dgvsfile = STAT.filepath  +filename+".xml";

        try
        {
            DevExpress.XtraGrid.Views.Grid.GridView gv = ((DevExpress.XtraGrid.Views.Grid.GridView)(grid.MainView));
            gv.RestoreLayoutFromXml(dgvsfile);
            //fxmr_Einstellungen.XMLSET set = new fxmr_Einstellungen.XMLSET();
            //foreach (DataGridViewColumn item in dgv.Columns)
            //{

            //    item.Width = set.read_Int(dgvsfile, dgv.Name + "W" + item.HeaderText);
            //    item.DisplayIndex = set.read_Int(dgvsfile, dgv.Name + "I" + item.HeaderText);


            //}

            //ListSortDirection sort = ListSortDirection.Descending;
            //if (set.read_String(dgvsfile, dgv.Name + "SO") == "Ascending")
            //{
            //    sort = ListSortDirection.Ascending;
            //}
            //dgv.Sort(dgv.Columns[set.read_String(dgvsfile, dgv.Name + "SC")], sort);
        }
        catch (Exception)
        { }
    }

    // Nur als Beispiel - not in use Füllen Combobox Auftragsart
    public static List<DevExpress.XtraEditors.SearchLookUpEdit> boxlist_aagaufart = new List<DevExpress.XtraEditors.SearchLookUpEdit>();
    public static DataTable dt_aagaufart = null;
    public static void LoadKriterium(DevExpress.XtraEditors.SearchLookUpEdit box)
    {
      fxmr_Sql.SQL orac = STAT.getDBConnection();
      dt_aagaufart = orac.SelectTable("SELECT DISTINCT KRITERIUM FROM ARTLISTE").Tables[0];
      if (!boxlist_aagaufart.Contains(box))
      {
        boxlist_aagaufart.Add(box);
      }
      foreach (DevExpress.XtraEditors.SearchLookUpEdit item in boxlist_aagaufart)
      {
        item.Properties.PopupFormSize = new System.Drawing.Size(300, 350);
        item.Properties.ShowClearButton = false;
        item.Properties.ShowFooter = false;
        item.Properties.DataSource = dt_aagaufart;
        item.Properties.ValueMember = "bez";
        item.Properties.DisplayMember = "bez";
        item.Properties.PopulateViewColumns();
        //JRO20140216 item.Properties.View.Columns["nr"].Visible = true;
      }
    }      
           
    public static void WriteWindowSize(Form frm, string pathfile, bool showMenu)
    {
        if (!STAT.sofortBeenden)
        {
            fxmr_Einstellungen.XMLSET set = new fxmr_Einstellungen.XMLSET();
            if (pathfile == "")
            {
                pathfile = STAT.filepath + "size.xml";
            }
            set.createOrAddOrChange_Bool(pathfile, frm.Name + "_Menu", showMenu);
            set.createOrAddOrChange_Bool(pathfile, frm.Name + "_LockView", STAT._Haup.dockPanel1.AllowEndUserDocking);
            set.createOrAddOrChange_Int(pathfile, frm.Name + "_LocX", frm.Location.X);
            set.createOrAddOrChange_Int(pathfile, frm.Name + "_LocY", frm.Location.Y);
            set.createOrAddOrChange_Int(pathfile, frm.Name + "_SizeW", frm.Size.Width);
            set.createOrAddOrChange_Int(pathfile, frm.Name + "_SizeH", frm.Size.Height);
            switch (frm.WindowState)
            {
                case FormWindowState.Maximized:
                    set.createOrAddOrChange_Bool(STAT.filepath + "size.xml", frm.Name + "_Maximized", true);
                    break;
                default:
                    set.createOrAddOrChange_Bool(STAT.filepath + "size.xml", frm.Name + "_Maximized", false);
                    break;
            }  
        }
           
    }

    public static bool ReadWindowSize(Form frm, string pathfile)
    {
        if (pathfile == "")
        {
            pathfile = STAT.filepath + "size.xml";
        }

        if (System.IO.File.Exists(pathfile))
        {
            fxmr_Einstellungen.XMLSET set = new fxmr_Einstellungen.XMLSET();
            STAT._Haup.dockPanel1.AllowEndUserDocking = set.read_Bool(pathfile, frm.Name + "_LockView");
            //if (STAT._Haup.dockPanel1.AllowEndUserDocking)
            //{ STAT._Haup.btChangeLock.Caption = "Ansicht sperren"; }
            //else
            //{ STAT._Haup.btChangeLock.Caption = "Ansicht lösen"; }

            frm.Location = new System.Drawing.Point(set.read_Int(pathfile, frm.Name + "_LocX"), set.read_Int(pathfile, frm.Name + "_LocY"));
            frm.Size = new System.Drawing.Size(set.read_Int(pathfile, frm.Name + "_SizeW"), set.read_Int(pathfile, frm.Name + "_SizeH"));
            if ( frm.Location.X ==-1 &&  frm.Location.Y == -1)
            {
                frm.Location = new Point(10, 10);
                frm.Size = new Size(900, 600);
            }
            if (set.read_Bool(pathfile, frm.Name + "_Maximized"))
            {
                frm.WindowState = FormWindowState.Maximized;
            }
            return set.read_Bool(STAT.filepath + "size.xml", frm.Name + "_Menu");
        }
        return false;
    }

    public static void SetFloatStyle(WeifenLuo.WinFormsUI.Docking.DockContent frm)
    {
        frm.DockStateChanged += new EventHandler(frm_DockStateChanged);
    }

    static void frm_DockStateChanged(object sender, EventArgs e)
    {
        WeifenLuo.WinFormsUI.Docking.DockContent frm = (WeifenLuo.WinFormsUI.Docking.DockContent)sender;
        if (frm.DockState == WeifenLuo.WinFormsUI.Docking.DockState.Float)
        {
            frm.FloatPane.FindForm().FormBorderStyle = FormBorderStyle.Sizable;
            frm.FloatPane.FindForm().ShowInTaskbar = true;
            frm.FloatPane.FindForm().Owner = null;
            frm.FloatPane.FindForm().Icon = frm.FindForm().Icon;
            //JRO20140407 frm.FloatPane.FindForm().MaximizeBox = false;
            frm.FloatPane.FindForm().MaximizeBox = true;
            //frm.FloatPane.FindForm().Size = frm.FindForm().MinimumSize;
            //frm.FloatPane.FindForm().MinimumSize = frm.FindForm().MinimumSize; 
        }
    }

    public static double minDB = 0.1;
    public static bool autoCalc = true;
    public static void KalkulationDeckungsbeitrag(string DB, string EK, string VK, bool calcPreis, DevExpress.XtraEditors.TextEdit cnt, DevExpress.XtraEditors.TextEdit DBcnt)
    {
        double res = 0;

        if (autoCalc)
        {
            EK = EK.Replace('.', ',');
            VK = VK.Replace('.', ',');
            DB = DB.Replace('.', ',');
                
            try
            {
                if (calcPreis)    //  Preis errechnen
                {
                    if (DB != "n. def." & DB != "" & DB != "0" & EK != "n. def." & EK != "" & EK != "0")
                    {
                        double ek = (double.Parse(EK));
                        double db = (double.Parse(DB));
                        res = Math.Round(1*ek/(1-db), 4);
                        //'=1*ek/(1-db%)
                        autoCalc = false;
                        if (res.ToString() != cnt.EditValue.ToString())
                        { cnt.EditValue = res; }
                        autoCalc = true;
                    }
                }
                else   //DB errechnen
                {
                    if (EK != "n. def." & EK != "" & EK != "0" & VK != "n. def." & VK != "" & VK != "0")
                    {
                        double ek = (double.Parse(EK));
                        double vk = (double.Parse(VK));
                        res = Math.Round((vk - ek) / vk, 4);

                        autoCalc = false;
                        if (res.ToString() != cnt.EditValue.ToString())
                        { cnt.EditValue = res; }
                        autoCalc = true;
                    }
                }

                if (double.Parse(DBcnt.EditValue.ToString()) < minDB)
                { DBcnt.ForeColor = Color.Red; }
                else
                { DBcnt.ForeColor = Color.Black; }
            }



            catch (Exception)
            { }

        }

    }



    public static void ExecuteBatFile(string batfile)
   {
      // Task als Admin starten !
      const int ERROR_CANCELLED = 1223; //The operation was canceled by the user.
      ProcessStartInfo info = new ProcessStartInfo(batfile);
      info.UseShellExecute = true;
      info.Verb = "runas";
      try
      {
        Process.Start(info);
      }
      catch (Win32Exception ex)
      {
        if (ex.NativeErrorCode == ERROR_CANCELLED)
          MessageBox.Show("Why you no select Yes?");
        else
          throw;
      }
      //FileInfo BatDateiInfo = new FileInfo(batfile);
      //string baName = BatDateiInfo.Name;
      //string baPfad = BatDateiInfo.Directory + @"\";
      //Process proc = null;
      //string _batDir = string.Format(baPfad);
      //proc = new Process();
      //proc.StartInfo.WorkingDirectory = _batDir;
      //proc.StartInfo.FileName = baName;
      //proc.StartInfo.CreateNoWindow = true;
      //proc.StartInfo.UseShellExecute = true;
      //proc.StartInfo.Verb = "runas";
      //proc.Start();
      //proc.WaitForExit();
      //proc.Close();
      //MessageBox.Show("Bat " + batfile + " file executed...");
  }


  public static void WRITE_ASC_FILE(string FileName, string Inhalt, bool EncodeStandard)
    {
      // NET-Framework stellt folgende Implementierungen der Encoding-Klasse zur Unterstützung der aktuellen Unicode-Codierungen und anderer Codierungen bereit: 
      //ASCIIEncoding codiert Unicode-Zeichen als einzelne 7-Bit-ASCII-Zeichen. Diese Codierung unterstützt nur Zeichenwerte zwischen U+0000 und U+007F. Codepage 20127. Auch verfügbar durch die ASCII-Eigenschaft.
      //UTF7Encoding codiert Unicode-Zeichen mit der UTF-7-Codierung. Diese Codierung unterstützt alle Unicode-Zeichenwerte. Codepage 65000. Auch verfügbar durch die UTF7-Eigenschaft.
      //UTF8Encoding codiert Unicode-Zeichen mit der UTF-8-Codierung. Diese Codierung unterstützt alle Unicode-Zeichenwerte. Codepage 65001. Auch verfügbar durch die UTF8-Eigenschaft.
      //UnicodeEncoding codiert Unicode-Zeichen mit der UTF-16-Codierung. Dabei werden die Bytereihenfolgen Little-Endian (Codepage 1200) und Big-Endian (Codepage 1201) unterstützt. Auch verfügbar durch die Unicode-Eigenschaft und die BigEndianUnicode-Eigenschaft.
      //UTF32Encoding codiert Unicode-Zeichen mit der UTF-32-Codierung. Dabei werden die Bytereihenfolgen Little-Endian (Codepage 65005) und Big-Endian (Codepage 65006) unterstützt. Auch verfügbar durch die UTF32-Eigenschaft.
      FileInfo ProDatei = new FileInfo(FileName);
      string Path = ProDatei.DirectoryName;
      if (!System.IO.Directory.Exists(Path))
      {
        try
        {
          System.IO.Directory.CreateDirectory(Path);
        }
        catch (Exception ex)
        {
          Console.WriteLine("Fehler beim Erzeugen des Ordners '" +
              STAT.MTW_DownloadPfad + "': " + ex.Message);
        }
      }
      if (EncodeStandard == true)
      {
        System.IO.StreamWriter analyse_testsql = new System.IO.StreamWriter(FileName, true);
        analyse_testsql.WriteLine(Inhalt);
        analyse_testsql.Close();
      }
      else
      {
        System.Text.Encoding Xcode = System.Text.Encoding.Default;
        System.IO.StreamWriter analyse_testsql = new System.IO.StreamWriter(FileName, true, Xcode);
        analyse_testsql.WriteLine(Inhalt + Environment.NewLine);
        analyse_testsql.Close();
      }
    }

    public static void ASC_2_FILE(string Path, string FileName, string Inhalt, bool EncodeStandard)
     {
      // NET-Framework stellt folgende Implementierungen der Encoding-Klasse zur Unterstützung der aktuellen Unicode-Codierungen und anderer Codierungen bereit: 
      //ASCIIEncoding codiert Unicode-Zeichen als einzelne 7-Bit-ASCII-Zeichen. Diese Codierung unterstützt nur Zeichenwerte zwischen U+0000 und U+007F. Codepage 20127. Auch verfügbar durch die ASCII-Eigenschaft.
      //UTF7Encoding codiert Unicode-Zeichen mit der UTF-7-Codierung. Diese Codierung unterstützt alle Unicode-Zeichenwerte. Codepage 65000. Auch verfügbar durch die UTF7-Eigenschaft.
      //UTF8Encoding codiert Unicode-Zeichen mit der UTF-8-Codierung. Diese Codierung unterstützt alle Unicode-Zeichenwerte. Codepage 65001. Auch verfügbar durch die UTF8-Eigenschaft.
      //UnicodeEncoding codiert Unicode-Zeichen mit der UTF-16-Codierung. Dabei werden die Bytereihenfolgen Little-Endian (Codepage 1200) und Big-Endian (Codepage 1201) unterstützt. Auch verfügbar durch die Unicode-Eigenschaft und die BigEndianUnicode-Eigenschaft.
      //UTF32Encoding codiert Unicode-Zeichen mit der UTF-32-Codierung. Dabei werden die Bytereihenfolgen Little-Endian (Codepage 65005) und Big-Endian (Codepage 65006) unterstützt. Auch verfügbar durch die UTF32-Eigenschaft.
       if (!System.IO.Directory.Exists(Path))
       {
         try
         {
           System.IO.Directory.CreateDirectory(Path);
         }
         catch (Exception ex)
         {
           Console.WriteLine("Fehler beim Erzeugen des Ordners '" +
               STAT.MTW_DownloadPfad + "': " + ex.Message);
         }
       }
       if (EncodeStandard == true)
       {
         System.IO.StreamWriter analyse_testsql = new System.IO.StreamWriter(Path + FileName, true);
         analyse_testsql.WriteLine(Inhalt);
         analyse_testsql.Close();
       }
       else
       {
         System.Text.Encoding Xcode = System.Text.Encoding.Default;
         System.IO.StreamWriter analyse_testsql = new System.IO.StreamWriter(Path + FileName, true, Xcode);
         analyse_testsql.WriteLine(Inhalt + Environment.NewLine);
         analyse_testsql.Close();
       }
     }

    public static DataTable LogTable = null;

    public static void AddLogEntry(string name, string beschreibung, Form sender, int level)
    {
        if (LogTable == null)
        {
            LogTable = new DataTable();
            LogTable.Columns.Add("Timestamp", typeof(DateTime));
            LogTable.Columns.Add("Name");
            LogTable.Columns.Add("Beschreibung");
            LogTable.Columns.Add("Ort");
            LogTable.Columns.Add("Level", typeof(int));
        }
        LogTable.Rows.Add(DateTime.Now, name, beschreibung, sender.Name, level);

    }


    // das muss später in eine dll, da der server die methoden auch benötigt
    static public string CreateSalt(string UserName)
    {
        System.Security.Cryptography.Rfc2898DeriveBytes hasher = new System.Security.Cryptography.Rfc2898DeriveBytes(UserName,
            System.Text.Encoding.Default.GetBytes("ye$6sQwbH|aFR~a+_0o"), 10000);
        return Convert.ToBase64String(hasher.GetBytes(25));
    }

    static public string HashPassword(string Salt, string Password)
    {
        System.Security.Cryptography.Rfc2898DeriveBytes Hasher = new Rfc2898DeriveBytes(Password,
            System.Text.Encoding.Default.GetBytes(Salt), 10000);
        return Convert.ToBase64String(Hasher.GetBytes(25));
    }
       
    public static void AddLogEntry(Exception ex,Form sender)
    {
        if (LogTable == null)
        {
            LogTable = new DataTable();
            LogTable.Columns.Add("Timestamp", typeof(DateTime));
            LogTable.Columns.Add("Name");
            LogTable.Columns.Add("Beschreibung");
            LogTable.Columns.Add("Ort");
            LogTable.Columns.Add("Level", typeof(int));
        }
        LogTable.Rows.Add(DateTime.Now, ex.Message, ex.HelpLink, sender.Name, 2);

    }

    public static void RunningThreadsCount()
    {

        if (threadFrameUpdate != null)
        {
            _Haup.lbTasks.Caption = "Tasks: " + threadFrameUpdate.runningThreads.ToString() + " / " + threadFrameUpdate.waitingThreads.ToString() + "\t\n";

        }
          
          
    }

    public static Int64[] ExtractNumbers(this string source, bool extractOnlyIntegers)
    {
      // Muster für den regulären Ausdruck definieren
      string pattern;
      if (extractOnlyIntegers)
      {
        pattern = @"\d{1,}";
      }
      else
      {
        string decimalSeparator = Regex.Escape(
           Thread.CurrentThread.CurrentCulture.NumberFormat.
           NumberDecimalSeparator);
        pattern = @"\d{1,}" + decimalSeparator + @"{0,1}\d{0,}";
      }

      // Die Treffer ermitteln
      MatchCollection matches = Regex.Matches(source, pattern);

      // Das Ergebnis in ein double-Array kopieren
      Int64[] result = new Int64[matches.Count];
      for (int i = 0; i < matches.Count; i++)
      {
        result[i] = Convert.ToInt64(matches[i].Value);
      }

      return result;
    }

    public static String SaeubernFuerDateinamen(this string source)
    {
      string cLegal = "ABCDEFGHIJKLMNOPQRTSTUVWXYZ0123456789 -.()[]";
      string cZeichen = "";
      string cRueckgabe = "";
      source = source.Trim().ToUpper();
      for (int i = 0; i < source.Length; i++)
      {
        cZeichen = source.Substring(i, 1).ToUpper();
        if (cLegal.IndexOf(cZeichen) >= 0)
        {
          cRueckgabe = cRueckgabe + cZeichen;
        }
        else
        { cRueckgabe = cRueckgabe + "_"; }
      }
      return cRueckgabe;
    }

    public static String DezimalPunkt(this string source)
    {
      string cRueckgabe = "";
      int Pos = 0;
      Pos = source.IndexOf(",");
      if (Pos >= 0)
      {
        cRueckgabe = source.Substring(0, Pos) + "." + source.Substring(Pos + 1);
      }
      else
      { cRueckgabe = source; }
      return cRueckgabe;
    }

    public static String DezimalKomma(this string source)
    {
      string cRueckgabe = "";
      int Pos = 0;
      Pos = source.IndexOf(".");
      if (Pos >= 0)
      {
        cRueckgabe = source.Substring(0, Pos) + "," + source.Substring(Pos + 1);
      }
      else
      { cRueckgabe = source; }
      return cRueckgabe;
    }

    public static String SaeubernFuerSQL(this string source)
    {
      string cIllegal = "'`´&";
      string cZeichen = "";
      string cRueckgabe = "";
      source = source.Trim();
      for (int i = 0; i < source.Length; i++)
      {
        cZeichen = source.Substring(i, 1);
        if (cIllegal.IndexOf(cZeichen) >= 0)
        { cRueckgabe = cRueckgabe + " "; }
        else
        { cRueckgabe = cRueckgabe + cZeichen; }
      }
      return cRueckgabe;
    }

    public static Decimal WandelInDezimal(this string source)
    {//JRO20110629
      Decimal nRueckgabe = 0;
      if (source == "")
      { source = "0"; }

      nRueckgabe = Convert.ToDecimal(source);
      return nRueckgabe;
    }

    public static String SaeubernNumerisch(this string source)
    {
      string cLegal = "0123456789,. ";
      string cZeichen = "";
      string cRueckgabe = "";
      source = source.Trim().ToUpper();
      for (int i = 0; i < source.Length; i++)
      {
        cZeichen = source.Substring(i, 1).ToUpper();
        if (cLegal.IndexOf(cZeichen) >= 0)
        {
          cRueckgabe = cRueckgabe + cZeichen;
        }
      }
      if (cRueckgabe == "")
      { cRueckgabe = "0"; }
      return cRueckgabe;
    }

    public static string DATUMFORMAT_YYMMTT(this DateTime heute)
    {
      string cDatum = "";
      cDatum = DateTime.Now.Year.ToString().Substring(2, 2) + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0');
      return cDatum;
    }
    public static string DATUMFORMAT_YYYYMMTT(this DateTime heute)
    {
      string cDatum = "";
      cDatum = DateTime.Now.Year.ToString().Substring(0, 4) + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0');
      return cDatum;
    }
    public static string ZEITFORMAT_HHMMSS(this DateTime heute)
    {
      string cDatum = "";
      cDatum = DateTime.Now.Hour.ToString().PadLeft(2, '0') + DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');
      return cDatum;
    }
    public static bool TestRechner()
    {
      string MeinRechnerName = Environment.MachineName.ToUpper().Trim(); // Rechnername  Environment.MachineName = VMWARE-XPG 
      bool Tester = (MeinRechnerName == "VMWARE-XPG" || MeinRechnerName == "W7E-00-30G-DE");
      return Tester;
    }

    public static string MailSender(string SUBJEKT, string Inhalt, string Datei, bool bAnTeam)
    {
      string Status = "";
      // Mailzugang //
      if (Inhalt == "") { Inhalt = SUBJEKT; }
      string smptServer = "192.168.3.1";
      string userName = "Admin";
      string password = "WAg3dPD3";
      if (STAT.TestRechner())
      {
        smptServer = "smtp.strato.de";
        userName = "software@coro.de";
        password = "hallo14927";
      }
      string cBody = "";
      try
      {
        MailMessage myMessage = new MailMessage(); // http://msdn.microsoft.com/de-de/library/system.net.mail.mailmessage(v=vs.110).aspx
        // Empfänger :
        myMessage.To.Add("meldungen@coro.de");               // Hier die Empängeradresse(n) !
        if (bAnTeam == true)
        {
          myMessage.To.Add("2@coro.de");                     // Hier die Empängeradresse(n) !
        }
        if (STAT.MAIL_SAMMELBOX != "")
        {
          myMessage.To.Add(STAT.MAIL_SAMMELBOX);                     // Hier die Empängeradresse(n) !
        }
        //
        myMessage.From = new MailAddress("dmas@dma-hoelling.de", "Automatische Benachrichtigung OrderPro");
        myMessage.Subject = SUBJEKT;
        myMessage.Priority = MailPriority.High;
        // Anhang //
        if (Datei != "")
        {
          FileInfo UploadDatei = new FileInfo(Datei);
          if (UploadDatei.Exists)
          {
            Attachment Attachdata = new Attachment(Datei);
            myMessage.Attachments.Add(Attachdata);
          }
        }
        cBody = Inhalt;
        cBody = cBody + Environment.NewLine + Environment.NewLine + Environment.NewLine + "2@coro.de"; // Testmail
        myMessage.Body = cBody;
        // von extern SmtpClient client = new SmtpClient("kugelmann.dyndns.org", 25);
        //SmtpClient client = new SmtpClient("remote.dma-hoelling.de", 25); // Mailserver
        SmtpClient client = new SmtpClient(smptServer, 25); // Mailserver
        client.Credentials = new System.Net.NetworkCredential(userName, password);
        client.Timeout = 360000;
        client.Send(myMessage);
        Thread.Sleep(1);
      }
      catch (Exception ex)
      {
        string s = ex.Message.ToString();
        Status = s;
      }

      return Status;
    }

    public static string MailSender_Empfaenger(string Empfaenger1, string Empfaenger2, string Empfaenger3, string SUBJEKT, string Inhalt, string Datei, string VersenderAdresse)
    {
      
      // return "";
      // Testweise ausschalten ! erstmal Mail ausgeschaltet ... 

      string Status = "";
      // Mailzugang //
      if (Inhalt == "") { Inhalt = SUBJEKT; }
      string smptServer = "192.168.3.1";
      string userName = "Admin";
      string password = "WAg3dPD3";
      smptServer = "smtp.strato.de";
      userName = "software@coro.de";
      password = "hallo14927";
      smptServer = "hvh-11-lotus.heinzvonheiden.de";
      userName = "hello";
      password = "hallo";
      string cBody = "";
      try
      {
        MailMessage myMessage = new MailMessage(); // http://msdn.microsoft.com/de-de/library/system.net.mail.mailmessage(v=vs.110).aspx
        // Empfaenger1
        myMessage.To.Add(Empfaenger1);               // Hier die Empängeradresse(n) !
        if (Empfaenger2 != "")
        {
          myMessage.To.Add(Empfaenger2);                     // Hier die Empängeradresse(n) !
        }
        if (Empfaenger3 != "")
        {
          myMessage.To.Add(Empfaenger3);                     // Hier die Empängeradresse(n) !
        }
        if (STAT.MAIL_SAMMELBOX != "")
        {
          myMessage.To.Add(STAT.MAIL_SAMMELBOX);                     // Hier die Empängeradresse(n) !
        }
        //
        VersenderAdresse = VersenderAdresse == "" ? @"software@coro.de" : VersenderAdresse;
        myMessage.From = new MailAddress(VersenderAdresse, SUBJEKT);
        myMessage.Subject = SUBJEKT;
        myMessage.Priority = MailPriority.High;
        // Anhang //
        if (Datei != "")
        {
          FileInfo UploadDatei = new FileInfo(Datei);
          if (UploadDatei.Exists)
          {
            Attachment Attachdata = new Attachment(Datei);
            myMessage.Attachments.Add(Attachdata);
          }
        }
        cBody = Inhalt;
        cBody = cBody + Environment.NewLine + Environment.NewLine + Environment.NewLine + "2@coro.de"; // Testmail
        myMessage.Body = cBody;
        // von extern SmtpClient client = new SmtpClient("kugelmann.dyndns.org", 25);
        // SmtpClient client = new SmtpClient("remote.dma-hoelling.de", 25); // Mailserver
        SmtpClient client = new SmtpClient(smptServer, 25); // Mailserver
        client.Credentials = new System.Net.NetworkCredential(userName, password);
        client.Timeout = 360000;
        client.Send(myMessage);
        Thread.Sleep(1);

        
      }
      catch (Exception ex)
      {
        string s = ex.Message.ToString();
        Status = s;
      }
      return Status;
    }

    // nur Test  !!!!!// nur Test  !!!!!// nur Test  !!!!!// nur Test  !!!!!// nur Test  !!!!!// nur Test  !!!!!// nur Test  !!!!!// nur Test  !!!!!
    public static string MailSender_Empfaenger_SBS(string Empfaenger1, string Empfaenger2, string Empfaenger3, string SUBJEKT, string Inhalt, string Datei, string VersenderAdresse)
    {
      //  Zum Senden von Emails über den SBS muss man einfach als SMTP-Server die IP 
      //192.168.3.1 angeben und als Absenderadresse admin@dma-hoelling.de.
      string Status = "";
      // Mailzugang //
      if (Inhalt == "") { Inhalt = SUBJEKT; }
      string smptServer = "192.168.3.1";
      string userName = "Admin";
      string password = "WAg3dPD3";
      if (STAT.TestRechner())
      {
        smptServer = "smtp.strato.de";
        userName = "software@coro.de";
        password = "hallo14927";
      }
      string cBody = "";
      try
      {
        MailMessage myMessage = new MailMessage(); // http://msdn.microsoft.com/de-de/library/system.net.mail.mailmessage(v=vs.110).aspx
        // Empfaenger1
        myMessage.To.Add(Empfaenger1);               // Hier die Empängeradresse(n) !
        if (Empfaenger2 != "")
        {
          myMessage.To.Add(Empfaenger2);                     // Hier die Empängeradresse(n) !
        }
        if (Empfaenger3 != "")
        {
          myMessage.To.Add(Empfaenger3);                     // Hier die Empängeradresse(n) !
        }
        if (STAT.MAIL_SAMMELBOX != "")
        {
          myMessage.To.Add(STAT.MAIL_SAMMELBOX);                     // Hier die Empängeradresse(n) !
        }
        //
        VersenderAdresse = VersenderAdresse == "" ? @"admin@dma-hoelling.de" : VersenderAdresse;
        myMessage.From = new MailAddress(VersenderAdresse, SUBJEKT);
        myMessage.Subject = SUBJEKT;
        myMessage.Priority = MailPriority.High;
        // Anhang //
        if (Datei != "")
        {
          FileInfo UploadDatei = new FileInfo(Datei);
          if (UploadDatei.Exists)
          {
            Attachment Attachdata = new Attachment(Datei);
            myMessage.Attachments.Add(Attachdata);
          }
        }
        cBody = Inhalt;
        cBody = cBody + Environment.NewLine + Environment.NewLine + Environment.NewLine; // Testmail
        myMessage.Body = cBody;
        // von extern SmtpClient client = new SmtpClient("kugelmann.dyndns.org", 25);
        //SmtpClient client = new SmtpClient("remote.dma-hoelling.de", 25); // Mailserver
        SmtpClient client = new SmtpClient(smptServer, 25); // Mailserver
        client.Credentials = new System.Net.NetworkCredential(userName, password);
        client.Timeout = 360000;
        client.Send(myMessage);
        Thread.Sleep(1);
      }
      catch (Exception ex)
      {
        string s = ex.Message.ToString();
        Status = s;
      }
      return Status;
    }

    public static void PrintScreen()
    {
      if (!Directory.Exists(@"C:\Temp\"))
      { try
        { Directory.CreateDirectory(@"C:\Temp\");
        }
        catch (Exception ex)
        { MessageBox.Show("Fehler beim Erzeugen des Ordners '" + @"C:\Temp\" + "': " + ex.Message);
        }
      }
      string DateiBild = @"C:\Temp\Bildschirm_" + DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Ticks.ToString() + ".jpg";
      Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
      Graphics graphics = Graphics.FromImage(printscreen as Image);
      graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);
      printscreen.Save(DateiBild, ImageFormat.Jpeg);
      PrintDocument pd = new PrintDocument();
      //pd.DefaultPageSettings.PrinterSettings.PrinterName = "PDFCreator";
      pd.DefaultPageSettings.Landscape = true; //or false!
      pd.PrintPage += (sender, args) =>
      {
        Image i = Image.FromFile(DateiBild);
        Rectangle m = args.MarginBounds;

        if ((double)i.Width / (double)i.Height > (double)m.Width / (double)m.Height) // image is wider
        {
          m.Height = (int)((double)i.Height / (double)i.Width * (double)m.Width);
        }
        else
        {
          m.Width = (int)((double)i.Width / (double)i.Height * (double)m.Height);
        }
        args.Graphics.DrawImage(i, m);
      };
      pd.Print();
      MessageBox.Show(@"Die gedruckte Datei liegt als JPG-Bilddatei auf " + DateiBild + " (Ablage/Versand/Mail) vor.", " BILDSCHIRM-AUSDRUCK ERSTELLT ");
    }

    public static void PrintAllScreens()
    {
      if (!Directory.Exists(@"C:\Temp\"))
      {
        try
        {
          Directory.CreateDirectory(@"C:\Temp\");
        }
        catch (Exception ex)
        {
          MessageBox.Show("Fehler beim Erzeugen des Ordners '" + @"C:\Temp\" + "': " + ex.Message);
        }
      }
      string DateiBild = @"C:\Temp\Bildschirm_" + DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Ticks.ToString(); //  +".jpg";
      string TempDateiBild = "";
      int iBild = 0;
      Screen[] AllScreens = Screen.AllScreens;
      foreach (Screen item in AllScreens)
      {
        iBild++;
        TempDateiBild = DateiBild + "_" + iBild.ToString() + ".jpg";
        Bitmap printscreen = new Bitmap(item.Bounds.Width, item.Bounds.Height);
        Graphics graphics = Graphics.FromImage(printscreen as Image);
        graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);
        printscreen.Save(TempDateiBild, ImageFormat.Jpeg);
        PrintDocument pd = new PrintDocument();
        //pd.DefaultPageSettings.PrinterSettings.PrinterName = "PDFCreator";
        pd.DefaultPageSettings.Landscape = true; //or false!
        pd.PrintPage += (sender, args) =>
        {
          Image i = Image.FromFile(TempDateiBild);
          Rectangle m = args.MarginBounds;

          if ((double)i.Width / (double)i.Height > (double)m.Width / (double)m.Height) // image is wider
          {
            m.Height = (int)((double)i.Height / (double)i.Width * (double)m.Width);
          }
          else
          {
            m.Width = (int)((double)i.Width / (double)i.Height * (double)m.Height);
          }
          args.Graphics.DrawImage(i, m);
        };
        pd.Print();
      }
      //MessageBox.Show(@"Die gedruckte Datei liegt als JPG-Bilddatei auf " + DateiBild + " (Ablage/Versand/Mail) vor.", " BILDSCHIRM-AUSDRUCK ERSTELLT ");
    }
}

  public class FormStart
  {
      // public List<frm_PBSEAN> LIST_PBSZUORDNUNG = new List<frm_PBSEAN>();
      // public List<frm_TABELLE_STAMM_SUCHE> LIST_TABELLE = new List<frm_TABELLE_STAMM_SUCHE>();
      // public List<frm_INFO> LIST_INFO = new List<frm_INFO>();
      // public List<frm_PBSDATEN> LIST_PBSDATEN = new List<frm_PBSDATEN>();
      // public List<frm_PBSZUORD> LIST_PBSZUORD = new List<frm_PBSZUORD>();

  }
    
  public class Last10
  {
      public Last10( string bez, string Lfdnr)
      {
          this.bez = bez;
          this.lfdnr = Lfdnr;
      }

      public string lfdnr = string.Empty;
      public string bez = string.Empty;
  }

  public class FormInfoObject
  {
    public Size FormDefaultSize = new Size(100, 100);
    public FormInfoObject(Size formDefaultSize)
    {
        FormDefaultSize = formDefaultSize;
    }
  }

  public class IniFile
  {
    public string path;

    [System.Runtime.InteropServices.DllImport("kernel32")]
    private static extern long WritePrivateProfileString(string section,
      string key, string val, string filePath);

    [System.Runtime.InteropServices.DllImport("kernel32")]
    private static extern int GetPrivateProfileString(string section,
      string key, string def, StringBuilder retVal,
      int size, string filePath);

    public IniFile(string INIPath)
    {
      path = INIPath;
    }

    public void IniWriteValue(string Section, string Key, string Value)
    {
      WritePrivateProfileString(Section, Key, Value, this.path);
    }

    public string IniReadValue(string Section, string Key)
    {
      StringBuilder temp = new StringBuilder(255);
      int i = GetPrivateProfileString(Section, Key, "", temp, 255, this.path);
      return temp.ToString();
    }
  }
}
