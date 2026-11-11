using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Threading;
using System.ComponentModel;
using System.Collections;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using Microsoft.Win32;


// Bemærk at visual studio skal køres som administrator ellers virker det ikke

//namespace MathMenuLib  // giver vidst nok problemer med com
//{

//interface som viser det som ses fra com i vba
public interface IMaximaConn
    {
        void StartMaximaProcess();
        void SetSettings();
        void RunFile(string MaxSti, string Arguments);
        void SetMaximaPath(string MaxSti);
        void Reset(string extrakommando);
        void CloseProcess();
        void ConsoleInterrupt();
        void ExecuteMaximaCommand(string inputstring, int wait);
        void AnswerQuestion(string inputText);
        void TurnUnitsOn(string prevcomm, string postcomm);
        void UpdateUnits();
        string MaximaInputArray(int nr);
        string MaximaOutputArray(int nr);
        string CheckForUpdate();
        string GetVersionNews();
        string RegKeyRead(string keyName);
        void RegKeySave(string fullPath, object value);
        bool RegKeyExists(string fullPath);
        void RegKeyDelete(string fullPath);
        void OpenFolder(string path);
        void CreateZipFile(string zipPath, string filePath);
        string LastMaximaOutput
        {
            get;
        }
        string MaximaOutput
        {
            get;
        }
        int Question
        {
            get;
        }
        int ErrCode
        {
            get;
        }
        int AntalCifre
        { get; set; }
        int Exact
        { get; set; }
        int Complex
        { get; set; }
        int Units
        { get; set; }
        int Finished
        { get; set; }
        string OutUnits
        { get; set; }
        bool IsMaximaStarted
        { get; }
        // New members are added here, at the end, so the existing members keep their places
        int RunAndWait(string filePath, string arguments, string workingDirectory, int timeoutSeconds);
        int UpdateDownloadStart(bool beta, bool nonAdmin);
        int UpdateDownloadProgress();
        void UpdateDownloadCancel();
        int UpdateInstall();
        int LicenseCheckStart(bool force, string userName);
        int LicenseCheckResult();
        string LicenseName();
        int LicensePurchasePage(bool danish);
        int VersionCheckStart(bool partnership);
        int VersionCheckResult();
        string VersionCheckText();
        void Pause(int milliseconds);
        string ScreenInfo(int hwnd);
        int WindowToFront(string className);
        int HideConsoleWindow();

}/* end interface IMaximaConn */

//    [GuidAttribute("CE99F51A-727A-11DF-B26D-D5FCDFD72085")]  // burde blive autogenereret men er måske ikke consistent så
//    [ComVisible(true)] // ikke nødvendigt hvis sat i assemblyinfo
//    [ProgId("MathMenuLib.MaximaConn")]
[ClassInterface(ClassInterfaceType.None)]   // AutoDispatch og AutoDual  er også muligt men giver problemer med nye versioner og virker kun med latebind
    public class MaximaProcessClass : IMaximaConn
        {
            // Define static variables shared by class methods.
            // Der må ikke være noget der er static
            private StringBuilder MaximaOutputBuilder = null;
            private StringBuilder LastMaximaOutputBuilder = null;
            private string[] maximaoutput= new string[1000]; // array der indeholder output fra Maxima efter nr i Maxima
            private string[] maximainput = new string[1000];
            private Process maximaProcess;
            private StreamWriter maximaStreamWriter;
            private int numOutputLines = 0;
            private int KommandoNr = 0; //senest udførte kommandonr.
            private int mindex = 0; // index der angiver hvilket (%inr) der er i Maxima pt. 
            private bool isMaximaStarted = false;
            private bool finish;
            private bool outputstarted; // sættes så snart der kommer det første output
            private bool outputstarted2; // sættes så snart sidste "(%onn" kommer i ouput
            private int question;
            public int errCode = 0; //0 ingen fejl, 1 - Maxima ikke installeret, 
            public bool LastWasOutput; // true hvis forrige linje var en outputlinje false hvis input
            public string MaximaPath=""; // Til hovedmappen hvor Maxima er.
            private int complex = 0; // true hvis der regnes med komplekse tal
            private int units = 0; // true hvis der regnes med enheder.
            private int exact = 0; // 0- auto 1-exact 2-numer  (tidligere:0-exact 1-num 2-både exact og num)
            private int antalcifre= 7; // 0-exact 1-num 2-både exact og num
            private string moutunits="";

        // der skal være en constructor og den må ikke tage nogen parametre
            public MaximaProcessClass()
            {
//                StartMaximaProcess(); ' bedre at vente så der kan sættes indstillinger inden f.eks. units
            }


            ~MaximaProcessClass()
            {
                try
                {
                    maximaStreamWriter.Close();
                }
                catch { }
                try
                {
                    maximaProcess.Close();
                }
                catch { }
            }

    public void SetMaximaPath(string MaxSti)
    {
        MaximaPath = MaxSti;
    }

    public void RunFile(string filePath, string arguments = "")
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = filePath,
            Arguments = arguments,
            UseShellExecute = true
        });
    }

    // Runs a program (not through a command interpreter) with its window hidden, in workingDirectory, and waits
    // until it has ended, at most timeoutSeconds (0 = no limit). Used by WordMat+ for pdflatex.
    // Returns the program's exit code, -1 if it was stopped at the timeout, -2 if it could not be started.
    public int RunAndWait(string filePath, string arguments, string workingDirectory, int timeoutSeconds)
    {
        try
        {
            using (Process p = new Process())
            {
                p.StartInfo.FileName = filePath;
                p.StartInfo.Arguments = arguments ?? "";
                if (!string.IsNullOrEmpty(workingDirectory))
                    p.StartInfo.WorkingDirectory = workingDirectory;
                p.StartInfo.UseShellExecute = false; // the program itself: no shell, no cmd.exe
                p.StartInfo.CreateNoWindow = true;   // no console window; the output is not read (no pipes to fill up)
                p.Start();
                if (timeoutSeconds > 0)
                {
                    if (!p.WaitForExit(timeoutSeconds * 1000))
                    {
                        try { p.Kill(); } catch { }
                        return -1;
                    }
                }
                else
                {
                    p.WaitForExit();
                }
                return p.ExitCode;
            }
        }
        catch
        {
            return -2;
        }
    }

    // ---------------------------------------------------------------------------------------------------------
    // The update of WordMat (WordMat+: WMPUpdateWordMat in ModuleHjaelp).
    // The caller gives no address and no file name: WordMat's installer is downloaded from WordMat's release page
    // (UpdateUrl) to the user's Downloads folder as WordMatU.exe, and it is only kept, and only started, when it has
    // a valid digital signature of UpdateSigner. The download runs on its own thread, so Word can show its progress:
    //   UpdateDownloadStart(beta, nonAdmin)   0 = started, -4 = a download is already running
    //   UpdateDownloadProgress()              0-99 = percent downloaded, 100 = downloaded and its signature is valid,
    //                                         -1 = could not be downloaded, -2 = no valid signature of UpdateSigner (the
    //                                         file is deleted), -3 = stopped, -4 = no download was started
    //   UpdateDownloadCancel()                stops the download
    //   UpdateInstall()                       starts the installer (silent; it asks the user to close Word):
    //                                         0 = started, -2 = the signature is not valid, -4 = nothing downloaded,
    //                                         -5 = it could not be started
    // ---------------------------------------------------------------------------------------------------------
    private const string UpdateUrl = "https://github.com/Eduap-com/WordMat/releases/download/v1.17/";
    // the names (the certificate's CN) WordMat's installers are signed with: the company's, and its name before 2026
    private static readonly string[] UpdateSigner = { "Eduap ApS", "Eduap v/Mikael Samsøe Sørensen" };
    private const long UpdateMinSize = 10000000; // an installer is far larger; an error page is far smaller
    private volatile int updateState = -4;       // UpdateDownloadProgress
    private volatile bool updateCancel;
    private string updateFile;                   // the downloaded installer, when updateState is 100
    private Thread updateThread;

    public int UpdateDownloadStart(bool beta, bool nonAdmin)
    {
        if (updateThread != null && updateThread.IsAlive) return -4;
        string name = (nonAdmin ? "WordMatNA" : "WordMat") + (beta ? "b" : "") + ".exe";
        string url = UpdateUrl + name;
        string file = Path.Combine(UpdateDownloadsFolder(), "WordMatU.exe");
        updateFile = null;
        updateCancel = false;
        updateState = 0;
        updateThread = new Thread(() => UpdateDownload(url, file));
        updateThread.IsBackground = true;
        updateThread.Start();
        return 0;
    }

    public int UpdateDownloadProgress()
    {
        return updateState;
    }

    public void UpdateDownloadCancel()
    {
        updateCancel = true;
    }

    public int UpdateInstall()
    {
        string file = updateFile;
        if (updateState != 100 || string.IsNullOrEmpty(file) || !File.Exists(file)) return -4;
        if (!UpdateSignatureValid(file)) return -2; // again: the file has been in the Downloads folder since the download
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = file,
                Arguments = "/SILENT /SUPPRESSMSGBOXES",
                UseShellExecute = true // the installer may ask for administrator rights
            });
            return 0;
        }
        catch
        {
            return -5;
        }
    }

    private void UpdateDownload(string url, string file)
    {
        string part = file + ".download";
        try
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.UserAgent = "WordMat";
            req.Timeout = 30000;
            req.ReadWriteTimeout = 30000;
            using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
            using (Stream s = resp.GetResponseStream())
            using (FileStream f = new FileStream(part, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                long total = resp.ContentLength, done = 0;
                byte[] buf = new byte[81920];
                int n;
                while (!updateCancel && (n = s.Read(buf, 0, buf.Length)) > 0)
                {
                    f.Write(buf, 0, n);
                    done += n;
                    if (total > 0) updateState = (int)Math.Min(99, done * 100 / total);
                }
            }
            if (updateCancel)
            {
                UpdateDelete(part);
                updateState = -3;
                return;
            }
            if (new FileInfo(part).Length < UpdateMinSize)
            {
                UpdateDelete(part);
                updateState = -1;
                return;
            }
            UpdateDelete(file);
            File.Move(part, file);
            if (!UpdateSignatureValid(file))
            {
                UpdateDelete(file);
                updateState = -2;
                return;
            }
            updateFile = file;
            updateState = 100;
        }
        catch
        {
            UpdateDelete(part);
            updateState = updateCancel ? -3 : -1;
        }
    }

    private static void UpdateDelete(string file)
    {
        try { if (File.Exists(file)) File.Delete(file); } catch { }
    }

    // The user's Downloads folder
    private static string UpdateDownloadsFolder()
    {
        try
        {
            IntPtr p;
            if (SHGetKnownFolderPath(new Guid("374DE290-123F-4565-9164-39C4925E467B"), 0, IntPtr.Zero, out p) == 0)
            {
                string s = Marshal.PtrToStringUni(p);
                Marshal.FreeCoTaskMem(p);
                if (!string.IsNullOrEmpty(s) && Directory.Exists(s)) return s;
            }
        }
        catch { }
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
    }

    // True when the file has a valid Authenticode signature (Windows' own check, WinVerifyTrust: the file is unchanged
    // since it was signed, and the certificate chain is trusted) and the signer is UpdateSigner
    private static bool UpdateSignatureValid(string file)
    {
        IntPtr pFile = IntPtr.Zero;
        try
        {
            WINTRUST_FILE_INFO fi = new WINTRUST_FILE_INFO();
            fi.cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_FILE_INFO));
            fi.pcwszFilePath = file;
            pFile = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WINTRUST_FILE_INFO)));
            Marshal.StructureToPtr(fi, pFile, false);
            WINTRUST_DATA wd = new WINTRUST_DATA();
            wd.cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_DATA));
            wd.dwUIChoice = 2;      // WTD_UI_NONE
            wd.dwUnionChoice = 1;   // WTD_CHOICE_FILE
            wd.pFile = pFile;
            wd.dwProvFlags = 0x10;  // WTD_REVOCATION_CHECK_NONE (no extra requests to the internet)
            if (WinVerifyTrust(new IntPtr(-1), new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE"), ref wd) != 0) return false;
            using (var cert = new System.Security.Cryptography.X509Certificates.X509Certificate2(
                System.Security.Cryptography.X509Certificates.X509Certificate.CreateFromSignedFile(file)))
            {
                string name = cert.GetNameInfo(System.Security.Cryptography.X509Certificates.X509NameType.SimpleName, false);
                foreach (string signer in UpdateSigner)
                    if (string.Equals(name, signer, StringComparison.OrdinalIgnoreCase)) return true;
                return false;
            }
        }
        catch
        {
            return false;
        }
        finally
        {
            if (pFile != IntPtr.Zero)
            {
                Marshal.DestroyStructure(pFile, typeof(WINTRUST_FILE_INFO));
                Marshal.FreeHGlobal(pFile);
            }
        }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WINTRUST_FILE_INFO
    {
        public uint cbStruct;
        [MarshalAs(UnmanagedType.LPWStr)] public string pcwszFilePath;
        public IntPtr hFile;
        public IntPtr pgKnownSubject;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WINTRUST_DATA
    {
        public uint cbStruct;
        public IntPtr pPolicyCallbackData;
        public IntPtr pSIPClientData;
        public uint dwUIChoice;
        public uint fdwRevocationChecks;
        public uint dwUnionChoice;
        public IntPtr pFile;
        public uint dwStateAction;
        public IntPtr hWVTStateData;
        public IntPtr pwszURLReference;
        public uint dwProvFlags;
        public uint dwUIContext;
        public IntPtr pSignatureSettings;
    }

    [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
    private static extern int WinVerifyTrust(IntPtr hwnd, [MarshalAs(UnmanagedType.LPStruct)] Guid pgActionID, ref WINTRUST_DATA pWVTData);

    [DllImport("shell32.dll")]
    private static extern int SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);

    // ---------------------------------------------------------------------------------------------------------
    // The WordMat+ license check and the check for a new version (WordMat+: PQActivePartnership in
    // ModulePartnerskab, WMPCheckForUpdate in ModuleHjaelp). Moved here from WordMat+'s VBA code (2026-11-11), so the
    // macros do not read the computer's ids, the Office accounts or the license data, and make no requests to the
    // internet. The caller gives no address: the requests go to WordMat's servers (www.wordmat.dk, www.eduap.com).
    //
    // The license check is a line by line port of the VBA function PQActivePartnership and the functions it used
    // (GetHardwareUUID, GetHardwareUUID2, GetUserEmailFromRegistry, GetRegistryValue, SetRegistryValue, Scramble,
    // UnScramble, GetHTML). It must give the same as the VBA code, which WordMat (WordMatP.dotm) and WordMat+ on Mac
    // still use, with the same license data in the registry (HKCU\Software\WordMat\Settings: MyKeys, TipModCounter,
    // Skolenavn). So:
    //   - the registry is read and written with the same functions as VBA used (advapi32's ANSI functions, in the
    //     registry view of the process: a 32 bit Word does not see MachineGuid, and must not see it here either)
    //   - dates are read and written with the functions VBA uses (oleaut32: VarDateFromStr, VarBstrFromDate), in the
    //     user's date format
    //   - where the VBA code got a run-time error and went to its error handler, an exception is thrown here
    // WordMat+'s tests compare this code with the VBA code (Tests\DllTests.ps1).
    //
    //   LicenseCheckStart(force, userName)   starts the check (userName: Word's user name, the name of a license
    //                                        bought by one person). 0 = started, -4 = a check is already running
    //   LicenseCheckResult()                 -100 = running, -4 = no check was started, otherwise the sum of
    //                                        1 = a license, 2 = not checked now: no license was found less than 7 days
    //                                        ago, 4 = no connection to the internet, 8 = the server gave an error
    //   LicenseName()                        the name of the school (or the user) of the license
    //   LicensePurchasePage(danish)          the page where WordMat+ is bought, in the browser. 0 = opened
    //   VersionCheckStart(partnership)       starts reading WordMat's version file. 0 = started, -4 = already running
    //   VersionCheckResult()                 -100 = running, -4 = not started, 0 = done
    //   VersionCheckText()                   the version file ("version:1.42" and the news), "" if it was not read
    // ---------------------------------------------------------------------------------------------------------
    private const string LicKey = "dkfyenbskdhfnbaozmxzfasebmzasfdmehsbuemcziyaqpeudkgflfdtqmapewiusmcnvhsudifqpoaydtlftzpwiy";
    private const int LicNeedsNetwork = -1000;
    private const int LicLcid = 0x0400;          // LOCALE_USER_DEFAULT: the user's formats, as VBA
    private static readonly IntPtr LicHKCU = new IntPtr(unchecked((int)0x80000001));
    private static readonly IntPtr LicHKLM = new IntPtr(unchecked((int)0x80000002));
    private string licSettings = @"Software\WordMat\Settings"; // (the tests: another key)
    private string licTestAnswers;               // the tests: the servers' answers, no request is made
    private StringBuilder licTestLog;            // the tests: the addresses asked for
    private volatile int licState = -4;
    private volatile string licName = "";
    private Thread licThread;
    private volatile int verState = -4;
    private volatile string verText = "";
    private Thread verThread;

    public int LicenseCheckStart(bool force, string userName)
    {
        if (licThread != null && licThread.IsAlive) return -4;
        string user = userName ?? "";
        licState = -100;
        int r;
        try { r = LicenseRun(force, user, false); } catch { r = 0; }
        if (r != LicNeedsNetwork)
        {
            licState = r; // decided by the license data in the registry: no request
            return 0;
        }
        licThread = new Thread(() =>
        {
            int x;
            try { x = LicenseRun(force, user, true); } catch { x = 0; }
            licState = x;
        });
        licThread.IsBackground = true;
        licThread.Start();
        return 0;
    }

    public int LicenseCheckResult()
    {
        return licState;
    }

    public string LicenseName()
    {
        return licName ?? "";
    }

    public int LicensePurchasePage(bool danish)
    {
        try
        {
            Process.Start(new ProcessStartInfo { FileName = LicPurchaseUrl(danish), UseShellExecute = true });
            return 0;
        }
        catch
        {
            return -5;
        }
    }

    public int VersionCheckStart(bool partnership)
    {
        if (verThread != null && verThread.IsAlive) return -4;
        verText = "";
        verState = -100;
        verThread = new Thread(() =>
        {
            string s = "";
            try
            {
                s = LicVersionFrom("www.wordmat.dk", partnership);
                if (s.Length == 0) s = LicVersionFrom("www.eduap.com", partnership);
            }
            catch { s = ""; }
            verText = s;
            verState = 0;
        });
        verThread.IsBackground = true;
        verThread.Start();
        return 0;
    }

    public int VersionCheckResult()
    {
        return verState;
    }

    public string VersionCheckText()
    {
        return verText ?? "";
    }

    private static string LicPurchaseUrl(bool danish)
    {
        return "https://www.eduap.com/payment/" + (danish ? "da" : "en") + "/purchase.php?ClientID=" + LicHardwareId(1) +
            "&ClientID2=" + LicHardwareId(2);
    }

    // WordMat's version file on a server ("version:1.42.1" and the news), "" if it cannot be read
    private string LicVersionFrom(string server, bool partnership)
    {
        bool answered;
        string s = LicGet("https://" + server + "/api/get-wordmat-version.php?platform=win" + (partnership ? "p" : ""), out answered);
        if (s == "error" || s.Contains("<head>") || s.Contains("404 Not Found")) s = "";
        return s;
    }

    // PQActivePartnership, from where the VBA function had found no answer in its variables (mPartnerSkabOK).
    // network = false: only the license data in the registry; LicNeedsNetwork when the servers must be asked.
    // Returns the value of LicenseCheckResult
    private int LicenseRun(bool force, string userName, bool network)
    {
        string skole = "", glSkole = "", uuid = "", uuid2 = "", s;
        double today = DateTime.Today.ToOADate(), lastCheck = 0, d;
        bool errorOccured = false, noInternet = false, answered1, answered2;
        try
        {
            uuid = LicHardwareId(1);
            uuid2 = LicHardwareId(2);
            s = LicScramble(LicRegGet("MyKeys"), false);
            string[] arr = VbSplit(s, ";");
            string uuidReg = "", uuidReg2 = "";
            if (arr.Length >= 2)
            {
                lastCheck = VbIsDate(arr[0], out d) ? d : 0;
                skole = arr[1];
                glSkole = skole;
                if (arr.Length > 2) uuidReg = arr[2];
                if (arr.Length > 3) uuidReg2 = arr[3];
            }
            if (arr.Length >= 5 && VbIsNumeric(arr[4])) LicRegSet("TipModCounter", arr[4]);
            // the license data is from another computer (or has no id): checked again
            if ((uuidReg.Length == 0 && uuidReg2.Length == 0) || !(uuid == uuidReg || uuid2 == uuidReg2)) lastCheck = 0;
            if (VbDays(lastCheck, today) < 7 && skole.Length == 0 && !force)
            {
                licName = skole;
                return 2; // no license was found less than 7 days ago: the servers are not asked again
            }
            if (VbDays(lastCheck, today) < 30 && skole.Length != 0)
            {
                licName = skole;
                return 1; // a license was found less than 30 days ago
            }
            if (!network) return LicNeedsNetwork;

            // a partnership of the school: by the IP address or the domain of an e-mail address of an Office account
            string emails = "";
            try { emails = LicEmailDomains(LicOfficeEmails()); } catch { emails = ""; }
            s = LicCheckPartnership("www.wordmat.dk", emails, out answered1);
            if (s.Length == 0 || s == "error")
            {
                s = LicCheckPartnership("www.eduap.com", emails, out answered2);
                if (s.Length == 0 || s == "error")
                {
                    if (!answered1 && !answered2)
                    {
                        noInternet = true; // no answer from either server
                        goto slut;
                    }
                    s = "";
                    errorOccured = true;
                }
            }
            if (s.Length > 9)
            {
                s = VbTrim(s.Replace("\n", ""));
                int p = s.IndexOf("Skolenavn:", StringComparison.Ordinal);
                if (p >= 0)
                {
                    s = s.Substring(p + 10).Replace("\r\n", "").Replace("\r", "").Replace("\n", "");
                    skole = VbTrim(s);
                    if (skole.Length != 0)
                    {
                        LicRegSet("MyKeys", LicScramble(VbDateStr(today) + ";" + skole + ";" + uuid + ";" + uuid2 + ";0", true));
                        LicRegSet("Skolenavn", skole);
                        licName = skole;
                        return 1;
                    }
                }
            }

            // WordMat+ bought by one person: the server answers with the date the license runs out
            s = LicCheckPayment("www.eduap.com/payment/checkpaymentmonths.php", uuid, uuid2);
            if (s.Length == 0 || s == "error") s = LicCheckPayment("www.wordmat.dk/api/checkpaymentmonths.php", uuid, uuid2);
            if (s.Length == 0 || s == "error")
            {
                errorOccured = true;
            }
            else if (VbLCase(VbTrim(s)) != "false")
            {
                if (VbIsDate(s, out d))
                {
                    double end = DateTime.FromOADate(d).Date.ToOADate(); // DateValue
                    if (VbMonths(today, end) > 0)
                    {
                        skole = userName;
                        LicRegSet("MyKeys", LicScramble(VbDateStr(end) + ";" + skole + ";" + uuid + ";" + uuid2 + ";0", true));
                        LicRegSet("Skolenavn", skole);
                        licName = skole;
                        return 1;
                    }
                }
            }

            // no license was found, and there is a connection
            short tip = 0;
            string settings = LicGetSettings("wordmat.dk", "mandatory");
            if (settings.Length == 0 || settings == "error") settings = LicGetSettings("eduap.com", "mandatory");
            if (!(settings.Length == 0 || settings == "error"))
            {
                foreach (string line in VbSplit(settings, "\n"))
                {
                    string[] a2 = VbSplit(line, "=");
                    // a line without = is skipped (the file ends with a line break: the empty line after it stopped the
                    // check here until 2026-11-11, before the date below was stored)
                    if (a2.Length < 2) continue;
                    if (VbTrim(a2[0]) == "TipModCounter")
                    {
                        tip = VbCInt(VbTrim(a2[1]));
                        LicRegSet("TipModCounter", tip.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    }
                }
            }
            // the date of the check, when no school is known: the servers are then not asked again for 7 days. With
            // the ids of the computer (since 2026-11-11): a date without ids does not count (see above), so the
            // servers were asked in every Word session
            if (glSkole.Length == 0)
                LicRegSet("MyKeys", LicScramble(VbDateStr(today) + ";;" + uuid + ";" + uuid2 + ";" + tip.ToString(System.Globalization.CultureInfo.InvariantCulture), true));
            LicRegSet("SkoleNavn", "");
        }
        catch
        {
            try { LicRegSet("Skolenavn", ""); } catch { }
        }
    slut:
        // without a license found now: one found less than 60 days ago still holds, and so does one found earlier
        // when the server gave an error
        int r = 0;
        if ((VbDays(lastCheck, today) < 60 || errorOccured) && (skole.Length != 0 || glSkole.Length != 0))
        {
            if (skole.Length == 0) skole = glSkole;
            r = 1;
        }
        else if (errorOccured)
        {
            r = 8;
        }
        if (noInternet) r |= 4;
        licName = skole;
        return r;
    }

    // Scramble (scramble = true) and UnScramble of the VBA code: the license data as it is stored in the registry.
    // The character at a position that is a multiple of the key's length is not changed (VBA: Mid$ from position 0
    // gave an error), nor one whose calculation does not fit a 16 bit integer (VBA: overflow)
    private static string LicScramble(string s, bool scramble)
    {
        StringBuilder sb = new StringBuilder(s.Length);
        for (int i = 1; i <= s.Length; i++)
        {
            char ch = s[i - 1];
            int c = (short)ch; // AscW
            int j = i % LicKey.Length;
            if (j == 0)
            {
                sb.Append(ch);
                continue;
            }
            int k = LicKey[j - 1];
            int v = scramble ? c - k : c + k;
            if (v < short.MinValue || v > short.MaxValue)
            {
                sb.Append(ch);
                continue;
            }
            v = scramble ? v + 97 : v - 97;
            if (v < short.MinValue || v > short.MaxValue)
            {
                sb.Append(ch);
                continue;
            }
            sb.Append((char)(ushort)(short)v); // ChrW$
        }
        return sb.ToString();
    }

    // GetHardwareUUID (1) and GetHardwareUUID2 (2) of the VBA code: the ids of the computer a license is bound to
    private static string LicHardwareId(int number)
    {
        string s = "";
        try
        {
            if (number == 1)
                s = LicRegRead(LicHKLM, @"SYSTEM\CurrentControlSet\Control\IDConfigDB\Hardware Profiles\0001", "HwProfileGuid");
            else
                s = LicRegRead(LicHKLM, @"SOFTWARE\Microsoft\Cryptography", "MachineGuid");
        }
        catch { s = ""; }
        s = VbTrim(s);
        if (s.StartsWith("{", StringComparison.Ordinal)) s = s.Substring(1);
        if (s.EndsWith("}", StringComparison.Ordinal)) s = s.Substring(0, s.Length - 1);
        return s;
    }

    // GetUserEmailFromRegistry and the ADUserName of the VBA code: the e-mail addresses of the Office accounts
    private static List<string> LicOfficeEmails()
    {
        List<string> c = new List<string>();
        string[] keys = {
            @"Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext\LicenseIdToEmailMapping",
            @"Software\Microsoft\Office\16.0\Common\Identity\IdToAuthorityUrlMapping\EmailAddress",
            @"Software\Microsoft\Office\16.0\Common\Identity\Identities" };
        foreach (string key in keys)
        {
            try { LicRegListData(key, c); } catch { }
        }
        try
        {
            string email = LicRegRead(LicHKCU, @"Software\Microsoft\OneDrive\Accounts\Business1", "UserEmail");
            if (email.Length != 0) c.Add(email);
        }
        catch { }
        string ad = LicRegRead(LicHKCU, @"Software\Microsoft\Office\16.0\Common\Identity", "ADUserName");
        if (ad.Length != 0) c.Add(ad);
        return c;
    }

    // EmailDomains of the VBA code: the domains of the addresses (each once, lower case, URL-encoded), separated by
    // commas. Only the domains are sent to the server
    private static string LicEmailDomains(List<string> emails)
    {
        string result = "";
        foreach (string email in emails)
        {
            int p = email.LastIndexOf('@');
            if (p >= 0)
            {
                string dom = LicUrlEncode(VbLCase(VbTrim(email.Substring(p + 1))));
                if (dom.Length != 0 && ("," + result + ",").IndexOf("," + dom + ",", StringComparison.Ordinal) < 0)
                {
                    if (result.Length != 0) result += ",";
                    result += dom;
                }
            }
        }
        return result;
    }

    // UrlEncode of the VBA code: letters, digits and - . _ ~ as they are, other characters as %XX of their UTF-8 bytes
    private static string LicUrlEncode(string s)
    {
        StringBuilder sb = new StringBuilder();
        foreach (char ch in s)
        {
            int c = ch;
            if ((c >= 48 && c <= 57) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c == 45 || c == 46 || c == 95 || c == 126)
                sb.Append(ch);
            else if (c < 0x80)
                sb.Append('%').Append(c.ToString("X2"));
            else if (c < 0x800)
                sb.Append('%').Append((0xC0 | (c >> 6)).ToString("X2")).Append('%').Append((0x80 | (c & 0x3F)).ToString("X2"));
            else
                sb.Append('%').Append((0xE0 | (c >> 12)).ToString("X2")).Append('%').Append((0x80 | ((c >> 6) & 0x3F)).ToString("X2"))
                    .Append('%').Append((0x80 | (c & 0x3F)).ToString("X2"));
        }
        return sb.ToString();
    }

    // CheckPartnership of the VBA code
    private string LicCheckPartnership(string server, string emails, out bool answered)
    {
        string s = LicGet("https://" + server + "/api/check-partnerskab.php?emails=" + emails, out answered);
        if (s.Length == 0 || s.Contains("<head>") || VbLCase(s).Contains("error") || s.Contains("404 Not Found")) s = "error";
        return s;
    }

    // GetSettingsFromServer of the VBA code
    private string LicGetSettings(string server, string settingFile)
    {
        bool answered;
        string s = LicGet("https://" + server + "/api/get-settings.php?skolenavn=" + settingFile, out answered);
        if (s.Length == 0 || s.Contains("<head>") || VbLCase(s).Contains("error") || s.Contains("404 Not Found")) s = "error";
        return s;
    }

    // CheckWordMatPlusFromServerAPI of the VBA code
    private string LicCheckPayment(string endpoint, string uuid, string uuid2)
    {
        bool answered;
        string s = LicGet("https://" + endpoint + "?ClientID=" + uuid + "&ClientID2=" + uuid2, out answered);
        if (s.Length == 0 || s.Contains("<head>") || s.Contains("error") || s.Contains("404 Not Found")) s = "error";
        return s;
    }

    // GetHTML of the VBA code: the text at url, "error" if it cannot be read (no answer within 10 seconds, or a status
    // other than 200). answered: a server answered (any status); False: no connection. Windows' proxy settings are
    // used, as before
    private string LicGet(string url, out bool answered)
    {
        answered = false;
        if (licTestAnswers != null)
        {
            if (licTestLog != null) licTestLog.Append(url).Append('\n');
            foreach (string entry in licTestAnswers.Split('\u0001'))
            {
                string[] f = entry.Split('\u0002');
                if (f.Length == 3 && url.Contains(f[0]))
                {
                    answered = (f[1] == "1");
                    return f[2];
                }
            }
            return "error";
        }
        try
        {
            // the time in the address makes sure it is not a stored answer
            long cb = (long)Math.Round(DateTime.Now.TimeOfDay.TotalSeconds * 100);
            url += (url.IndexOf('?') >= 0 ? "&cb=" : "?cb=") + cb.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.UserAgent = "WordMat";
            req.Timeout = 10000;
            req.ReadWriteTimeout = 10000;
            try
            {
                IWebProxy proxy = WebRequest.GetSystemWebProxy();
                proxy.Credentials = CredentialCache.DefaultCredentials; // a proxy that asks for the Windows user, as before
                req.Proxy = proxy;
            }
            catch { }
            using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
            {
                answered = true;
                if (resp.StatusCode != HttpStatusCode.OK) return "error";
                Encoding enc = Encoding.UTF8;
                string ct = resp.ContentType ?? "";
                int i = ct.IndexOf("charset=", StringComparison.OrdinalIgnoreCase);
                if (i >= 0)
                {
                    string name = ct.Substring(i + 8);
                    int semi = name.IndexOf(';');
                    if (semi >= 0) name = name.Substring(0, semi);
                    try { enc = Encoding.GetEncoding(name.Trim(' ', '"', '\'')); } catch { enc = Encoding.UTF8; }
                }
                using (Stream st = resp.GetResponseStream())
                using (StreamReader rd = new StreamReader(st, enc, true))
                {
                    char[] buf = new char[1000000]; // (the answers are a few lines)
                    int n = 0, k;
                    while (n < buf.Length && (k = rd.Read(buf, n, buf.Length - n)) > 0) n += k;
                    return new string(buf, 0, n);
                }
            }
        }
        catch (WebException e)
        {
            if (e.Response != null) answered = true;
            return "error";
        }
        catch
        {
            return "error";
        }
    }

    // ---- the registry, as the VBA code read and wrote it (GetRegistryValue, SetRegistryValue, GetReg, SetReg) ----
    private string LicRegGet(string name)
    {
        return LicRegRead(LicHKCU, licSettings, name);
    }

    private void LicRegSet(string name, string value)
    {
        IntPtr h;
        int disp;
        if (LicRegCreateKey(LicHKCU, licSettings, 0, null, 0, 0x20006, IntPtr.Zero, out h, out disp) != 0) return;
        // the text in the ANSI code page, converted as VBA converts a text it gives to a Windows function: a character
        // that is not in the code page becomes the one that looks most like it (Marshal.StringToHGlobalAnsi makes it ?)
        string v = value ?? "";
        IntPtr p = IntPtr.Zero;
        try
        {
            int n = v.Length == 0 ? 0 : LicWideCharToMultiByte(0, 0, v, v.Length, IntPtr.Zero, 0, IntPtr.Zero, IntPtr.Zero);
            p = Marshal.AllocHGlobal(n + 1);
            if (n > 0) LicWideCharToMultiByte(0, 0, v, v.Length, p, n, IntPtr.Zero, IntPtr.Zero);
            Marshal.WriteByte(p, n, 0);
            LicRegSetValue(h, name, 0, 1, p, n + 1); // REG_SZ, with its null character
        }
        finally
        {
            if (p != IntPtr.Zero) Marshal.FreeHGlobal(p);
            LicRegCloseKey(h);
        }
    }

    // GetRegistryValue: a text (REG_SZ), a number (REG_DWORD) as text, binary data as hex, "" when it is not there
    private static string LicRegRead(IntPtr root, string path, string name)
    {
        IntPtr h;
        if (LicRegOpenKey(root, path, 0, 0x20019, out h) != 0) return "";
        IntPtr buf = IntPtr.Zero;
        try
        {
            int type, size = 0;
            if (LicRegQueryValue(h, name, IntPtr.Zero, out type, IntPtr.Zero, ref size) != 0) return "";
            if (size <= 0) throw new InvalidOperationException(); // (VBA: ReDim of an empty array)
            int alloc = Math.Max(size, 4);
            buf = Marshal.AllocHGlobal(alloc);
            for (int i = 0; i < alloc; i++) Marshal.WriteByte(buf, i, 0);
            int size2 = size;
            if (type == 1) // REG_SZ: the text up to its null character, converted from the ANSI code page
            {
                LicRegQueryValue(h, name, IntPtr.Zero, out type, buf, ref size2);
                int n = 0;
                while (n < size && Marshal.ReadByte(buf, n) != 0) n++;
                if (n >= size) throw new InvalidOperationException(); // (VBA: Left$ with -1)
                return Marshal.PtrToStringAnsi(buf, n);
            }
            if (type == 4) // REG_DWORD
            {
                size2 = 4;
                LicRegQueryValue(h, name, IntPtr.Zero, out type, buf, ref size2);
                return Marshal.ReadInt32(buf).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            if (type == 3) // REG_BINARY
            {
                LicRegQueryValue(h, name, IntPtr.Zero, out type, buf, ref size2);
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < size2 && i < alloc; i++) sb.Append(Marshal.ReadByte(buf, i).ToString("X2")).Append(' ');
                return VbTrim(sb.ToString());
            }
            return "";
        }
        finally
        {
            if (buf != IntPtr.Zero) Marshal.FreeHGlobal(buf);
            LicRegCloseKey(h);
        }
    }

    // ListRegistryValueNamesAndData (the data only): the values of a key of HKEY_CURRENT_USER
    private static void LicRegListData(string keyPath, List<string> coll)
    {
        IntPtr h;
        if (LicRegOpenKey(LicHKCU, keyPath, 0, 0x20019, out h) != 0) return;
        IntPtr pName = Marshal.AllocHGlobal(16384), pData = Marshal.AllocHGlobal(1024);
        try
        {
            for (int i = 0; i < 1024; i++) Marshal.WriteByte(pData, i, 0);
            int index = 0, r;
            do
            {
                int nameLen = 16383, dataLen = 1024, type;
                r = LicRegEnumValue(h, index, pName, ref nameLen, IntPtr.Zero, out type, pData, ref dataLen);
                if (r == 0)
                {
                    if (type == 1)
                    {
                        int n = 0;
                        while (n < 1024 && Marshal.ReadByte(pData, n) != 0) n++;
                        if (n >= 1024) throw new InvalidOperationException();
                        coll.Add(Marshal.PtrToStringAnsi(pData, n));
                    }
                    else if (type == 4)
                    {
                        coll.Add(Marshal.ReadInt32(pData).ToString(System.Globalization.CultureInfo.InvariantCulture));
                    }
                    index++;
                }
            } while (r == 0); // (a value longer than the buffer ends the list, as in the VBA code)
        }
        finally
        {
            Marshal.FreeHGlobal(pName);
            Marshal.FreeHGlobal(pData);
            LicRegCloseKey(h);
        }
    }

    // ---- VBA's functions ----
    private static string[] VbSplit(string s, string separator)
    {
        if (s.Length == 0) return new string[0]; // (Split of an empty text is an empty array)
        return s.Split(new string[] { separator }, StringSplitOptions.None);
    }

    private static string VbTrim(string s)
    {
        return s.Trim(' ');
    }

    private static string VbLCase(string s)
    {
        return s.ToLower(System.Globalization.CultureInfo.CurrentCulture);
    }

    private static bool VbIsDate(string s, out double date)
    {
        return LicVarDateFromStr(s, LicLcid, 0, out date) == 0;
    }

    private static string VbDateStr(double date)
    {
        string s;
        if (LicVarBstrFromDate(date, LicLcid, 0, out s) != 0) throw new InvalidOperationException();
        return s;
    }

    private static bool VbIsNumeric(string s)
    {
        double x;
        return LicVarR8FromStr(s, LicLcid, 0, out x) == 0;
    }

    private static short VbCInt(string s)
    {
        short x;
        if (LicVarI2FromStr(s, LicLcid, 0, out x) != 0) throw new InvalidOperationException(); // (VBA: type mismatch)
        return x;
    }

    // DateDiff("d", from, to)
    private static int VbDays(double from, double to)
    {
        return (int)(DateTime.FromOADate(to).Date - DateTime.FromOADate(from).Date).TotalDays;
    }

    // DateDiff("m", from, to)
    private static int VbMonths(double from, double to)
    {
        DateTime a = DateTime.FromOADate(from), b = DateTime.FromOADate(to);
        return (b.Year - a.Year) * 12 + b.Month - a.Month;
    }

    [DllImport("advapi32.dll", EntryPoint = "RegOpenKeyExA", CharSet = CharSet.Ansi)]
    private static extern int LicRegOpenKey(IntPtr hKey, string subKey, int options, int sam, out IntPtr result);

    [DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA", CharSet = CharSet.Ansi)]
    private static extern int LicRegQueryValue(IntPtr hKey, string name, IntPtr reserved, out int type, IntPtr data, ref int size);

    [DllImport("advapi32.dll", EntryPoint = "RegCreateKeyExA", CharSet = CharSet.Ansi)]
    private static extern int LicRegCreateKey(IntPtr hKey, string subKey, int reserved, string cls, int options, int sam, IntPtr security, out IntPtr result, out int disposition);

    [DllImport("advapi32.dll", EntryPoint = "RegSetValueExA", CharSet = CharSet.Ansi)]
    private static extern int LicRegSetValue(IntPtr hKey, string name, int reserved, int type, IntPtr data, int size);

    [DllImport("advapi32.dll", EntryPoint = "RegEnumValueA", CharSet = CharSet.Ansi)]
    private static extern int LicRegEnumValue(IntPtr hKey, int index, IntPtr name, ref int nameLen, IntPtr reserved, out int type, IntPtr data, ref int size);

    [DllImport("advapi32.dll", EntryPoint = "RegCloseKey")]
    private static extern int LicRegCloseKey(IntPtr hKey);

    [DllImport("kernel32.dll", EntryPoint = "WideCharToMultiByte")]
    private static extern int LicWideCharToMultiByte(uint codePage, uint flags, [MarshalAs(UnmanagedType.LPWStr)] string wide, int wideLength, IntPtr multi, int multiLength, IntPtr defaultChar, IntPtr usedDefaultChar);

    [DllImport("oleaut32.dll", EntryPoint = "VarDateFromStr", CharSet = CharSet.Unicode)]
    private static extern int LicVarDateFromStr(string text, int lcid, int flags, out double date);

    [DllImport("oleaut32.dll", EntryPoint = "VarBstrFromDate")]
    private static extern int LicVarBstrFromDate(double date, int lcid, int flags, [MarshalAs(UnmanagedType.BStr)] out string text);

    [DllImport("oleaut32.dll", EntryPoint = "VarR8FromStr", CharSet = CharSet.Unicode)]
    private static extern int LicVarR8FromStr(string text, int lcid, int flags, out double value);

    [DllImport("oleaut32.dll", EntryPoint = "VarI2FromStr", CharSet = CharSet.Unicode)]
    private static extern int LicVarI2FromStr(string text, int lcid, int flags, out short value);

    // ---------------------------------------------------------------------------------------------------------
    // Windows functions WordMat+'s VBA code declared until 2026-11-11 (Declare of kernel32, user32 and gdi32: Defender's
    // rule "Block Win32 API calls from Office macros" looks at them). The same calls, so the same results:
    //   Pause(ms)                 kernel32's Sleep (ModuleWindowsFunctions.Sleep): waits without using the processor
    //   ScreenInfo(hwnd)          the monitor of the window (MonitorFromWindow, nearest; GetMonitorInfo) and the screen's
    //                             dots per inch (GetDeviceCaps of the window's device context; of the screen when hwnd
    //                             is 0): "monitor left;monitor top;monitor right;work area bottom;dpi x;dpi y" in pixels
    //                             (ModuleWindowsFunctions.GetWorkArea and ConvertX/YPixelsToPoints), 0 when not found
    //   WindowToFront(class)      the first window of the class to the front (FindWindow, SetForegroundWindow; Excel's
    //                             XLMAIN, Excelsubs.SetExcelForeground): 1 = found
    //   HideConsoleWindow()       the first console window (ConsoleWindowClass, else Windows Terminal's
    //                             CASCADIA_HOSTING_WINDOW_CLASS) hidden (VBACASmaxima.HideConsoleWindow, Maxima through
    //                             WSH): 1 = a window was hidden
    // ---------------------------------------------------------------------------------------------------------
    public void Pause(int milliseconds)
    {
        if (milliseconds > 0) Thread.Sleep(milliseconds);
    }

    public string ScreenInfo(int hwnd)
    {
        IntPtr h = new IntPtr(hwnd);
        int left = 0, top = 0, right = 0, bottom = 0, dpiX = 0, dpiY = 0;
        try
        {
            IntPtr mon = WinMonitorFromWindow(h, 2); // MONITOR_DEFAULTTONEAREST
            WIN_MONITORINFO mi = new WIN_MONITORINFO();
            mi.cbSize = Marshal.SizeOf(typeof(WIN_MONITORINFO));
            if (WinGetMonitorInfo(mon, ref mi))
            {
                left = mi.rcMonitor.left;
                top = mi.rcMonitor.top;
                right = mi.rcMonitor.right;
                bottom = mi.rcWork.bottom;
            }
        }
        catch { }
        try
        {
            IntPtr dc = WinGetDC(h);
            if (dc == IntPtr.Zero && h != IntPtr.Zero)
            {
                h = IntPtr.Zero; // (as the VBA code: the screen when the window's could not be had)
                dc = WinGetDC(h);
            }
            dpiX = WinGetDeviceCaps(dc, 88); // LOGPIXELSX
            dpiY = WinGetDeviceCaps(dc, 90); // LOGPIXELSY
            WinReleaseDC(h, dc);
        }
        catch { }
        return left + ";" + top + ";" + right + ";" + bottom + ";" + dpiX + ";" + dpiY;
    }

    public int WindowToFront(string className)
    {
        try
        {
            IntPtr w = WinFindWindow(className, null);
            if (w == IntPtr.Zero) return 0;
            WinSetForegroundWindow(w);
            return 1;
        }
        catch
        {
            return 0;
        }
    }

    public int HideConsoleWindow()
    {
        try
        {
            IntPtr w = WinFindWindowEx(IntPtr.Zero, IntPtr.Zero, "ConsoleWindowClass", null);
            if (w != IntPtr.Zero && WinShowWindow(w, 0)) return 1; // (0 = hide; true: it was visible)
            w = WinFindWindowEx(IntPtr.Zero, IntPtr.Zero, "CASCADIA_HOSTING_WINDOW_CLASS", null);
            if (w != IntPtr.Zero)
            {
                WinShowWindow(w, 0);
                return 1;
            }
        }
        catch { }
        return 0;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WIN_RECT
    {
        public int left, top, right, bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WIN_MONITORINFO
    {
        public int cbSize;
        public WIN_RECT rcMonitor;
        public WIN_RECT rcWork;
        public uint dwFlags;
    }

    [DllImport("user32.dll", EntryPoint = "MonitorFromWindow")]
    private static extern IntPtr WinMonitorFromWindow(IntPtr hwnd, uint flags);

    [DllImport("user32.dll", EntryPoint = "GetMonitorInfoW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool WinGetMonitorInfo(IntPtr monitor, ref WIN_MONITORINFO info);

    [DllImport("user32.dll", EntryPoint = "GetDC")]
    private static extern IntPtr WinGetDC(IntPtr hwnd);

    [DllImport("user32.dll", EntryPoint = "ReleaseDC")]
    private static extern int WinReleaseDC(IntPtr hwnd, IntPtr dc);

    [DllImport("gdi32.dll", EntryPoint = "GetDeviceCaps")]
    private static extern int WinGetDeviceCaps(IntPtr dc, int index);

    [DllImport("user32.dll", EntryPoint = "FindWindowW", CharSet = CharSet.Unicode)]
    private static extern IntPtr WinFindWindow(string className, string windowName);

    [DllImport("user32.dll", EntryPoint = "FindWindowExW", CharSet = CharSet.Unicode)]
    private static extern IntPtr WinFindWindowEx(IntPtr parent, IntPtr childAfter, string className, string windowName);

    [DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool WinSetForegroundWindow(IntPtr hwnd);

    [DllImport("user32.dll", EntryPoint = "ShowWindow")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool WinShowWindow(IntPtr hwnd, int cmdShow);

    // a function to read a key from the regisry
    public string RegKeyRead(string fullPath)
    {
        string value = "";
        try
        {
            string keyName = Path.GetDirectoryName(fullPath);
            string keyValue = Path.GetFileName(fullPath);
            object regValue = Microsoft.Win32.Registry.GetValue(keyName, keyValue, null);
            if (regValue != null)
            {
                value = regValue.ToString();
            }
            else
            {
                value = "";
            }
        }
        catch
        {
            value = "";
        }
        return value;
    }

    // function to save a key to the registry
    public void RegKeySave(string fullPath, object value)
    {
        try
        {
            string keyName = Path.GetDirectoryName(fullPath);
            string keyValue = Path.GetFileName(fullPath);
            if (value is int || value is string)
            {
                Microsoft.Win32.Registry.SetValue(keyName, keyValue, value);
            }
            else
            {
                throw new ArgumentException("Value must be either a string or an integer.");
            }
        }
        catch
        {
            // do nothing
        }
    }

    //function to delete a key from the registry
    public void RegKeyDelete(string fullPath)
    {
        try
        {
            string keyName = Path.GetDirectoryName(fullPath);
            string keyValue = Path.GetFileName(fullPath);
            Microsoft.Win32.Registry.SetValue(keyName, keyValue, null);
        }
        catch
        {
            // do nothing
        }
    }

    //function to check if a key exists in the registry
    public bool RegKeyExists(string fullPath)
    {
        try
        {
            string keyName = Path.GetDirectoryName(fullPath);
            string keyValue = Path.GetFileName(fullPath);
            object regValue = Microsoft.Win32.Registry.GetValue(keyName, keyValue, null);
            if (regValue != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        catch
        {
            return false;
        }
    }

    // A function that will open an explorer folder in the specied path
    public void OpenFolder(string path)
    {
        Process.Start("explorer.exe", path);
    }

    // A function that will create a zipfile, containing only 1 file, in the specified path
    public void CreateZipFile(string zipPath, string filePath)
    {
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        {
            using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
            {
                ZipArchiveEntry entry = archive.CreateEntry(Path.GetFileName(filePath));
                using (FileStream fileToCompress = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (Stream entryStream = entry.Open())
                    {
                        fileToCompress.CopyTo(entryStream);
                    }
                }
            }
        }
    }



    public void StartMaximaProcess()
            {
                // Initialize the process and its StartInfo properties.

                string maximasti;
                numOutputLines = 0;
                KommandoNr = 0;
                mindex = 0;
                errCode = 0;
                LastWasOutput = false;
                outputstarted = false;
                outputstarted2 = false;
                try
                {
                    maximaProcess = new Process();
            //maxima aldrig 64                    maximasti = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            if (MaximaPath == "")
            {
                try
                {
                    maximasti = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                    DirectoryInfo di = new DirectoryInfo(maximasti);
                    di = di.GetDirectories(@"WordMat\maxima*", SearchOption.TopDirectoryOnly)[di.GetDirectories(@"WordMat\maxima*", SearchOption.TopDirectoryOnly).Length - 1];
                    maximasti = di.FullName;
                }
                //if (di.FullName == "")
                catch
                {
                    try
                    {
                        maximasti = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                        DirectoryInfo di = new DirectoryInfo(maximasti);
                        di = di.GetDirectories(@"WordMat\maxima*", SearchOption.TopDirectoryOnly)[di.GetDirectories(@"WordMat\maxima*", SearchOption.TopDirectoryOnly).Length - 1];
                        maximasti = di.FullName;
                    }
                    catch
                    {
                        maximasti = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                        DirectoryInfo di = new DirectoryInfo(maximasti);
                        di = di.GetDirectories("maxima*", SearchOption.TopDirectoryOnly)[di.GetDirectories("maxima*", SearchOption.TopDirectoryOnly).Length - 1];
                        maximasti = di.FullName;
                    }
                }
            }
            else
            {
                maximasti = MaximaPath;
            }
                    if (units == 1)
                    {
                        maximasti = maximasti + @"\bin\maximaunit.bat";
                    }
                    else 
                    {
                        maximasti = maximasti + @"\bin\maxima.bat";
                    }


                    maximaProcess.StartInfo.FileName = maximasti;
                }
                catch
                {
                    errCode = 1; // maxima not installed  
                    maximaProcess = null;
                    return;
                }

                try
                {
                    MaximaOutputBuilder = new StringBuilder("");
                    LastMaximaOutputBuilder = new StringBuilder("");
                    finish = false;
                    outputstarted = false;
                    outputstarted2 = false;

                    // Set UseShellExecute to false for redirection.
                    maximaProcess.StartInfo.UseShellExecute = false;
                    maximaProcess.StartInfo.CreateNoWindow = true;
                    // Redirect the standard output  
                    // This stream is read asynchronously using an event handler.
                    maximaProcess.StartInfo.RedirectStandardOutput = true;

                    // Set our event handler to asynchronously read the output.
                    maximaProcess.OutputDataReceived += new DataReceivedEventHandler(MaximaOutputHandler);

                    // Redirect standard input as well.  This stream
                    // is used synchronously.
                    maximaProcess.StartInfo.RedirectStandardInput = true;

                    // Start the process.
                    maximaProcess.Start();

                    // Start the asynchronous read of the output stream.
                    maximaProcess.BeginOutputReadLine();

                    // Use a stream writer to synchronously write the input.
                    maximaStreamWriter = maximaProcess.StandardInput;



                    WaitForOutput(); // kan måske flyttes senere hen. Burde ikke være nødv. da executemaxima venter til finish=true

                    // startindstillinger mm. sættes
                    StringBuilder initialsetup = new StringBuilder("[");
                    //initialsetup.Append("load(to_poly_solver),load(solvereal),load(draw),"); to_poly_solve er nu i solvereal
//                    initialsetup.Append("load(solvereal),load(draw),"); // er nu i Maxima image

// Er nu i maximaunit.exe image
//                    if (units == 1)
//                    {
//                        initialsetup.Append("usersetunits:[N,J,W,Pa,C,V,F,Ohm,T,H,K],load(unit),");
//                    }

                    initialsetup.Append(SettingsString);
//                    initialsetup.Append("," + FuncDefinitions);  // er nu i solvereal.mac
                    // tre linjer flyttet til settingsstring
                    initialsetup.Append("]$");

                    ExecuteMaximaCommand(initialsetup.ToString(), 0); // settings der altid skal sættes
            isMaximaStarted = true;
        }
                catch 
                {
                    errCode = 2; // fejl ved kørsel af process  
                    maximaProcess = null;
                    return;
                }
            }

        // alle basisindstillinger
            private string SettingsString
            {
                get 
                {
                    StringBuilder initialsetup = new StringBuilder();
                    // nu i solvereal.mac
//                    initialsetup.Append("display2d:false,ratprint:false,");
//                    initialsetup.Append("numer:false,"); //alt numer håndteres med ,numer
//                    initialsetup.Append("dotscrules:true,"); // prikprodukt -> * når skalar indgår
//                    initialsetup.Append("numer_pbranch:true,"); //(-2)^0,75 konverteres da til imaginærtal. Må gøre det lidt nemmere at identificere imag løsninger
//                    initialsetup.Append("ratsimpexpons:true,"); //simplificer også eksponenter
//                    initialsetup.Append("assumescalar:all,"); //variable antages at være skalare eller må declare(x,nonscalar) og senere kill(x)
//                    initialsetup.Append("logabs:true,"); //integrate(1/x,x)=log(abs(x))
//                    initialsetup.Append("breakup:false,"); //3. og 4. gradspolynomier deles ikke op i underudtryk
//                    initialsetup.Append("algebraic:false,"); // false er default value, simplifikation af algebraiske heltal bl.a. noget med %i. Udtryk der er relle bliver ikke reduceret så %i fjernes helt hvis denne er false
//                    initialsetup.Append("logexpand:true,"); //alle log regneregler   true: log(a^x)=x*log(a)  all: log(a*b)=log(a)+log(b)  super: log(a/b)=log(a)-log(b)
//                    initialsetup.Append("triginverses:all,"); //noget med sin^-1(sin(x))=x altid og begge veje
//                    //                    initialsetup.Append("erfflag:true,"); // true er default. med false vises ikke erf funktion i udtryk fra integraler, men så regnes følgende forkert qua_qagi(normalfordeling,x,minf,10) = -1/2 istedet for 1/2 . men Integrate(0.132980760370601*x^1.5*%e^(-x/2),x,0,1500);  %,numer giver 0 istedet for 1 hvis ikke sat til false. dilemma?
//                    initialsetup.Append("intanalysis:false,"); // Kan løse flere bestemte integraler. Noget med ikke at søge efter poler i området og altid integrere. 
//                    initialsetup.Append("listconstvars:true,"); // for at %i også kommer med blandt variable så kompleks kan findes
//                    //                    initialsetup.Appelistconstvars:truend("matchdeclare (iszeroqwz, lambda ([qwz], qwz =0)),tellsimp(iszeroqwz^iszeroqwz,1),"); // def af 0^0=1 gav problemer ved integration af normalfordeling så nu i funktion: mu:2;sigma:1;Integrate(1/(sqrt(2*%pi)*sigma)*%e^(-1/2*((y-mu)/sigma)^2),y,(minf),1000);
//                    //test: [domain:real,display2d:false,ratprint:false,dotscrules:true,numer_pbranch:true,ratsimpexpons:true,assumescalar:all,logabs:true,breakup:false,algebraic:false,logexpand:true,triginverses:all,intanalysis:false,listconstvars:true,ev(tellsimp(0 ^ 0, 1), simp: false)]
//                    //                initialsetup.Append("radexpand:true,"); // default true, styrer reduktion af rødder. hvis all er sqrt(x^2)=x hvis true sqrt(x^2)=abs(x) ' all giver problemer ved reduktion af imaginære udtryk til reelle
//                    if (units == 1)
//                    {
//                        initialsetup.Append("keepfloat:false,"); // keepfloat=true giver problemer for nogle enheds beregninger. Der meldes en fejl.
//                        initialsetup.Append("setunits([N,J,W,Pa,C,V,F,Ohm,T,H,K]),");
//                        if (moutunits.Length > 0)
//                        {
//                            string[] outunits;
//                            outunits = moutunits.Split(",".ToArray());
//                            foreach (string unit in outunits)
//                            {
//                                initialsetup.Append("errcatch(setunits(");
//                                initialsetup.Append(unit);
//                                initialsetup.Append(")),");
//                            }
//                        }
//                    }
//                    else
//                        initialsetup.Append("keepfloat:true,"); // sørger for at der ikke omdannes fra decimaltal til brøk ved eksakt når det er nødvendigt at bruge numeriske metoder

//                    // FLyttet til slovereal.mac
////                    initialsetup.Append(@"errcatch(remrule(""^"",all)),ev(tellsimp(0 ^ 0, 1), simp: false),"); // def af 0^0=1 gav problemer ved integration af normalfordeling så nu i funktion: mu:2;sigma:1;Integrate(1/(sqrt(2*%pi)*sigma)*%e^(-1/2*((y-mu)/sigma)^2),y,(minf),1000);
////                    //fjernet alizvisnot til scinotrul da f.eks 3/2*kg så ikke blev til 1.5*kg        matchdeclare(alizvidnot,lambda([z],if numberp(z) then (if abs(z)<1 or abs(z)>10 then true) else false ))
////                    initialsetup.Append("matchdeclare(aliz,all),matchdeclare(aliznum,numberp),matchdeclare(alizfloat,floatnump),matchdeclare(aliz2,lambda([z],not(numberp(z)))),defrule(scinotrul,aliznum,scinot(aliznum)),defrule(dectalrul,aliznum,dectal(aliznum)),defrule(floatrul,alizfloat,dectal(alizfloat)),defrule(msinrul,sin(aliz),sing(180/%pi*aliz)),defrule(mcosrul,cos(aliz),cosg(180/%pi*aliz)),defrule(mtanrul,tan(aliz),tang(180/%pi*aliz)),defrule(masinrul,asin(aliz),%pi/180*asing(aliz)),defrule(macosrul,acos(aliz),%pi/180*acosg(aliz)),defrule(matanrul,atan(aliz),%pi/180*atang(aliz)),defrule(mlogrul,log(aliz),lgog(aliz)/lgog(%e)),defrule(merul,%e^aliz2,10^(lgog(%e)*aliz2))");
////                    initialsetup.Append(@",defrule(unitrule,aliz,convert(aliz,[])),applyunitrule(qudt):=(applyb1(qudt,unitrule))");
////                    initialsetup.Append(@",dectalallNum(udt):=block([qout],udt:ev(udt,lgog(x):=if x=%e then 0.4342944819032518 elseif numberp(x) then ev(log(x)/log(10),numer) else 'lgog(x),numer),qout:errcatch(apply1(udt,dectalrul)),if qout=[] then udt else qout[1])");
////                    initialsetup.Append(@",dectalall(udt):=block([qout],qout:errcatch(apply1(udt,floatrul)),if qout=[] then udt else qout[1]),");

//                    initialsetup.Append(GuiSettingsString);
                    
                    return initialsetup.ToString();
                }
            }

            private string FuncDefinitions
            {
                get 
                {
                    StringBuilder initialsetup = new StringBuilder();
                    // ikke-bruger funktioner der skal defineres

                    // vektorprodukt, længde af vektor, tværvektor
                    initialsetup.Append(@"infix(""~"",135,135),infix(""×"",135,135)]$[""~""(qwa,qwb):=if not(scalarp(qwa) or scalarp(qwb)) then (if not(matrixp(qwa) and matrixp(qwb)) then qwa×qwb elseif length(qwa)=3 and length(qwb)=3 then matrix( qwa[2]*qwb[3]-qwa[3]*qwb[2], qwa[3]*qwb[1]-qwa[1]*qwb[3], qwa[1]*qwb[2]-qwa[2]*qwb[1]) else(error(""vektorproduktet er kun defineret for vektorer af dimension 3""))) else qwa*qwb");
                    initialsetup.Append(@",infix(""¤"",115,115),""¤""(qwa,qwb):=block(if scalarp(qwa) or scalarp(qwb) then return(qwa*qwb) else return(qwa.qwb))"); // matrix multiplikationsoperatoren har forkert binding power her rettet ved at indføre ny operator. 2/3 . matrix([2],[3]) nu bedre 
                    initialsetup.Append(@",vecmag(qwa):=if nonscalarp(qwa) then (if matrixp(qwa) then sqrt(conjugate(transpose(qwa)) . qwa) else 'vecmag(qwa)) else abs(qwa)");
                    initialsetup.Append(@",tvvec(qwa):=if scalarp(qwa) then error(""Tværvektor er ikke defineret for en skalar"") else (if length(qwa)=2 then matrix(-qwa[2],qwa[1]) else error(""Tværvektor er kun defineret to-dimensionelle vektorer""))");

                    //                    initialsetup.Append("HasInvTrig(udt):=block([i,del],partswitch:true,i:0,start,del:inpart(udt,i),if del=end then return (false),if del=asin then return (true),if del=acos then return (true),if del=atan then return (true), i:i+1, go (start))");
//                    initialsetup.Append("ConvertToDegr2(udt):=block([aliz],matchdeclare(aliz,all),tellsimp(sin(aliz),sing(180/%pi*aliz)),tellsimp(cos(aliz),cosg(180/%pi*aliz)),tellsimp(tan(aliz),tang(180/%pi*aliz)),tellsimp(asin(aliz),%pi/180*asing(aliz)),tellsimp(acos(aliz),%pi/180*acosg(aliz)),tellsimp(atan(aliz),%pi/180*atang(aliz)),slut,udt:ev(udt,simp),remrule(sin,all),remrule(cos,all),remrule(tan,all),remrule(asin,all),remrule(acos,all),remrule(atan,all),udt)"); // clear_rules() fjernet
                    initialsetup.Append(",ConvertToDegr(udt):=block([],applyb1(udt,msinrul,mcosrul,mtanrul,masinrul,macosrul,matanrul))");
//                    initialsetup.Append(@",ConvertLog(udt):=block([aliz],matchdeclare(aliz,all),tellsimp(log(aliz),lgog(aliz)/lgog(%e)),tellsimp(%e^aliz,10^(lgog(%e)*aliz)),udt:ratsimp(udt),udt:ev(udt,lgog(10)=1),remrule(log,all),remrule(""^"",all),udt)"); //clear_rules fjernet
                    initialsetup.Append(",ConvertLog(udt):=block([],udt:ratsimp(apply1(udt,merul,mlogrul)),ev(udt,lgog(10)=1))");
                    initialsetup.Append(",ConvertLogN(udt):=block([],udt:ratsimp(apply1(udt,merul,mlogrul)),ev(udt,lgog(10)=1,lgog(%e)=0.4342944819032518))");
                    //                    initialsetup.Append(@",scinot(udt):=block([talfakt],keepfloat:false,if numberp(udt) then (if udt<-10 or (udt>-1 and udt<1) or udt>10 then return(printf(false,concat(""~"",fpprec+4,""e""),udt)) else return(udt)),talfakt:numfactor(udt),if talfakt>=1 and talfakt<10 then return(udt),udt:simplode([printf(false,concat(""~"",fpprec+4,""e""),talfakt),ratsimp(ev(udt/talfakt))],""*""),keepfloat:true,udt)");
// keepfloat fjernet fra scinot som forsøg. Giver problemer med units og scinot
//                    initialsetup.Append(@",scinot(udt):=block([talfakt],if numberp(udt) then (if udt<-10 or (udt>-1 and udt<1) or udt>10 then return(printf(false,concat(""~"",fpprec+4,""e""),udt)) else return(dectal(udt))),talfakt:numfactor(udt),if talfakt>=1 and talfakt<10 then return(udt),udt:simplode([printf(false,concat(""~"",fpprec+4,""e""),talfakt),ratsimp(ev(udt/talfakt))],""*""),udt)");
// ny scinot der kun virker på tal da den nu anvendes i rule
                    initialsetup.Append(@",scinot(udt):=block([talfakt,antbsf],if not(numberp(udt)) then return(udt),if udt=0 then return(0),if abs(udt)<1 then antbsf:NoSigFig+4 else antbsf:NoSigFig+4,if udt<0 then antbsf:antbsf+1,(if udt<-10 or (udt>-1 and udt<1) or udt>10 then return(sremove("" "",printf(false,concat(""~"",antbsf,""e""),udt))) else return(dectal(udt))))");

                    initialsetup.Append(@",scinotall(udt):=block([],apply1(udt,scinotrul))");
//                    initialsetup.Append(@",dectal(udt):=block([talfakt,antbsf,kpfq],kpfq:keepfloat,keepfloat:false,if numberp(udt) then (if abs(udt)<1 then antbsf:fpprec+2 else antbsf:fpprec,return(printf(false,concat(""~"",antbsf,""g""),udt))),talfakt:numfactor(udt),if talfakt<1 then antbsf:fpprec+2 else antbsf:fpprec,udt:simplode([printf(false,concat(""~"",antbsf,""f""),talfakt),ratsimp(ev(udt/talfakt))],""*""),keepfloat:kpfq,udt)");
                    // kun for tal
                    initialsetup.Append(@",dectal(udt):=block([qudt,antbsf],if not(numberp(udt)) then return(udt),if udt=0 or udt=0.0 then return(0),qudt:abs(udt), if qudt<10^-4 then return(scinot(udt)) elseif qudt>10^NoSigFig then return(scinot(udt)),if qudt>1 then antbsf:NoSigFig+1 else antbsf:NoSigFig+2-floor(log(qudt)/log(10)),if udt<0 then antbsf:antbsf+1,udt:printf(false,concat(""~"",antbsf,""f""),udt),udt:striml("" "",udt),udt:strimr(""0"",udt),strimr(""."",udt) )");

                    //scinotsolve erstattet af scinotall
//                    initialsetup.Append(@",scinotsolve(udt):=block([iqw,lignq,outudt],outudt:[],iqw:1,startq,lignq:udt[iqw],if not(freeof(%if,lignq)) then lignq:args(lignq)[2][1],outudt:append(outudt,[eval_string(sconcat(lhs(lignq),""="",scinotall(rhs(lignq))))]),iqw:iqw+1,if iqw<=length(udt) then go(startq),outudt)");
                    //                    initialsetup.Append(",DefinePowerZero(udt):=block([iszeroqwz],matchdeclare (iszeroqwz, lambda ([qwz], qwz =0)),tellsimp(iszeroqwz^iszeroqwz,1),udt:ratsimp(udt),clear_rules(),udt)");
                    //                    initialsetup.Append(",ConvertToDegr(udt):=block([aliz],remfunction(sing),matchdeclare(aliz,all),tellsimp(sin(aliz),sing(180/%pi*aliz)),tellsimp(cos(aliz),cosg(180/%pi*aliz)),tellsimp(tan(aliz),tang(180/%pi*aliz)),if not(HasInvTrig(udt)) then go(slut),tellsimp(asin(aliz),%pi/180*asing(aliz)),tellsimp(acos(aliz),%pi/180*acosg(aliz)),tellsimp(atan(aliz),%pi/180*atang(aliz)),slut,udt:ev(udt,simp),clear_rules(),udt)");
                    //                initialsetup.Append(@",scinot(udt):=block([talfakt],if numberp(udt) then (if udt<1 or udt>=10 then return(printf(false,""~" & MaximaCifre + 4 & @"e"",udt)) else return(udt)),talfakt:numfactor(udt),if talfakt>=1 and talfakt<10 then return(udt),simplode([printf(false,""~" & MaximaCifre + 4 & @"e"",talfakt),ratsimp(ev(udt/talfakt))],""*""))$"");
                    initialsetup.Append(@",funp1(fun,exp):=block([inflag],inflag:true,if mapatom(exp) then false else (if inpart(exp,0) = fun then true else member(true,maplist(lambda([q],funp1(fun,q)),exp))))");
                    initialsetup.Append(@",funp2(fun,exp,var):=block([inflag],inflag:true,if mapatom(exp) then false else (if inpart(exp,0) = fun and member(var,listofvars(exp)) then true else member(true,maplist(lambda([q],funp2(fun,q,var)),exp))))");
                    initialsetup.Append(@",funp(fun,exp,[var]):=if var = [] then funp1(fun,exp) else (if length(var) = 1 then funp2(fun,exp,first(var)) else error(""too many arguments to funp""))"); // afgør om funktion er i udtryk 

                    initialsetup.Append(",GetReal(lign):=block([inlign,kond],if not(freeof(%if,lign)) then (inlign:args(lign)[2][1],kond:args(lign)[1]) else inlign:lign,if not(freeof(%i,carg,parg,kond)) then kond:[],if numberp(ev(float(ev(imagpart(rhs(inlign)),simp)),simp)) then  (if (is(abs(ev(imagpart(ev(float(ev(rhs(inlign),simp)),simp,numer)),simp))<10^(-NoSigFig))) then (if freeof(%i,inlign) then go(slutqw) else (inlign:ev(inlign,numer),go(slutqw))) else (return([]))), if freeof(%i,ev(inlign,%z0:0,%z1:0)) then go(slutqw),inlign:trigsimp(rectform(inlign)),slutqw,if not(freeof(%if,lign)) and kond#[] then inlign:(%if(kond,[inlign],args(lign)[3])),if freeof(%i,inlign) then return(inlign) else return([]) )");
 //                   initialsetup.Append(",GetReal(lign):=block([inlign,kond],if freeof(%i,parg,carg,lign) then return (lign),if not(freeof(%if,lign)) then (inlign:args(lign)[2][1],kond:args(lign)[1]) else inlign:lign,if not(freeof(%i,carg,parg,kond)) then kond:[],if numberp(ev(float(ev(imagpart(rhs(inlign)),simp)),simp)) then  (if (is(abs(ev(imagpart(ev(float(ev(rhs(inlign),simp)),simp,numer)),simp))<10^(-fpprec+1))) then (inlign:ev(inlign,numer),go(slutqw))), if freeof(%i,ev(inlign,%z0:0,%z1:0)) then go(slutqw),inlign:trigsimp(rectform(inlign)),slutqw,if not(freeof(%if,lign)) and kond#[] then inlign:(%if(kond,[inlign],args(lign)[3])),if freeof(%i,inlign) then return(inlign) else return([]) )");
                    initialsetup.Append(@",RemoveImag(lignliste):=block([iqw,outlistq],if lignliste=all then return(all),if lignliste=[] then return([]),lignliste:ev(lignliste,%z0:0,%z1:0),outlistq:[],iqw:1,startq, outlistq:append(outlistq,[GetReal(lignliste[iqw])]),iqw:iqw+1,if iqw<=length(lignliste) then go(startq),outlistq:flatten(outlistq),outlistq)");
//                    initialsetup.Append(@",RemoveImag(lignliste):=block([iqw,outlistq],if lignliste=all then return(all),if lignliste=[] then return([]),if freeof(%i,carg,parg,lignliste) then return(lignliste),lignliste:ev(lignliste,%z0:0,%z1:0),outlistq:[],iqw:1,startq, outlistq:append(outlistq,[GetReal(lignliste[iqw])]),iqw:iqw+1,if iqw<=length(lignliste) then go(startq),outlistq:flatten(outlistq),outlistq)");
                    //                    initialsetup.Append(",IsReal(lign):=block([],if not(freeof(%if,lign)) then lign:args(lign)[2][1],if freeof(%i,lign) then return (true),if numberp(ev(float(ev(imagpart(rhs(lign)),simp)),simp)) then return (is(abs(ev(imagpart(ev(float(ev(rhs(lign),simp)),simp,numer)),simp))<10^(-fpprec+1))), if freeof(%i,ev(lign,%z0:0,%z1:0)) then return(true),if freeof(%i,trigsimp(rectform(rhs(lign)))) then return(true) )");
                    //                    initialsetup.Append(@",RemoveImag(lignliste):=block([iqw,outlistq],if lignliste=all then return(all),if lignliste=[] then return([]),if freeof(%i,lignliste) then return(lignliste),lignliste:ev(lignliste,%z0:0,%z1:0),outlistq:[],iqw:1,startq,if IsReal(lignliste[iqw]) then outlistq:append(outlistq,[lignliste[iqw]]),iqw:iqw+1,if iqw<=length(lignliste) then go(startq),outlistq)");
                    //initialsetup.Append(@",RemoveImagOld(lign):=ev(block([l],if lign=all then return(all),lign:sublist(lign,[z],),l:sublist(lign,lambda([z],numberp(ev(float(ev(imagpart(rhs(z)),simp)),simp)))),l2:sublist(lign,lambda([z],not(numberp(ev(float(ev(imagpart(rhs(z)),simp)),simp))))),l:sublist(l,lambda([z],is(abs(ev(imagpart(ev(float(ev(rhs(z),simp)),simp,numer)),simp))<10^(-fpprec+1)))),l:append(l,l2) ,if not(freeof(%i,l)) then l:trigsimp(rectform(l)) else return(l),if not(freeof(%i,l)) then l:ev(realpart(l),simp) else return(l),l:sublist(l,lambda([z],freeof(%i,z) )),l ),simp)");
                    initialsetup.Append(",ContainsTrig(udt):=block([],not(freeof(sin,cos,tan,acos,asin,atan,udt)))");
                    initialsetup.Append(",IsSolved(lign,varqw):=block([],if listofvars(lhs(lign))=[] then return(false),if funp(%if,lign) then return(true),not(member(varqw,listofvars(rhs(lign)))) and varqw=lhs(lign))");
                    initialsetup.Append(",IsAnySolved(lignliste,varqw):=block([i],i:1,if lignliste=[] then return(lignliste),start,if IsSolved(lignliste[i],varqw) then return(true),i:i+1,if i<=length(lignliste) then go(start),false)");
                    initialsetup.Append(",IsAllSolved(lignliste,varqw):=block([i],i:1,if lignliste=[] then return(lignliste),start,if not(IsSolved(lignliste[i],varqw)) then return(false),i:i+1,if i<=length(lignliste) then go(start),true)");
                    //                    initialsetup.Append(",CheckSolutions(udt,var,soluts):=block([],sublist(soluts,lambda([z],is(ev(abs(ev(subst(rhs(z),var,udt),numer)),numer)<10^(-fpprec+3)) or not(numberp(ev(rhs(z),numer))))))");
                    initialsetup.Append(",CheckSolutions(udtind,var,soluts):=block([udt,qweks,boutq],udt:lhs(udtind)-rhs(udtind),sublist(soluts,lambda([z],if not(freeof(%if,z)) then z:args(z)[2][1],if not(numberp(ev(rhs(z),numer))) then true else (errcatch(qweks:ceiling(log(ev(abs(ev(lhs(udtind),z,numer)),numer))/log(10))),if not(integerp(qweks)) then qweks=1, if fpprintprec>7 then qweks:qweks+fpprintprec-7,if qweks<0 then qweks:0 ,is(ev(abs(ev(udt,z,numer)),numer)<10^(-fpprintprec+qweks+3))) )))");
                    //burde egentlig flyttet til gui nu, men nok sjældent at to_poly_bruges med enheder
                    if (units == 1) // to_poly_solve giver problemer for unit package f.eks. efter CSolve(P=E/t,t) giver J/S Newton istedet for W
                        initialsetup.Append(@",CSolve(lignqw,varqw):=block([solu_1a,solu_2b],solu_1a:[],errcatch(if domain=real then (solu_1a:solvereal(lignqw,varqw)),if IsAnySolved(solu_1a,varqw) then go(slutqw)),errcatch(if polynomialp(ev(lhs(lignqw)-rhs(lignqw),simp),[varqw]) then (if domain=real then solu_1a:flatten(ev(realroots(lignqw),numer)) else solu_1a:flatten(algsys([lignqw],[varqw]))) else go(nxqsol0)),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw) ,nxqsol0,solu_1a:ev(solve(lignqw,varqw),simp),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol3,solu_1a:ev(solve(lignqw,varqw),simp,numer,radexpand),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol4,solu_1a:ev(solve(lignqw,varqw),simp,radcan,numer),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a),slutqw, solu_1a )");
                    else
                        initialsetup.Append(@",CSolve(lignqw,varqw):=solvereal2(lignqw,varqw)");
                    //initialsetup.Append(@",CSolve(lignqw,varqw):=block([solu_1a,solu_2b],solu_1a:[],errcatch(if domain=real then (solu_1a:solvereal(lignqw,varqw)),if IsAllSolved(solu_1a,varqw) then go(slutqw)),errcatch(if polynomialp(ev(lhs(lignqw)-rhs(lignqw),simp),[varqw]) then solu_1a:flatten(algsys([lignqw],[varqw])) else go(nxqsol0)),if domain=real then solu_1a:RemoveImag(solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw) ,nxqsol0 ,if ContainsTrig(lignqw) then go(nxqsol2),solu_1a:ev(errcatch(nicedummies(ev(to_poly_solve(lignqw,varqw),realonly:false)))[1]),if not(freeof(to_poly_solve,solu_1a)) then go(nxqsol1),solu_1a:flatten(errcatch(args(solu_1a))),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol1,solu_1a:errcatch(nicedummies(ev(to_poly_solve(lignqw,varqw),numer,realonly:false,algexact:false)))[1],if not(freeof(to_poly_solve,solu_1a)) then go(nxqsol2),solu_1a:flatten(errcatch(args(solu_1a))),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol2,solu_1a:ev(solve(lignqw,varqw),simp),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol3,solu_1a:ev(solve(lignqw,varqw),simp,numer,radexpand),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol4,solu_1a:ev(solve(lignqw,varqw),simp,radcan,numer),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a),slutqw, if not(freeof(%if,solu_1a)) then solu_1a else RemoveIncorrect(solu_1a) )");
                    //                    initialsetup.Append(@",CSolve(lignqw,varqw):=block([solu_1a,solu_2b],solu_1a:[],errcatch(if polynomialp(ev(lhs(lignqw)-rhs(lignqw),simp),[varqw]) then solu_1a:flatten(algsys([lignqw],[varqw])) else go(nxqsol0)),if domain=real then solu_1a:RemoveImag(solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw) ,nxqsol0 ,if ContainsTrig(lignqw) then go(nxqsol2),solu_1a:ev(errcatch(nicedummies(ev(to_poly_solve(lignqw,varqw),realonly:false)))[1]),if not(freeof(to_poly_solve,solu_1a)) then go(nxqsol1),solu_1a:flatten(errcatch(args(solu_1a))),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol1,solu_1a:errcatch(nicedummies(ev(to_poly_solve(lignqw,varqw),numer,realonly:false,algexact:false)))[1],if not(freeof(to_poly_solve,solu_1a)) then go(nxqsol2),solu_1a:flatten(errcatch(args(solu_1a))),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol2,solu_1a:ev(solve(lignqw,varqw),simp),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol3,solu_1a:ev(solve(lignqw,varqw),simp,numer,radexpand),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol4,solu_1a:ev(solve(lignqw,varqw),simp,radcan,numer),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a),slutqw, solu_1a )");

                    // generate error if complex. Bør nok have numer, men pt ikke nødvendig da kun bruges til draw2d
                    initialsetup.Append(",RealOnly(zqw):=if freeof(%i,zqw) then zqw else error()");
                    initialsetup.Append(",CheckDef(zqw,qvar):=block(if not(freeof(%i,zqw)) then error() elseif maybe(equal(parse_string(qvar),eval_string(qvar)))=false then error() else zqw)");

// kun numerisk nyere laver eksakt                    initialsetup.Append(",NIntegrate(udt,var,x0,x1):=block([res1,res2],if x0=minf or x1=inf then res1:quad_qagi(udt,var,x0,x1) else res1:quad_qags(udt,var,x0,x1),if res1[4]=0 then return (res1[1]) else res1:quad_qag(udt,var,x0,x1,1),if res1[4]=0 then return (res1[1]),'integrate(udt,var,x0,x1))");
//                    initialsetup.Append(",NIntegrate(udt,var,x0,x1):=block([res1,res2,numer],if x0=minf or x1=inf then res1:quad_qagi(udt,var,x0,x1) else res1:quad_qags(udt,var,x0,x1),if res1[4]=0 then return (res1[1]) else res1:quad_qag(udt,var,x0,x1,1),if res1[4]=0 then return (res1[1]) else res1:IntegrateDef(udt,var,x0,x1),if res1#[] then return (res1),integrate(udt,var,x0,x1))"); // flyttet til solvereal.mac
//                    initialsetup.Append(",Integrate(udt,var,x0,x1)::=block([res1,res2,numer],numer:false,errcatch(res1:ev(integrate(udt,var,x0,x1),numer=false)),if not(numberp(x0) and numberp(x1)) and x0#minf and x1#inf  then return(res1),if numberp(ev(res1,numer)) then return(res1),res1:ev(res1,numer),if numberp(res1) then return(res1),errcatch(res2:ldefint(udt,var,x0,x1)),if numberp(res2) then return(res2),res2:NIntegrate(udt,var,x0,x1),if numberp(res2) then return(res2),res1)");  // flyttet til solvereal.mac
                    // Integrate ændret da resultat fra første integrale ikke blev udskrivet så altid fortsatte til numer. Problem?
//                    initialsetup.Append(",Integrate(udt,var,x0,x1):=block([res1,res2],numer:false,errcatch(res1:ev(integrate(udt,var,x0,x1))),if not(numberp(x0) and numberp(x1)) and x0#minf and x1#inf  then return(res1),res1:ev(res1,numer),if numberp(res1) then return(res1),errcatch(res2:ldefint(udt,var,x0,x1)),if numberp(res2) then return(res2),res2:NIntegrate(udt,var,x0,x1),if numberp(res2) then return(res2),res1)");

// flyttet til solvereal-pakke     initialsetup.Append(@",SolveIneq(lignq,varq):=block([soluqw],soluqw:[],load(solve_rat_ineq),errcatch(soluqw:solve_rat_ineq(lignq),soluqw),if soluqw#[] then return(soluqw),load(fourier_elim),soluqw:fourier_elim([lignq],[varq]),if soluqw#[] then return([soluqw]),print(""kunne ikke løses""))");

                    //                initialsetup.Append(","); CSolve(x*%e^(2*x)+%e^(2*x)=0,x); CSolve(y2/y1=(x2/x1)^a,a); CSolve(0.1*x^4+0.8*x^3+0.6*x^2-2*x-1=0,x);
                    return initialsetup.ToString();
                }
            }

            // settings der kan ændres via gui og skal derfor køres hver gang der udføres ny kommando da de kan være ændret
            private string GuiSettingsString // bruges ikke mere
            {
                get
                {
                    StringBuilder initialsetup = new StringBuilder();
                    //if (complex == 1)
                    //{
                    //    initialsetup.Append("domain:complex,");
                    //    initialsetup.Append("realonly:false,");
                    //}
                    //else
                    //{
                    //    initialsetup.Append("domain:real,");
                    //    //                        initialsetup.Append("realonly:true,");  // giver problemer for to_poly_solve
                    //}

//                    if (exact == 2) // num
//                        initialsetup.Append(@"dectalall(udt):=dectalallNum(udt),");
//                    else // exact og auto så afrundes kun float dvs kommatal eller Exp not.
//                        initialsetup.Append(@"dectalall(udt):=block([qout],qout:errcatch(apply1(udt,floatrul)),if qout=[] then udt else qout[1]),");

                    //dectalall(carg(1+%i));
                    // håndteres nu via $%,numer;     dette giver problemer med nogle to_poly_solve
                    //if (exact == 2)
                    //{
                    //    initialsetup.Append("numer:true,");
                    //}
                    //else
                    //{
                    //    initialsetup.Append("numer:false,");
                    //}
                    //if (exact == 1) //exact  duer heller ikke da der kommer afrundede decimaltal der bliver omdannet brøker
                    //{
                    //    initialsetup.Append("keepfloat:false,");
                    //}
                    //else //num eller auto
                    //{
                    //    initialsetup.Append("keepfloat:true,");
                    //}


                    //                initialsetup.Append("");

//                    initialsetup.Append("NoSigFig:" + Convert.ToString(AntalCifre)); // +"," fjernet da nu sidste
// Disse to forsøges låst fast på høje værdier i solvereal.mac
//                    initialsetup.Append("rootsepsilon:1E-" + Convert.ToString(AntalCifre + 2) + ",");
//                    initialsetup.Append("algepsilon:1E+" + Convert.ToString(AntalCifre + 2) + ",");

//                    initialsetup.Append("ratepsilon:2E-" + Convert.ToString(AntalCifre + 1) + ","); // styrer hvor præcist der omregnes fra dec til brøk 2E-8 er default. Skal være mindre end rootsepsilon. Integrate(sin(x^(2+x) )/x^4 ,x,1,3),numer; stopper ikke hvis der pilles ved den
//                    initialsetup.Append("fpprec:" + Convert.ToString(AntalCifre + 1) + ","); // var  +1 før dec tal. påvirker kun bigfloat
//                    initialsetup.Append("fpprintprec:0"); // + Convert.ToString(AntalCifre + 2)); // var +0 før dectalall  // gav problemer når integers og små decimaltal skulle trækkes fra hianden: 1-1.5•(10)^(-8) giver 0 når fpprintrec er sat til 9
//                    initialsetup.Append("fpprintprec:" + antalcifre); //antal cifre der skal vises fpprec styrer hvor mange der skal bruges i beregninger
                    return initialsetup.ToString();
//solve(Integrate(sqrt(3*x+9)-x-3,x,0,k),k); 
//Integrate(sin(%pi/180*x^(2+x))/x^4,x,1,3); 	

                }
            }
            
            public void SetSettings()
            {
                StringBuilder settext = new StringBuilder("[");
                settext.Append(SettingsString);
                settext.Append("]$");

                ExecuteMaximaCommand(settext.ToString(),0); 
            }

            // nulstiller variable mm i kørende process
            public void Reset(string extrakommando)
            {
//                ExecuteMaximaCommand("[kill(allbut(ConvertToDegr,ConvertLog,DefinePowerZero,IsSolved,RemoveImag,IsAnySolved,CheckSolutions,CSolve,NIntegrate,Integrate)),reset()]$[" + SettingsString + "]$",0);
//                ExecuteMaximaCommand("reset(linenum,%,features,%rnum)$[" + SettingsString + "]$", 0);
                if (extrakommando == null)
                    extrakommando = string.Empty;
                if (extrakommando.Length>1)
                {
                    extrakommando = extrakommando + ",";
                }
//                ExecuteMaximaCommand("[" + extrakommando + "reset(linenum,%,features,%rnum)," + GuiSettingsString + "]$", 0);
                ExecuteMaximaCommand("[" + extrakommando + "reset(linenum,%,features,%rnum)]$", 0);
                KommandoNr = 2;
                mindex = 2;
                LastWasOutput = false;
            }

            private void WaitForOutput()
            {
                string gemlout;
                // Waits for the maxima process to finish
                try
                {
                    // first wait for the output to begin
                    int n = 0;
                    do
                    {
                        Thread.Sleep(20);
                        n++;
                    }
                    while (outputstarted == false && n < 100);

                    // hvad skal det? fjernet som forsøg på at fjerne opstartsfejl
                    //if (n == 100)
                    //{
                    //    return;
                    //}

                    // first wait for the output (%o to begin
                    n = 0;
                    do
                    {
                        Thread.Sleep(20);
                        n++;
                    }
                    while (outputstarted2 == false && n < 1600);

                    // så ventes på at output printes færdigt. 
                    // metoden vil dog aldrig være 100% sikker.

                    n = 0;
                    do
                    {
                        Thread.Sleep(20);
                        n++;
                    }
                    while (finish == false && n < 1600);
//                    finish = true; // Hvis ikke processen selv har kunnet afgøre at den er færdig antages nu at det må den være
//                    outputstarted = true;
//                    outputstarted2 = true;

                    // sæt fejl og spørgsmålsflag
                    if (n == 1600)
                    {
                        gemlout = LastMaximaOutput.ToString();
                        ConsoleInterrupt();
                        //CloseProcess();
                        //StartMaximaProcess();
                        LastMaximaOutputBuilder.Append(gemlout);
                        errCode = 5;
                    }
                    if (LastMaximaOutput.Contains("syntax error"))
                        errCode = 2;
                    else if (LastMaximaOutput.Contains("incorrect syntax:"))
                        errCode = 2;
                    else if (LastMaximaOutput.Contains("?") && !LastMaximaOutput.Contains("?explicit"))
                        question = 1;
                }
                catch { }
            }


        // wait=1 betyder vent på at kommandoen bliver udført af maxima 
        public void ExecuteMaximaCommand(string inputText,int wait)
            {
//                inputText = inputText + "\r\n";
                errCode = 0;
                try
                {
                    if (maximaProcess == null)
                    {
                        StartMaximaProcess();
                    }
                    else if (maximaProcess.HasExited)
                    {
                        StartMaximaProcess();
                    }

                }
                catch
                {
                    StartMaximaProcess();
                }
                if (errCode > 0)
                {
                    CloseProcess();
                    return;
                }
                try
                {
                    if (finish == false)
                        WaitForOutput();

                    // Write each line to the redirected input stream
                    //                maximaStreamWriter.WriteLine("5+6;\r\n");
                    finish = false;
                    errCode = 0;
                    question = 0;
                    outputstarted = false;
                    outputstarted2 = false;
                    LastMaximaOutputBuilder.Clear();
                    int numInputLines = 0;
                    if (!String.IsNullOrEmpty(inputText))
                    {
                        maximaStreamWriter.WriteLine("[]$" + inputText + "slutprut;"); // forsøg med ikke at printe slutprut da det giver problemer for read() i solvesystem
//                        maximaStreamWriter.WriteLine("[" + GuiSettingsString + "]$" + inputText + "print(\"slutprut\");");
                        //                        maximaStreamWriter.WriteLine("[" + GuiSettingsString + "]$" + inputText + "print(\"slutprut\",linenum:linenum-1,%)$");
                        numInputLines++;
                        KommandoNr = KommandoNr + CountOccInStr(inputText, ';') + CountOccInStr(inputText, '$');
                        if (wait == 1)
                            WaitForOutput();
                    }
                }
                catch 
                {
                    CloseProcess();
                }
//                maximaProcess.Close();
            }

        
        public void AnswerQuestion(string inputText)
            {
                //                inputText = inputText + "\r\n";
                if (maximaProcess == null)
                {
                    return;
                }
                else if (maximaProcess.HasExited)
                {
                    return;
                }

                if (errCode > 0)
                {
                    CloseProcess();
                    return;
                }

                if (finish == false)
                    WaitForOutput();

                finish = false;
                errCode = 0;
                question = 0;
                outputstarted = false;
                outputstarted2 = false;
                LastMaximaOutputBuilder.Clear();
                if (!String.IsNullOrEmpty(inputText))
                {
//                    maximaStreamWriter.WriteLine(inputText + ";" + Environment.NewLine);
                    maximaStreamWriter.WriteLine(inputText + ";slutprut;");
                    //CSolve(matrix([x+y],[x-y])=matrix([a],[2]),x)

                    //                    WaitForOutput(); // fjernet i version 1.05 for at få wait-form op, også når der er spørgsmål. ellers kunne der være beregninger som ikke kunne stoppes
                }

            }



        
        public void CloseProcess()
            {
                // End the input stream
                try
                {
                    //ConsoleInterrupt();
                    //maximaProcess.Close();
                    maximaProcess.Kill();
                    Process[] localByName = Process.GetProcessesByName("maxima");
                    
                    foreach (Process p in localByName)
                    {
                        p.Kill();
                    }
                    Process[] localByName2 = Process.GetProcessesByName("sbcl");

                    foreach (Process p in localByName2)
                    {
                        p.Kill();
                    }
        }
        catch { }
                try
                {
                    maximaStreamWriter.Close();
                }
                catch { }
            }

        
        public void ConsoleInterrupt()
        {
            maximaStreamWriter.WriteLine("\x3"); // send ctrl+c
        }

            private int CountOccInStr(string source, char occ)
            {
                int count = 0;
                foreach (char c in source)
                    if (c == occ) count++;
                return count;
            }

            private void MaximaOutputHandler(object sendingProcess,
                DataReceivedEventArgs outLine)
            {
                // Collect the output.
                string outp;
                outp = outLine.Data;

                if (!String.IsNullOrEmpty(outp))
                {
                    numOutputLines++;
                    outputstarted = true;

                    // Add the text to the collected output.

                    // hvis kun startup text så afslut hurtigt // fjernet for at se om det bliver mere stabilt hvis kun ses efter The function bug_
                    //if (numOutputLines > 7)
                    //{
                    //}
                    //else if (numOutputLines == 5)
                    //{
                    //    question = 0;
                    //    outputstarted2 = true;
                    //    finish = true;
                    //    return;
                    //}
                    if (outp.Contains("The function bug_") || outp.Contains("Maxima restarted."))
                    {
                        question = 0;
                        outputstarted2 = true;
                        finish = true;
                        return;
                    }
                    else if (outp.Contains("Console interrupt."))
                    {
                        question = 0;
                        outputstarted2 = true;
                        finish = true;
                        return;
                    }

                    // hvis output er afsluttet korrekt afsluttes hurtigt
                    //string it;
                    //it = "(%i" + (KommandoNr + 1) + ")";
                    //if (outp.Contains(it))
                    //{
                    //    finish = true;
                    //    outputstarted2 = true;
                    //    return;
                    //}
                    if (outp.Contains("slutprut") )
                    {
                        finish = true;
                        outputstarted2 = true;
                        return;
                    }
                    else if (outp.Contains("Acceptable answers are"))
                    {
                        finish = true;
                        outputstarted2 = true;
                        question = 1;
                        return; 
                    }
                    else if (outp.Contains("nable to solve") )
                    {
                        question = 0;
                        return;
                    }
                    else if (outp.Contains("Nonalgebraic argument given to 'topoly'") )
                    {
                        question = 0;
                        return;
                    }
                    else if (outp.Contains("log: log(0) has been generated.") )
                    {
                        question = 0;
                        return;
                    }

                    LastMaximaOutputBuilder.Append(Environment.NewLine + outp);
                    MaximaOutputBuilder.Append(Environment.NewLine + outp);
                    
                    if (outp.Contains("incorrect syntax:"))
                    {
                        outputstarted2 = true;
                        errCode = 3;
                        finish = true;                        
                    }
                    else if (outp.Contains("positive, negative or zero?") || outp.Contains("positive or negative?")) //rettet i version 5.3.31-1 fra "positive, negative, or zero?" altså fjernet et komma
                    {
                        finish = true;
                        outputstarted2 = true;
                        question = 1;
                    }
                    else if (outp.Contains("variables to solve for?"))
                    {
                        finish = true;
                        outputstarted2 = true;
                        question = 1;
                    }

                    // add input/output to array
                    if (outp.Contains("(%o"))
                    {
                        mindex = FindNo(outp);
                        if (mindex > 0)
                        {
                            maximaoutput[mindex] = outp.Substring(outp.IndexOf(')', 3) + 1);
                            LastWasOutput = true;
                        }
                    }
                    else if (outp.Contains("(%i"))
                    {
                        mindex = FindNo(outp);
                        if (mindex > 0)
                        {
                            maximainput[mindex] = outp.Substring(outp.IndexOf(')', 3) + 1);
                            LastWasOutput = false;
                        }
                    }
                    else if (LastWasOutput)
                    {
                        maximaoutput[mindex] = maximaoutput[mindex] + outp;
                    }
                    else if (!LastWasOutput)
                    {
                        maximainput[mindex] = maximainput[mindex] + outp;
                    }

                    // sætter flag der indikerer at sidste (%oxx er nået
                    if (outp.Contains("(%o" + KommandoNr + ")"))
                    {
                        outputstarted2 = true;
                    }

                    // set flag der indikerer at output er slut
                    //if (outp.Contains("slutprut"))
                    //{
                    //    finish = true;
                    //}
                    //if (outp.Contains(" ^"))
                    //{
                    //    Merror = true;
                    //    finish = true;
                    //}

                }

            }

            // finds the input/outputnumber ex: (%i23) or (%o24)
            private int FindNo(string text)
        {
                int spos;
                int n;
                int ascval=0;
                string nom="";
                spos = text.IndexOf("(%i");
                if (spos<0)
                    spos = text.IndexOf("(%o");
                n = 0;
                do
            	{
	                nom=nom + text[spos+3+n].ToString();
                        n++;
            	} while (text[spos+3+n]!=')' & n<3);
                try{
                ascval = Convert.ToInt16(nom);
                }
                catch
                {
                    ascval = 0;
                }
                return ascval;
        }


            public void TurnUnitsOn(string prevcomm, string postcomm)
            {
                string[] outunits;
                try
                {
                    StringBuilder initialsetup = new StringBuilder();
                    if (units == 1)
                    {
                        if (prevcomm.Length > 0)
                            initialsetup.Append(prevcomm + "$");
                        initialsetup.Append("reset()$[keepfloat:false,usersetunits:[N,J,W,Pa,C,V,F,Ohm,T,H,K]]$"); // reset sat ind for løse problem med f.eks. 2*J/t efter to_poly_solve(-0.1*x^(-0.5)-0.2*x+1=0,x); men giver måske andre problemer?
                        initialsetup.Append("load(unit)$");
                        initialsetup.Append("[" + SettingsString); // indført efter reset() sat ind
                        // to_poly_solve giver problemer for unit package f.eks. efter CSolve(P=E/t,t) giver J/S Newton istedet for W
                        initialsetup.Append(@",CSolve(lignqw,varqw):=block([solu_1a,solu_2b],solu_1a:[],errcatch(if domain=real then (solu_1a:solvereal(lignqw,varqw)),if IsAnySolved(solu_1a,varqw) then go(slutqw)),errcatch(if polynomialp(ev(lhs(lignqw)-rhs(lignqw),simp),[varqw]) then (if domain=real then solu_1a:flatten(ev(realroots(lignqw),numer)) else solu_1a:flatten(algsys([lignqw],[varqw]))) else go(nxqsol0)),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw) ,nxqsol0,solu_1a:ev(solve(lignqw,varqw),simp),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol3,solu_1a:ev(solve(lignqw,varqw),simp,numer,radexpand),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a), if IsAnySolved(solu_1a,varqw) then go(slutqw),nxqsol4,solu_1a:ev(solve(lignqw,varqw),simp,radcan,numer),if domain=real then solu_1a:RemoveImag(solu_1a),solu_1a: CheckSolutions(lignqw,varqw,solu_1a),slutqw, solu_1a )");
                        
//                        initialsetup.Append("," + FuncDefinitions);  // fjernet da funcdefintions nu er i solvereal.mac. Forhåbentlig er den ikke nødvendig
                        initialsetup.Append("]$");
                        if (moutunits.Length > 0)
                        {
                            outunits=moutunits.Split(",".ToArray());
                            foreach (string unit in outunits)
                            {
                                initialsetup.Append("setunits(");
                                initialsetup.Append(unit);
                                initialsetup.Append(")$");                                
                            }
                        }
                        if (postcomm != "")
                            initialsetup.Append(postcomm + "$");
                        ExecuteMaximaCommand(initialsetup.ToString(), 0);
                    }
                }
                catch { }
            }
            public void UpdateUnits()
            {
                string[] outunits;
                try
                {
                    StringBuilder initialsetup = new StringBuilder();
                    if (units == 1)
                    {
//                        initialsetup.Append("[");
//                        initialsetup.Append("setunits([N,J,W,Pa,C,V,F,Ohm,T,H,K])");
                        initialsetup.Append("uforget(append(globalbaseunitlisting,globalderivedunitlisting))$");
//                        initialsetup.Append("]$");
                        if (moutunits.Length > 0)
                        {
                            outunits = moutunits.Split(",".ToArray());
                            foreach (string unit in outunits)
                            {
                                initialsetup.Append("setunits(");
                                initialsetup.Append(unit);
                                initialsetup.Append(")$");
                            }
                        }
                        ExecuteMaximaCommand(initialsetup.ToString(), 0);
                    }
                }
                catch { }
            }

            public string MaximaOutput
            {
                get
                {
                    if (MaximaOutputBuilder == null)
                    {
                        return "";
                    }
                    else
                    {
                        return MaximaOutputBuilder.ToString(0, MaximaOutputBuilder.Length);
                    }
                }
            }

            public int Question
            {
                get
                {
                    return question;
                }
                set
                {
                    question = value;
                }
            }
            
            public string QuestionText
            {
                get 
                {
                    return LastMaximaOutput;
                }
            }

            
            public string LastMaximaOutput
            {
                get 
                {
                    if (LastMaximaOutputBuilder == null)
                    {
                        return ""; 
                    }
                    else
                    {
                        return LastMaximaOutputBuilder.ToString(0, LastMaximaOutputBuilder.Length);
                    }                
                }

             }

            public int Exact
            {
                get
                {
                    return exact;
                }
                set
                {
                    exact = value;
                }
            }
            public int AntalCifre
            {
                get 
                {
                    return antalcifre;
                }
                set 
                {
                    antalcifre = value;
                }
            }
            public int Complex
            {
                get
                {
                    return complex;
                }
                set
                {
                    complex = value;
                }
            }
            public int Units
            {
                get
                {
                    return units;
                }
                set
                {
                    units = value;
                }
            }
            public int ErrCode
            {
                get
                {
                    return errCode;
                }
            }

            public bool IsMaximaStarted
            {
                get
                {
                    return isMaximaStarted;
                }
            }

    public string MaximaInputArray(int nr)
            {
                if (nr < mindex)
                    return maximainput[nr];
                else
                    return null;
            }

            
            public string MaximaOutputArray(int nr)
            {
                if (nr < mindex)
                    return maximaoutput[nr];
                else
                    return null;
            }
            public int Finished
            {
                get
                {
                    if (finish)
                        return 1;
                    else
                        return 0;
                }
                set
                {
                    if (value == 1)
                        finish = true;
                    else
                        finish = false;
                }
            }

            public string OutUnits
            {
                get 
                {
                    return moutunits;
                }
                set 
                {
                    moutunits = value;
                }
            }

        public string modulnavn
        {
            get
            {
                string navne=maximaProcess.ProcessName;
                for (int i = 0; i < maximaProcess.Modules.Count-1; i++)
			{
                navne = navne + maximaProcess.Modules[i].FileName;
			}
                navne = "";
//                Process[] localByName = Process.GetProcesses();
                Process[] localByName = Process.GetProcessesByName("maxima");
                foreach (Process p in localByName)
                {
                    if (p.ProcessName.Contains("maxima"))
                        navne = navne + p.ProcessName;
                }


                return navne;
            }
        }

        public string CheckForUpdate()
        {
//            System.Data.ConnectionState Description = 0;
//            bool con = InternetGetConnectedState(ref Description, 0);

            if (ConnectionExists())
            {
                MathMenu.com.eduap.WordMat.WebServiceWordMat webs = new MathMenu.com.eduap.WordMat.WebServiceWordMat();
                try
                {
                    return webs.Version();
                }
                catch
                {
                    return "";
                }
            }
            else
                return "";
        }

        public string GetVersionNews()
        {
            //            System.Data.ConnectionState Description = 0;
            //            bool con = InternetGetConnectedState(ref Description, 0);

            if (ConnectionExists())
            {
                MathMenu.com.eduap.WordMat.WebServiceWordMat webs = new MathMenu.com.eduap.WordMat.WebServiceWordMat();
                try
                {
                    return webs.News();
                }
                catch
                {
                    return "";
                }
            }
            else
                return "";
        }


        bool ConnectionExists()
        {
            try
            {
                System.Net.Sockets.TcpClient clnt = new System.Net.Sockets.TcpClient("www.google.com", 80);
                clnt.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        }


    



//}
