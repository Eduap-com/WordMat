using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.ComponentModel;
using System.Collections;
// Bemærk at visual studio skal køres som administrator ellers virker det ikke

//namespace MathMenuLib  // giver vidst nok problemer med com
//{

    //interface som viser det som ses fra com i vba
    public interface IMaximaConn
    {
        void StartMaximaProcess();
        void SetSettings();
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
//                    ExecuteMaximaCommand("solve(sin(x)=1/2,x);", 1); // for at hente vigtige funktioner ind, men har ikke rigtig nogen effekt
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
            catch (System.Exception ex)
            {
                return false;
            }
        }

        }


    



//}
