//
//  MaximaFunctions.c
//  MaximaConnection
//
//  Created by NyborgGym on 19/09/13.
//  Copyright (c) 2013 EDUAP. All rights reserved.
//

// Der skal være en return i hver funktion. Typisk i slutningen. Ellers går Word ned.
// Det gælder også void-funktioner. De skal afsluttes med return ;

// stdin er keyboard fra den kan man læse
// sdtout er skærm. Dertil kan man skrive

#include <stdio.h>
#include <stdlib.h>
#include <unistd.h>
#include <string.h>
#include <errno.h>
#include <signal.h>
#include <ctype.h>
#include <stdint.h>
#include <time.h>
#include <math.h>
#include <sys/wait.h>
#include <sys/types.h>
#include <sys/fcntl.h>
#include <sys/ioctl.h>
#include "HelpFunctions.h"
#include "MaximaFunctions.h"

#define EXIT_FAILURE 1
#define PIPE_READ 0
#define PIPE_WRITE 1

pid_t ChildPID;
int fdWrite[2]; //2 FildeDescriptors for the pipe between the 2 processes. read from fd[0] write to fd[1]
int fdRead[2];
char *MaximaPath;  //[]= "/Applications/Maxima.app/Contents/Resources/maxima.sh";
long mindex; // index of last maximacommand
long CommandNo; //
long MaximaStarted; // 0 if maxima is not running. Set to 1 when maxima is started

//char* MaximaOutString;

int setNonblocking(int fd)
{
    int flags;
    
    /* If they have O_NONBLOCK, use the Posix way to do it */
//#if defined(O_NONBLOCK)
    /* Fixme: O_NONBLOCK is defined but broken on SunOS 4.1.x and AIX 3.2.5. */
    if (-1 == (flags = fcntl(fd, F_GETFL, 0)))
        flags = 0;
    return fcntl(fd, F_SETFL, flags | O_NONBLOCK);
//#else
    /* Otherwise, use the old way of doing it */
//    flags = 1;
//    return ioctl(fd, FIOBIO, &flags);
//#endif
}


// Forsøg på at løse problem at den går ned
// Indtil videre uden held dog. Hvis Der trykkes luk på debuggen og fx SetMaximaPath køres, så crasher den
static void catch_signal(int signal) {
    sleep(15);
//    CloseMaxima();
    _exit(0);
}

long SetMaximaPath(const char *SetPath)
{
    MaximaPath=malloc(strlen(SetPath)+1);
    strcpy(MaximaPath,SetPath);
    return 0;
}



// Starts Maxima as a hidden process and immidately returns control to the calling process
// return value 1 - maxima started ok
// -1 error maxima-file not found
// -2 signal error
long StartMaxima()
{
//    if (IsMaximaRunning()>0) // already started . Moved to VBA
//    {
//        return 0;
//    }
//    signal(SIGPIPE, SIG_IGN); // Sørgrer for at fejl der skyldes at der skrives til en lukket pipe ignoreres. Men crasher 64-bit
    //    signal(SIGINT, catch_signal); // terminal interrupt signal
//    if (signal(SIGFPE, catch_signal)==SIG_ERR){
//        return -2;
//    }; // Floating Point error
    //    signal(SIGFPE, SIG_IGN); // Floating Point error
    //    signal(SIGBUS, catch_signal); // bad access
    //    signal(SIGSEGV, catch_signal); // invalid memory reference
    //    signal(SIGSYS, catch_signal); // bad system call
    
    pipe(fdWrite);
    pipe(fdRead);
    ChildPID = vfork();
    
    if (ChildPID < 0){
        //fork error
        perror("fork");
        exit(EXIT_FAILURE);
    }
    else if (ChildPID == 0){
        // Child process
        /* Close up standard input of the child */
        /* Duplicate the input side of pipe to stdin */
        close(fdWrite[PIPE_WRITE]);
        dup2(fdWrite[PIPE_READ],0); //dup2(oldfd,newfd)   0=stdin 1=stdout
        close(fdRead[PIPE_READ]);
        dup2(fdRead[PIPE_WRITE],1);
        
        if (MaximaPath==NULL)
            SetMaximaPath("/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/Maxima.app/Contents/Resources/maxima.sh");
        
        execl (MaximaPath, MaximaPath ,(char *) NULL); // exec'd applications inherit stdiinput and stdoutput from calling process // if succesful the childprocess is left an dthe following lines never run
        
        _exit(0); // Only runs if execl fails. when using vfork you may not change variables and may only use _exec(). Never use exit(0) with fork or vfork
        
    }
    else if (ChildPID > 0)
    {
        // Parent process
        close(fdWrite[PIPE_READ]);
        close(fdRead[PIPE_WRITE]);
        dup2(fdRead[PIPE_READ],0);
        
//        setNonblocking(0);  ' 64-bit problem the next 4 lines replaces
        int flags;
        if (-1 == (flags = fcntl(0, F_GETFL, 0)))
            flags = 0;
        fcntl(0, F_SETFL, flags | O_NONBLOCK);
        

        MaximaStarted=1; //IsMaximaRunning();
        mindex=1;
        return MaximaStarted;
    }
    // Both process
    return 1; // never reached
}


long StartMaxima2()
{
    if (IsMaximaRunning()>0) // already started
    {
        return 0;
    }
    signal(SIGPIPE, SIG_IGN); // Sørgrer for at fejl der skyldes at der skrives til en lukket pipe ignoreres. Men crasher 64-bit
    //    signal(SIGINT, catch_signal); // terminal interrupt signal
    if (signal(SIGFPE, catch_signal)==SIG_ERR){
        return -2;
    }; // Floating Point error
    //    signal(SIGFPE, SIG_IGN); // Floating Point error
    //    signal(SIGBUS, catch_signal); // bad access
    //    signal(SIGSEGV, catch_signal); // invalid memory reference
    //    signal(SIGSYS, catch_signal); // bad system call
    
    pipe(fdWrite);
    pipe(fdRead);
    //    MaximaOutString=malloc(1000);
    
    //    setvbuf(stdout, NULL, _IONBF, BUFSIZ);
    //    setvbuf(stdin, NULL, _IONBF, BUFSIZ);
    ChildPID = vfork();
    
    if (ChildPID < 0){
        //fork error
        perror("fork");
        exit(EXIT_FAILURE);
    }
    else if (ChildPID == 0){
        // Child process
        /* Close up standard input of the child */
        /* Duplicate the input side of pipe to stdin */
        close(fdWrite[PIPE_WRITE]);
        dup2(fdWrite[PIPE_READ],0); //dup2(oldfd,newfd)   0=stdin 1=stdout
        close(fdRead[PIPE_READ]);
        dup2(fdRead[PIPE_WRITE],1);
        
        //        write(fdRead[PIPE_WRITE], "asdfg", 5);
        //            char Path[]= "/Applications/Maxima.app/Contents/MacOS/Maxima";  // starter først shell hvorfra maxima køres derved mistes kontakten til Maxima brug istedet stien: /Applications/Maxima.app/Contents/Resources/maxima.sh
        if (MaximaPath==NULL)
            SetMaximaPath("/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/Maxima.app/Contents/Resources/maxima.sh");
        
        execl (MaximaPath, MaximaPath ,(char *) NULL); // exec'd applications inherit stdiinput and stdoutput from calling process // if succesful the childprocess is left an dthe following lines never run
        //        sleep(5);
        //        MaximaStarted=0;
        //        close(fdWrite[PIPE_READ]);
        //        return -1;
        
        _exit(0); // Only runs if execl fails. when using vfork you may not change variables and may only use _exec(). Never use exit(0) with fork or vfork
        
    }
    else if (ChildPID > 0)
    {
        // Parent process
        
        //        printf("time: %i \n",sleep(4));
        //        printf("time: %i \n",sleep(4));
        close(fdWrite[PIPE_READ]);
        close(fdRead[PIPE_WRITE]);
        dup2(fdRead[PIPE_READ],0);
        //        char s[100];
        //        char command[]="2+4;\n 4+5;\n ";
        //    write(fdWrite[PIPE_WRITE], command, strlen(command));
        
        //        fprintf(stdout,command,strlen(command));
        
        //        fflush(stdout);
        //        wait(1);
        
        //        close(fdWrite[PIPE_WRITE]); // sends EOF
        
        setNonblocking(0);
        MaximaStarted=1; //IsMaximaRunning();
        mindex=1;
        return MaximaStarted;
        //        exit(0);
    }
    // Both process
    return 1; // never reached
}

long IsMaximaRunning()
{
    pid_t pid;
    int status;
    
    pid = waitpid (ChildPID,&status, WNOHANG|WUNTRACED); // WNOHANG means nonblocking, WUNTRACED means also report on stopped & terminated
    if (pid == -1) {            /* error calling waitpid, child not started or ...       */
        return -1;
    }
    else if (pid == 0) {        /* no process has reported. child still running or never run. Dont change maximastarted then  */
    }
    else if (pid == ChildPID) {  /* child ended                 */
        if (WIFEXITED(status))
            //            printf("Child ended normally.n");
            MaximaStarted=-2;
        else if (WIFSIGNALED(status))
            //            printf("Child ended because of an uncaught signal.n");
            MaximaStarted=-3;
        else if (WIFSTOPPED(status))
            //            printf("Child process has stopped.n");
            MaximaStarted=-4;
        else
            MaximaStarted=-5;
    }
    return MaximaStarted;
}

long GetCommandIndex()
{
    return mindex;
}

// Brugt til at teste om det er muligt at fange fejl
long ForceDivZero()
{
    long a=0;
    long b=1;
    long c;
    signal(SIGFPE,SIG_IGN);
    c=b/a;
    return 2;
}

void SetCommandIndex(const long ind)
{
    mindex=ind;
}

// Sender command til maxima. Fejler hvis den køres fra VBA. Brug SendToMaximaVB
// Der skal tilføjes \n før end kommandoen sendes
// Fra VB kan det gøres med vblf
long SendToMaxima( char* command)
{
    
//    long nocmd=0; //No of commands
    
//    if (IsMaximaRunning()<1)
//        StartMaximaWait(15);
//    if (IsMaximaRunning()<1)
//        return -1;
//    command="2+3;\n";
    
//    nocmd=CountInString(0, command,";") + CountInString(0, command,"$");
//    mindex = mindex + nocmd;
    
    //    const int ReadSize = 1000;
    //    char *buffer=malloc(ReadSize);
    //    char *s=malloc(1000); //[1000];
    // Hvis ikke kommandoen har \n i slutningen skal det tilføjes
    unsigned long l = strlen(command); // skal være strlen hvis køres fra c ellers VBstrlen

    //    if (command[l-1]!='\n')
//    {
//        char *commandn=malloc(l+1);
//        strcpy(commandn,command);
//        commandn[l]='\n';
//        l++;
//        write(fdWrite[PIPE_WRITE], commandn, l);
//        free(commandn);
//    }
//    else
//    {
        write(fdWrite[PIPE_WRITE], command, l); // der skal måske checkes for fejl ved write til pipe
//    }
    
    //    sleep(1); // uden denne kan der komme broken pipe fejl, men det er ikke sikkert. Forsøgt fjernet igen.
    return 1;
    
}

// Bestemmer længden af en streng fra VBA.Testet. Den virker fra Word
long VBstrlen(const char *vbstr) // size_t
{
    if (vbstr)
        return *(((uint32_t*)vbstr)-1); // ændret fra 32bit 18/10-2021
    return 0;
}

// To be called from c
long VBstrlen2(char *vbstr) // size_t
{
    return *(((uint32_t*)vbstr)-1);
}

// det ser faktisk ud til at denne også giver det rigtige ihbvertfald 64-bit
long VBstrlen3(const char *vbstr) // size_t
{
    return strlen(vbstr);
}

// chrstring er input, vbstring er output.
// Skal kaldes fra c
// Konverterer en c-streng (enkelt byte \0 terminated) til en VBA streng)
void CharToVbs(char* chrstring, char* vbstring)
{
    size_t strsiz;
    size_t strsizvb;
    strsiz=strlen(chrstring); // ændret fra sizeof da den gav forkert
    strsizvb=VBstrlen(vbstring);
    if (strsizvb<strsiz && strsizvb>0)
    {
        strsiz=strsizvb;
    }
    //    memset(vbstring, '\0', strsiz); // nulstring streng
    strcpy(vbstring,chrstring);
    // følgende skal nok ikke køres da det ændrer længden af strengen fra VB
    // Det kan give måske give memory leak. Det er bedre at ændre længden i VB
    //    *(vbstring-4)=(strsiz >> (8*0)) & 0xff;  //strsiz%255;
    //    *(vbstring-3)=(strsiz >> (8*1)) & 0xff;//(strsiz%65536-*(vbstring-4))/65536;
    //    *(vbstring-2)=(strsiz >> (8*2)) & 0xff;
    //    *(vbstring-1)=(strsiz >> (8*3)) & 0xff;
}

long ReadOutputIfAny(char *outstring)
{
    long ReadSize;
    ReadSize = strlen(outstring);
    
    long outlen=0;
    outlen=read(0,outstring,ReadSize); // 0 er stdin. return 0 if EOF, -1 if error. buffer pointer is automatically incremented by the bytes read according to doc, but I dont see it does. pread does not alter buffer pointer, but I cant make it work
    return outlen;
}

// 32bit version
long ReadOutputIfAny2(char *outstring)
{
    long ReadSize = VBstrlen(outstring); //strlen(outstring); doesn't work for string Malloc'ed
    if (ReadSize<1)
        return -1;
    long outlen=0;
    outlen=read(0,outstring+outlen,ReadSize); // 0 er stdin. return 0 if EOF, -1 if error. buffer pointer is automatically incremented by the bytes read according to doc, but I dont see it does. pread does not alter buffer pointer, but I cant make it work
    if (outlen>=ReadSize)
        outlen=ReadSize;
    return outlen;
}

long ReadOutputIfAnyVB(char *outstring)
{
    return ReadOutputIfAny(outstring);
}

// fetches the output from Maxima which is stored in a buffer
// The output is placed in outstring which must be large enough to hold the output
// Excess output is left in the buffer
// Returns when stoptext is found in the output, waittime is reached.
// returns the number of characters printed to outstring.
// returns 0 or a negative number if the expected (%ix) didn't come
// the negative number is the length of characters printed if any
long GetLastOutput(char *outstring, long WaitTime)
{
    char *stoptext=malloc(8); // "(%i";
    int digits[8];
    int p;
    long mi; // ændret fra int 18/10-2021
    int i=0;
    if (IsMaximaRunning()<1)
        return -1;
    
    stoptext[0]='(';
    stoptext[1]='%';
    stoptext[2]='i';
    mi=mindex;
    p=(int) floor(log10(mindex)); // no of digits
    while (i<=p)
    {
        //        stoptext[3+i]=mindex % 10*(p-i+1) + 48;
        digits[i]=(mi/pow(10,p-i));
        mi=mi-digits[i]*pow(10,p-i);
        stoptext[3+i]= digits[i]+ 48;
        i++;
    }
    stoptext[4+p]=')';
    stoptext[5+p]='\0';
    //    time_t waittime=(time_t)waittime2;
    
    //int ReadSize=strlen(outstring)+10;
    long ReadSize = strlen(outstring); //VBstrlen(outstring); doesn't work for string Malloc'ed
    if (ReadSize<1)
        return -1;
    //    const int MaxNoOfReads=1;
    char *buffer=malloc(ReadSize+10);
    char *ReturnString=malloc(ReadSize+10); //maxnoofread*readsize
    time_t StartTime;
    //    time_t waittime=1; // Det virker hvis man sætter tiden her men ikke via parameret. Datatype fra vb problem?
    StartTime = time(NULL);
    
    long n=0;
    long l=ReadSize;
    long outlen=0;
    struct timeval timeout;
    fd_set fdset;
    //    int waittimeloc=WaitTime;
    
    while (InString(0, outstring, stoptext)<0 && n<WaitTime)  //difftime(time(NULL),StartTime)<(double)waittime) keep reading until stoptext is found or waittime passed
    {
        timeout.tv_sec = 1;//waittime;  // seconds virker ikke med parameter. fdset og timeout skal nulstilles når select har været kørt
        timeout.tv_usec = 0; // microseconds
        FD_ZERO(&fdset);
        FD_SET(fileno(stdin), &fdset);
        n++;
        
        l=ReadSize;
        if (select(1,&fdset,NULL,NULL,&timeout) > 0) //Returns when there is something to read. -1 is error. 0 if timeout. >0 indicates something to read. The first parameter is the max filedescriptor in the fdset (filedescriptors are integers). In this case there is only one fd in the set. it is 0, hence the paramter is 0+1=1. Maybe FD_SETSIZE+1 is more general. The third parameter is a fdset to be tested for able to write.
        {
            //            while( l==ReadSize && n<MaxNoOfReads) // read whats waiting in the buffer
            //                while( l==ReadSize ) // read whats waiting in the buffer
            //            {
            l=read(0,outstring+outlen,ReadSize-outlen); // 0 er stdin. return 0 if EOF, -1 if error. buffer pointer is automatically incremented by the bytes read according to doc, but I dont see it does. pread does not alter buffer pointer, but I cant make it work
            if (l>0)
            {
                //                if (outlen<ReadSize) // there is more room in the returnstring
                //                {
                //                    if ((outlen+l)>ReadSize) // if the buffer is holding more than there is room for in returnstring discard the rest
                //                    {
                //                        l=ReadSize-outlen;
                //                    }
                //                    memcpy(ReturnString+outlen,buffer,l);
                outlen=outlen+l;
                if (outlen>=ReadSize)
                {
                    outlen=ReadSize;
                    break;
                }
                //                }
                //                    n++;
            }
            else if(l==-1)
            {
                break;
            }
            //            }
        }
        //        if(InString(0, buffer, stoptext)<0 && time(NULL)-StartTime<waittime)
        //        {
        //            usleep(200000);
        //        }
    }
    
    outstring[outlen]='\0';
    //    strcpy(outstring,ReturnString);
    //    outstring[outlen]='\0';
    //    CharToVbs(ReturnString, outstring);
    free(ReturnString);
    free(buffer);
    if (InString(0, outstring, stoptext)<0) // if timed out without finding stoptext return negative value. the negative value of the length of the output if any.
    {
        free(stoptext);
        return -outlen;
    }
    else
    {
        free(stoptext);
        return outlen;
    }
}


long GetLastOutputVB(char *outstring,const long WaitTime)
{
    //    int w=WaitTime;
    return GetLastOutput(outstring, WaitTime);
    //    if (w>3)
    //    {
    //        return 10;
    //    }
    //    else
    //    {
    //        return 5;
    //    }
    
}

// returns
// 0 - Maxima was closed
// 1 - Maxima was force closed
// 2 - Maxima was already closed/not running
long CloseMaxima()
{
    char command[]="\n quit();\n";  // Det forreste \n er for evt at afslutte en ventende kommando der ikke er sendt
    errno=0;
    write(fdWrite[PIPE_WRITE], command, strlen(command));
    if (errno==0)
    {
        close(fdWrite[PIPE_WRITE]); // sends EOF
        close(fdRead[PIPE_READ]);
        usleep(200000);
        kill(ChildPID,SIGQUIT);
        ChildPID=0;
    }
    else
    {
        if (ChildPID>0)
        {
            kill(ChildPID,SIGQUIT);
            ChildPID=0;
            return 1;
        }
        return 2;
    }
    //    free(MaximaOutString);
    MaximaStarted=0;
    return 0;
}

// Testet virker fra VBA
int OpenSendClose(const char *command, char *outstring)
{
    StartMaxima();
    //    sleep(3);
    GetLastOutput(outstring, 10);// venter max 10 s på at maxima returnerer klar
    SendToMaxima(command);
    //    sleep(3);
    GetLastOutput(outstring, 2); // venter max 2 s på at Maxima sender svar
    CloseMaxima();
    return 0;
}

//
//Sends command to maxima and returns the result in outstring
// The outstring must be declared in VBA and be of length 1000
long SendToMaximaWait(char *command, char *outstring, const long WaitTime)
{
    SendToMaxima(command);
    //    usleep(2000);
    //    return GetLastOutput(outstring, WaitTime,VBstrlen(outstring));  // til VB
    return GetLastOutput(outstring, WaitTime);
}

// Denne kan køres fra VB, men ikke SendToMaxima.
//Meget mærkeligt da den bare kører samme funktion
//SendToMaximaVB kan køres både i c og fra VB så endnu mere mærkeligt
long SendToMaximaVB(char *command)
{
    SendToMaxima(command);
    return 1;
}

// Starts Maxima as a hidden process and returns control when Maxima is ready to accept commands
long StartMaximaWait(const long WaitTime)
{
    if (IsMaximaRunning()>0)
    {
        return 2; // already running
    }
    char *outstring=malloc(1010); // simuler en VBstreng
    memset(outstring, ' ', 1000); // nulstring streng
    outstring[0]=200;
    outstring[1]=3; // 512+256
    outstring[2]='x';
    outstring[3]='x';
    outstring[968]='\0';

    errno=0;
    long ern;
    ern=StartMaxima();
    if (ern>0) // && errno==0)
    {
        usleep(50000); // wait for maxima to fail
        if (IsMaximaRunning()>0)
            GetLastOutput(outstring+4, WaitTime);// venter på at maxima returnerer klar
    }
    free(outstring);
    return ern;
}

// Starts Maxima as a hidden process and returns control when Maxima is ready to accept commands
// outputs the introprompt from Maxima to outstring
long StartMaximaWait2(long WaitTime,char *outstring)
{
    errno=0;
    StartMaxima();
    
    GetLastOutput(outstring, WaitTime);
    return errno;
}

long testlong2(){
    return 6;
}

void testlong3(){
    
}

long testlong(){
    testlong3();
    return 5;
}




