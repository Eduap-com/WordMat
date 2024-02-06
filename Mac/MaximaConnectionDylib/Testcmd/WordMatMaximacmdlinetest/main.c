//
//  MaximaFunctions.c
//  MaximaConnection
//
//  Created by NyborgGym on 19/09/13.
//  Copyright (c) 2013 EDUAP. All rights reserved.
//

// Der skal være en return i hver funktion. Typisk i slutningen. Ellers går Word ned.
// Det gælder også void-funktioner. De skal afsluttes med return ;

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
#include "HelpFunctions.h"
#include "MaximaFunctions.h"

int Test1()
{
    printf("Starting...\n");

    char *outs=malloc(1000); // simuler en output streng fra Maxima
    memset(outs, ' ', 1000); // nulstring streng
    outs[0]=200; // De første 4 bytes angiver størrelsen af strengen i VB - men ikke i c.
    outs[1]=3;
    outs[2]=0;
    outs[3]=0;
    int i=4;
    long j;
    while (i<=967)
    {
        outs[i]=' ';
        i++;
    }
    outs[967]='\0';

//    SetMaximaPath("/Applications/WordMat/Maxima");
//    SetMaximaPath("/Library/Application support/Microsoft/Office365/User Content/Add-ins/WordMat/Maxima.app/Contents/Resources/Maxima.sh");
    SetMaximaPath("""/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/Maxima.app/Contents/Resources/maxima.sh""");
//    SetMaximaPath("/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/Maxima");
    printf("VBstrlen(outs)= %li \n",VBstrlen(outs+4));
    printf("StartMaxima()= %li \n",StartShell());
    usleep(2000000);
//    j=ReadOutputIfAny(outs+4);
    //    SendToMaxima("'ls'");
    //    usleep(1000000);
        j=ReadOutputIfAny(outs+4);
        printf("LS length:\n %li \n",j);

    //   printf("LS:\n %s \n",outs+4);
//    SendToMaxima("'/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/Maxima.app/Contents/Resources/maxima.sh'");
    SendToMaxima("ls\n");
    usleep(2000000);
//        GetLastOutput(outs+4,5);
    j=ReadOutputIfAny(outs+4);
    printf("Velkomststreng fra Maxima:\n %s \n",outs+4);

    printf("IsMaximaRunning: %li \n",IsMaximaRunning());

    i=4;
    while (i<=967)
    {
        outs[i]=' ';
        i++;
    }
    outs[968]='\0';
//    SendToMaxima("2+3;");
//    j=ReadOutputIfAny(outs+4);

    GetLastOutput(outs+4,5);
    printf("Resultat af simpel beregning:\n %s \n",outs+4);

    return 0;
}




int Test2()
{
    printf("Starting...\n");

    char *outs=malloc(1000); // simuler en output streng fra Maxima
    memset(outs, ' ', 1000); // nulstring streng
    outs[0]=200; // De første 4 bytes angiver størrelsen af strengen i VB - men ikke i c.
    outs[1]=3;
    outs[2]=0;
    outs[3]=0;
    int i=4;
    while (i<=967)
    {
        outs[i]=' ';
        i++;
    }
    outs[968]='\0';
    
    SetMaximaPath("""/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/Maxima.app/Contents/Resources/rmaxima.sh""");
    printf("VBstrlen(outs)= %li \n",VBstrlen(outs+4));
    StartMaxima();
    usleep(3000000);
    ReadOutputIfAny(outs+4);
    printf("Velkomststreng fra Maxima:\n %s \n",outs+4);

    printf("IsMaximaRunning: %li \n",IsMaximaRunning());

    i=4;
    while (i<=967)
    {
        outs[i]=' ';
        i++;
    }
    outs[968]='\0';
    SendToMaxima("2+3;");

    GetLastOutput(outs+4,5);
    printf("Resultat af simpel beregning:\n %s \n",outs+4);

    return 0;
}

int Test12()
{
    
    printf("Starting...\n");

    char *outs=malloc(1000); // simuler en output streng fra Maxima
    memset(outs, ' ', 1000); // nulstring streng
    outs[0]=200;
    outs[1]=3;
    outs[2]=0;
    outs[3]=0;
    outs[968]='\0';

    printf("VBstrlen(outs)= %li \n",VBstrlen(outs+4));
    
//    char *command="2*64;\n 4+5;\n";
    char *command="factor(988776542345678987654345678914323784393756346329364420674526385);\n";
//    char *command=malloc(10);
//    memset(command, ' ', 10); // nulstring streng
//    command[0]=5;
//    command[1]=0;
//    command[2]=0;
//    command[3]=0;
//    command[4]='1';
//    command[5]='+';
//    command[6]='4';
//    command[7]=';';
//    command[8]='\n';
//    command[9]='\0';
//    char *command=malloc(30);
//    CharToVbs(command2, command);
//    free(command2);
    
    long waittime=1;
//    StartMaximaWait2(10,outs+4);
    printf("StartMaximaWait return %li \n",StartMaximaWait(15));
    printf("intro string: %s \n",outs+4);
    SendToMaximaWait(command,outs+4,waittime);
//    SendToMaxima(command+4);
    printf("Outputstring:\n %s \n",outs+4);
    GetLastOutput(outs+4,waittime);
    printf("Outputstring:\n %s \n",outs+4);
    waittime=10;
    GetLastOutput(outs+4,waittime);
    printf("Outputstring:\n %s \n",outs+4);
    
    CloseMaxima();

    free(outs);
    printf("\nend proc\n");
    return 0;

}


int main( int argc, const char* argv[] )
{
    Test2();
}

