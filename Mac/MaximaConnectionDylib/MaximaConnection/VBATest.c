//
//  VBATest.c
//  MaximaConnection
//
//  Created by NyborgGym on 06/10/13.
//  Copyright (c) 2013 EDUAP. All rights reserved.
//

#include <stdio.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>
#include <arpa/inet.h>
#include <sys/types.h>
#include <netinet/in.h>
#include <sys/socket.h>
int sum;

long TestFunc()
{
    return 1;
}

int AddThem( int a, int b ) {
    return a+b;
}

int GetInt(int n, int n2)
{
    return AddThem(n,n2);
}

int SetString(char *outstring)
{
    char *ts="Mejsa";
//    memset(outstring, '\0', sizeof(outstring)); // nulstring streng
    //    chrtovbs(ts, outstring);
    //    outstring[0]='s';
    //    outstring[1]='r';
    strcpy(outstring,ts);
    *(outstring-1)=0;
    *(outstring-2)=0;
    *(outstring-3)=0;
    *(outstring-4)=strlen(ts);
    return 4;
}
int SetString2(char *outstring)
{
//    memset(outstring, '\0', sizeof(outstring)); // nulstring streng
    //    chrtovbs(ts, outstring);
    //    outstring[0]='s';
    //    outstring[1]='r';
    
    *(outstring-1)=0;
    *(outstring-2)=0;
    *(outstring-3)=0;
    *(outstring-4)=5;
    *(outstring)='H';
    *(outstring+1)='e';
    *(outstring+2)='j';
    *(outstring+3)='s';
    *(outstring+4)='a';
    *(outstring+5)=0;
    return 5;
}

int GetLength(char *s)
{
    return strlen(s);
}
