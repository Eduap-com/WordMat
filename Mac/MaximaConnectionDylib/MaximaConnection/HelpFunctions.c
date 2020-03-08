//
//  HelpFunctions.c
//  MaximaConnection
//
//  Created by NyborgGym on 08/10/13.
//  Copyright (c) 2013 EDUAP. All rights reserved.
//

#include <stdio.h>
#include <string.h>
#include <stdlib.h>

//Searches searchstring for findstring from startpos (0 is start of string)
// returns -1 if string is not found or findstring is larger than searchstring
int InString(int StartPos,char *SearchString, const char *FindString)
{
    int ls,lf,dl;
    int i=StartPos;
    int comp=1;
    ls = strlen(SearchString);
    if (ls==0)
        return -1;
    lf = strlen(FindString);
    if (lf==0)
        return -1;
    dl=ls-lf;
    if (lf>ls)
        return -1;
    while (comp!=0 & i<=dl+1)
    {
        comp=memcmp(SearchString+i,FindString,lf);
        i++;
    };
    if (i>dl+1)
        return -1;
    else
        return i-1;
}

int CountInString(int StartPos, char *SearchString, const char *FindString)
{
    int c=0;
    int p=StartPos;
    
    while ((p=InString(p,SearchString,FindString)+1)>0) {
        c++;
    }
    
    return c;
}

// Removes all trimchar characters to left of trimstring
// The string is terminated by 0, but memory is not reallocated
char* TrimStringLeft(char *trimstring, char trimchar)
{
    char *tempstring=malloc(strlen(trimstring));
    tempstring=trimstring;
    int i=0;
    while (*(trimstring+i)==trimchar)
        i++;
    if (i>0)
    {
        memmove(trimstring,tempstring+i , 7);
    }
    return trimstring;
}

char* TrimStringRight(char *trimstring, char trimchar)
{
    int i=strlen(trimstring)-1;
    while (*(trimstring+i)==trimchar)
        i--;
    *(trimstring+i+1)='\0';
    return trimstring;
}

char* TrimString(char *trimstring, char trimchar)
{
    TrimStringRight(trimstring, trimchar);
    TrimStringLeft(trimstring,trimchar);
    return trimstring;
}