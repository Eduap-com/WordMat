//
//  HelpFunctions.h
//  MaximaConnection
//
//  Created by NyborgGym on 07/11/13.
//  Copyright (c) 2013 EDUAP. All rights reserved.
//

#ifndef MaximaConnection_HelpFunctions_h
#define MaximaConnection_HelpFunctions_h

int CountInString(int StartPos, char *SearchString, const char *FindString);
int InString(int StartPos,char *SearchString, const char *FindString);
char* TrimString(char *trimstring, char trimchar);
char* TrimStringLeft(char *trimstring, char trimchar);
char* TrimStringRight(char *trimstring, char trimchar);

#endif
