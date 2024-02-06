//
//  Header.h
//  MaximaConnection
//
//  Created by NyborgGym on 16/10/13.
//  Copyright (c) 2013 EDUAP. All rights reserved.
//

#ifndef MaximaConnection_Header_h
#define MaximaConnection_Header_h

long SetMaximaPath(const char *SetPath);
long StartMaxima();
long StartShell();
long IsMaximaRunning();
void SetCommandIndex(const long ind);
long GetCommandIndex();
long SendToMaxima(char* command);
long SendToMaximaVB(char *command);
long VBstrlen(const char *vbstr);
void CharToVbs(char* chrstring, char* vbstring);
long GetLastOutput(char *outstring, const long waittime);
long GetLastOutputVB(char *outstring, const long waittime);
long ReadOutputIfAny(char *outstring);
long ReadOutputIfAnyVB(char *outstring);
long CloseMaxima();
int OpenSendClose(const char *command, char *outstring);
long SendToMaximaWait(char *command, char *outstring, const long WaitTime);
long StartMaximaWait(const long WaitTime);  // Starts Maxima as a hidden process and returns control when Maxima is ready to accept commands
long StartMaximaWait2(const long WaitTime,char *outstring); // Starts Maxima as a hidden process and returns control when Maxima is ready to accept commands, outputs the introprompt from Maxima to outstring

long ForceDivZero();
void testlong3();
long testlong2();
long testlong();


#endif
