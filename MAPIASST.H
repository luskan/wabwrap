////////////////////////////////////////////////////////////////////////////////
//	MAPIASST.H
////////////////////////////////////////////////////////////////////////////////
//	This Header file is intended to be a general purpose
//	MAPI debug helper file.
//
//	MAPI_ASSERT will give you basic error output and kick you into the
//	debugger at the point where the error occurred.
//
//	MAPI_ASSERT_EX will give you all of the information that ASSERT does 
//	but it will also print the meaning of the HRESULT error code.
//	Again, you will be kicked back into the debugger at the point 
//	of error.
//
//	SHOWTABLE will use the tblvu32.dll function ViewMapiTable() 
//	to display the passed in MAPITable in a nice GUI Dialog.
//	(tblvu32.dll is downloadable from : 
//	ftp.microsoft.com\DEVELOPR\MAPI\SAMPLES\MFCAPPS.ZIP )
//
//	Written by Sam Charchian 11/18/97.  Microsoft Developer Support
////////////////////////////////////////////////////////////////////////////////

#ifdef _DEBUG   // entire file
#include  <stdio.h>		// For sprintf().	

// We need an ASSERT Macro here
#define MAPI_ASSERT(hr) \
	if (hr) \
	{ \
		char	sOutStr[255];\
		sprintf(sOutStr, "**ASSERT! File : %s\n", __FILE__); \
		OutputDebugString(sOutStr); \
		sprintf(sOutStr, "          Line : %u\n", __LINE__); \
		OutputDebugString(sOutStr); \
		sprintf(sOutStr, "         Value : 0x%x\n", hr); \
		OutputDebugString(sOutStr); \
		_asm int 3 \
	} 

//////////////////////////////////////////////////////////////////////////////////////////////////
// Stolen from MAPICode.H
#define MAKE_MAPI_SCODE(sev,fac,code) \
    ((SCODE) (((unsigned long)(sev)<<31) | ((unsigned long)(fac)<<16) | ((unsigned long)(code))) )

/* The following two macros are used to build OLE 2.0 style sCodes */

#define MAKE_MAPI_E( err )  (MAKE_MAPI_SCODE( 1, FACILITY_ITF, err ))
#define MAKE_MAPI_S( warn ) (MAKE_MAPI_SCODE( 0, FACILITY_ITF, warn ))
//////////////////////////////////////////////////////////////////////////////////////////////////

// This assert Macro will report MAPI ERROR MSGS as well as the normal ASSERT info.
#define MAPI_ASSERT_EX(hr) \
	if (hr) \
	{ \
		char	sOutStr[255];\
		sprintf(sOutStr, "**ASSERT! File : %s\n", __FILE__); \
		OutputDebugString(sOutStr); \
		sprintf(sOutStr, "          Line : %u\n", __LINE__); \
		OutputDebugString(sOutStr); \
		sprintf(sOutStr, "         Value : 0x%x\n", hr); \
		OutputDebugString(sOutStr);\
		switch(hr) \
		{\
			case (MAKE_MAPI_E( 0x102 )): \
				sprintf(sOutStr, "       Meaning : MAPI_E_NO_SUPPORT\n"); \
				break; \
			case (MAKE_MAPI_E( 0x103 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_BAD_CHARWIDTH\n"); \
				break; \
			case (MAKE_MAPI_E( 0x105 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_STRING_TOO_LONG\n"); \
				break; \
			case (MAKE_MAPI_E( 0x106 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_UNKNOWN_FLAGS\n"); \
				break; \
			case (MAKE_MAPI_E( 0x107 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_INVALID_ENTRYID\n"); \
				break; \
			case (MAKE_MAPI_E( 0x108 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_INVALID_OBJECT\n"); \
				break; \
			case (MAKE_MAPI_E( 0x109 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_OBJECT_CHANGED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x10A )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_OBJECT_DELETED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x10B )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_BUSY\n"); \
				break; \
			case (MAKE_MAPI_E( 0x10D )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NOT_ENOUGH_DISK\n"); \
				break; \
			case (MAKE_MAPI_E( 0x10E )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NOT_ENOUGH_RESOURCES\n"); \
				break; \
			case (MAKE_MAPI_E( 0x10F )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NOT_FOUND\n"); \
				break; \
			case (MAKE_MAPI_E( 0x110 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_VERSION\n"); \
				break; \
            case (MAKE_MAPI_E( 0x111 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_LOGON_FAILED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x112 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_SESSION_LIMIT\n"); \
				break; \
			case (MAKE_MAPI_E( 0x113 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_USER_CANCEL\n"); \
				break; \
			case (MAKE_MAPI_E( 0x114 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_UNABLE_TO_ABORT\n"); \
				break; \
			case (MAKE_MAPI_E( 0x115 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NETWORK_ERROR\n"); \
				break; \
			case (MAKE_MAPI_E( 0x116 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_DISK_ERROR\n"); \
				break; \
			case (MAKE_MAPI_E( 0x117 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_TOO_COMPLEX\n"); \
				break; \
			case (MAKE_MAPI_E( 0x118 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_BAD_COLUMN\n"); \
				break; \
			case (MAKE_MAPI_E( 0x119 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_EXTENDED_ERROR\n"); \
				break; \
			case (MAKE_MAPI_E( 0x11A )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_COMPUTED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x11B )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_CORRUPT_DATA\n"); \
				break; \
			case (MAKE_MAPI_E( 0x11C )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_UNCONFIGURED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x11D )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_FAILONEPROVIDER\n"); \
				break; \
			case (MAKE_MAPI_E( 0x11E )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_UNKNOWN_CPID\n"); \
				break; \
			case (MAKE_MAPI_E( 0x11F )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_UNKNOWN_LCID\n"); \
				break; \
			case (MAKE_MAPI_E( 0x120 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_PASSWORD_CHANGE_REQUIRED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x121 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_PASSWORD_EXPIRED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x122 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_INVALID_WORKSTATION_ACCOUNT\n"); \
				break; \
			case (MAKE_MAPI_E( 0x123 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_INVALID_ACCESS_TIME\n"); \
				break; \
			case (MAKE_MAPI_E( 0x124 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_ACCOUNT_DISABLED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x200 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_END_OF_SESSION\n"); \
				break; \
			case (MAKE_MAPI_E( 0x201 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_UNKNOWN_ENTRYID\n"); \
				break; \
			case (MAKE_MAPI_E( 0x202 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_MISSING_REQUIRED_COLUMN\n"); \
				break; \
			case (MAKE_MAPI_S( 0x203 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_W_NO_SERVICE\n"); \
				break; \
			case (MAKE_MAPI_E( 0x301 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_BAD_VALUE\n"); \
				break; \
			case (MAKE_MAPI_E( 0x302 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_INVALID_TYPE\n"); \
				break; \
			case (MAKE_MAPI_E( 0x303 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_TYPE_NO_SUPPORT\n"); \
				break; \
			case (MAKE_MAPI_E( 0x304 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_UNEXPECTED_TYPE\n"); \
				break; \
			case (MAKE_MAPI_E( 0x305 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_TOO_BIG\n"); \
				break; \
			case (MAKE_MAPI_E( 0x306 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_DECLINE_COPY\n"); \
				break; \
			case (MAKE_MAPI_E( 0x307 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_UNEXPECTED_ID\n"); \
				break; \
			case (MAKE_MAPI_S( 0x380 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_W_ERRORS_RETURNED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x400 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_UNABLE_TO_COMPLETE\n"); \
				break; \
			case (MAKE_MAPI_E( 0x401 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_TIMEOUT\n"); \
				break; \
			case (MAKE_MAPI_E( 0x402 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_TABLE_EMPTY\n"); \
				break; \
			case (MAKE_MAPI_E( 0x403 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_TABLE_TOO_BIG\n"); \
				break; \
			case (MAKE_MAPI_E( 0x405 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_INVALID_BOOKMARK\n"); \
				break; \
			case (MAKE_MAPI_S( 0x481 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_W_POSITION_CHANGED\n"); \
				break; \
			case(MAKE_MAPI_S( 0x482 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_W_APPROX_COUNT\n"); \
				break; \
			case (MAKE_MAPI_E( 0x500 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_WAIT\n"); \
				break; \
			case (MAKE_MAPI_E( 0x501 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_CANCEL\n"); \
				break; \
			case (MAKE_MAPI_E( 0x502 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NOT_ME\n"); \
				break; \
			case(MAKE_MAPI_S( 0x580 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_W_CANCEL_MESSAGE\n"); \
				break; \
			case (MAKE_MAPI_E( 0x600 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_CORRUPT_STORE\n"); \
				break; \
			case (MAKE_MAPI_E( 0x601 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NOT_IN_QUEUE\n"); \
				break; \
			case (MAKE_MAPI_E( 0x602 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NO_SUPPRESS\n"); \
				break; \
			case (MAKE_MAPI_E( 0x604 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_COLLISION\n"); \
				break; \
			case (MAKE_MAPI_E( 0x605 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NOT_INITIALIZED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x606 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NON_STANDARD\n"); \
				break; \
			case (MAKE_MAPI_E( 0x607 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_NO_RECIPIENTS\n"); \
				break; \
			case (MAKE_MAPI_E( 0x608 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_SUBMITTED\n"); \
				break; \
			case (MAKE_MAPI_E( 0x609 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_HAS_FOLDERS\n"); \
				break; \
			case (MAKE_MAPI_E( 0x60A )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_HAS_MESSAGES\n"); \
				break; \
			case (MAKE_MAPI_E( 0x60B )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_FOLDER_CYCLE\n"); \
				break; \
			case(MAKE_MAPI_S( 0x680 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_W_PARTIAL_COMPLETION\n"); \
				break; \
			case (MAKE_MAPI_E( 0x700 )):  \
				sprintf(sOutStr, "       Meaning : MAPI_E_AMBIGUOUS_RECIP\n"); \
				break; \
			default: \
				sprintf(sOutStr, "       Meaning : Meaning Unknown!\n"); \
		}\
		OutputDebugString(sOutStr); \
		_asm int 3 \
	} 

// SHOWTABLE will display a MAPITABLE in a dialog box.
// It uses tblvu32.dll's ViewMapiTable function to 
// show the passed in MAPITable pointer.
#define SHOWTABLE(MAPITable) \
{ \
	if (MAPITable) \
	{ \
		HINSTANCE hInstTableVu = LoadLibrary("tblvu32.dll"); \
		ULONG (PASCAL *lpfnViewMAPITable)(LPMAPITABLE FAR *, HWND); \
		(FARPROC&)lpfnViewMAPITable = GetProcAddress(hInstTableVu, "ViewMapiTable"); \
		lpfnViewMAPITable((LPMAPITABLE FAR *)&MAPITable, NULL); \
		FreeLibrary (hInstTableVu); \
	} \
	else \
	{ \
		MessageBox((HWND)GetActiveWindow(), "The Table Pointer Passed to SHOWTABLE is NULL", \
			"MAPI ASSERT ROUTINES", MB_OK | MB_ICONEXCLAMATION); \
		OutputDebugString("The Table Pointer Passed to SHOWTABLE is NULL! \n"); \
	} \
}

#endif // _DEBUG

///////////////////////////////////////////////////////////////////////////////
//	We now need to define our MACROs to make the compiler happy
//	when building in Release Mode.
#ifndef MAPI_ASSERT
#define MAPI_ASSERT 
#endif // MAPI_ASSERT

#ifndef MAPI_ASSERT_EX
#define MAPI_ASSERT_EX 
#endif // MAPI_ASSERTX

#ifndef SHOWTABLE
#define SHOWTABLE
#endif // SHOWTABLE
///////////////////////////////////////////////////////////////////////////////
