//
//File Maker plugin
//Function: read data from given COM port
//

#include "FMWrapper/FMXTypes.h"
#include "FMWrapper/FMXText.h"
#include "FMWrapper/FMXFixPt.h"
#include "FMWrapper/FMXData.h"
#include "FMWrapper/FMXCalcEngine.h"

#include<Windows.h>
#include <string>
#include <vector>
#include <atlbase.h>

//Exported plug-in functions ===========================================================================

static const char* kFMsp("FMsp");

//plug-in functions ID, min and max amount of parameters /my assumption/
enum {kFMsp_ReadPortID = 100, kFMsp_ReadPortMin = 5, kFMsp_ReadPortMax = 5};

//plug-in function information to show in File Maker application
static const char* kFMsp_ReadPortName("FMsp_ReadPort");
static const char* kFMsp_ReadPortDefinition("FMsp_ReadPort( PortName, Baudrate, ByteSize, StopBits, Parity )");
static const char* kFMsp_ReadPortDescription("Get a value from a given serial port");

// Exported plug-in script step

enum {kFMsp_ReadPortScriptStepID = 200, kFMsp_ReadPortScriptStepMin = 1, kFMsp_ReadPortScriptStepMax = 1};
static const char* kFMsp_ReadPortScriptStepName( "Read Port" );
static const char* kFMsp_ReadPortScriptStepDefinition( "<PluginStep>"
															"<Parameter Type=\"target\" Label=\"Destination\" ShowInLine=\"true\"/>"
															"<Parameter Type=\"calc\" DataType=\"text\" ShowInLine=\"true\" Label=\"Port name\" ID= \"0\" />"
															"<Parameter Type=\"calc\" DataType=\"text\" ShowInLine=\"true\" Label=\"BaudRate\" ID= \"1\" />"
															"<Parameter Type=\"calc\" DataType=\"text\" ShowInLine=\"true\" Label=\"ByteSize\" ID= \"2\" />"
															"<Parameter Type=\"calc\" DataType=\"text\" ShowInLine=\"true\" Label=\"StopBits\" ID= \"3\" />"
															"<Parameter Type=\"calc\" DataType=\"text\" ShowInLine=\"true\" Label=\"Parity\" ID= \"4\" />"
													   "</PluginStep>");
static const char* kFMsp_ReadPortScriptStepDescription( "Get a value from a given serial port" );

//code to implement plug-in function/script step /my function and scrip step make the same so I use this part of a code for both/
static FMX_PROC(fmx::errcode) Do_FMsp_ReadPort( short, const fmx::ExprEnv& environment, const fmx::DataVect& dataVect, fmx::Data& results )
{
	fmx::errcode errorResult(960);
	fmx::TextUniquePtr outText;

	if (dataVect.Size() == 5)
	{
		const fmx::Data& dat(dataVect.At(0));

		//to get a port name from the parameters
		const fmx::Text& t_portName(dataVect.At(0).GetAsText());
		double size = t_portName.GetSize();
		char *buffer = new char[size + 1];
		t_portName.GetBytes(buffer, size + 1);
		double buffer_size = strlen(buffer);
		wchar_t *portName = new wchar_t[buffer_size + 1];
		MultiByteToWideChar(CP_ACP, 0, buffer, -1, portName, buffer_size + 1);

		//to get a baudrate from the parameters
		const fmx::int32 baudrate(dataVect.At(1).GetAsNumber().AsLong());
		
		//to get a bytesite from the parameters
		const fmx::int16 bytesize(dataVect.At(2).GetAsNumber().AsLong());

		//to get a stopbits from the parameters (1 or 2)
		fmx::int16 stopbits(dataVect.At(3).GetAsNumber().AsLong());
		if (stopbits == 1) {
			stopbits = 0;
		}

		//to get a parity from the parameters
		const fmx::Text& t_parity(dataVect.At(4).GetAsText());
		size = t_parity.GetSize();
		char* buffer_p = new char[size + 1];
		t_parity.GetBytes(buffer_p, size + 1);
		std::string s_parity = buffer_p;
		short parity = 0;
		if (s_parity == "NONE") {
			parity = 0;
		}
		if (s_parity == "ODD") {
			parity = 1;
		}
		if (s_parity == "EVEN") {
			parity = 2;
		}
		if (s_parity == "MARK") {
			parity = 3;
		}
		if (s_parity == "SPACE") {
			parity = 4;
		}

		//Opening the serial port
		HANDLE hSerial;
		hSerial = CreateFile(portName, GENERIC_READ | GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
		if (hSerial == INVALID_HANDLE_VALUE) {
			if (GetLastError() == ERROR_FILE_NOT_FOUND) {
				//serial port does not exist
				outText->Assign("Port opening is failed");
				results.SetAsText(*outText, results.GetLocale());

				errorResult = fmx::ExprEnv::kPluginErrResult1;

				return errorResult;
			}
			//other error
		}

		//Setting Parameters
		DCB dcbSerialParams = { 0 };
		if (!GetCommState(hSerial, &dcbSerialParams)) {
			//Failed to get current serial parameters
		}
		//The baud rate at which the communications device operates
		dcbSerialParams.BaudRate = baudrate;

		//The number of bits in the bytes transmitted and received, 5-8
		dcbSerialParams.ByteSize = bytesize;

		//The number of stop bits to be used. ONESTOPBIT /0/, ONE5STOPBITS /1/, TWOSTOPBITS /2/ 
		dcbSerialParams.StopBits = stopbits;

		//The parity scheme to be used. NOPARITY /0/, ODDPARITY /1/, EVENPARITY /2/, MARKPARITY /3/, SPACEPARITY /4/
		dcbSerialParams.Parity = parity;

		if (!SetCommState(hSerial, &dcbSerialParams)) {
			//could not set serial port parameters
			outText->Assign("Could not set serial port parameters");
			results.SetAsText(*outText, results.GetLocale());
		}

		//Setting timeouts
		COMMTIMEOUTS timeouts = { 0 };
		timeouts.ReadIntervalTimeout = 50;
		timeouts.ReadTotalTimeoutConstant = 50;
		timeouts.ReadTotalTimeoutMultiplier = 10;
		timeouts.WriteTotalTimeoutConstant = 50;
		timeouts.WriteTotalTimeoutMultiplier = 10;
		if (!SetCommTimeouts(hSerial, &timeouts)) {
			//error
		}

		//A pointer to the variable that receives the number of bytes read/written when using a synchronous hFile parameter.
		DWORD dwBytesRead = 0;

		//Writing data
		char writeMessage[3] = { 'P', '\r\n'};
		if (!WriteFile(hSerial, writeMessage, 3, &dwBytesRead, NULL)) {
			//error
		}

		//Reading data
		const int n = 32;
		char szBuff[n + 1] = { 0 };
		if (!ReadFile(hSerial, szBuff, n, &dwBytesRead, NULL)) {
			//error
		}

		outText->Assign(szBuff, fmx::Text::kEncoding_UTF8);
		results.SetAsText(*outText, dat.GetLocale());
		errorResult = 0;

		//Closing down
		CloseHandle(hSerial);

		//to free allocated memory /my assumption/
		delete[] portName;
		delete[] buffer;
		delete[] buffer_p;

	}
	else
	{
		outText->Assign("Missed port settings (PortName, Baudrate, ByteSize, StopBits, Parity)");
		results.SetAsText(*outText, results.GetLocale());
	}
	return errorResult;
}

//plug-in function to get a list of COM ports

enum { kFMsp_PortListID = 300, kFMsp_PortListMin = 0, kFMsp_PortListMax = 0 };

static const char* kFMsp_PortListName("FMsp_PortList");
static const char* kFMsp_PortListDefinition("FMsp_PortList");
static const char* kFMsp_PortListDescription("Return list of available COM ports");

//plug-in script step to get a list of COM ports
enum { kFMsp_PortListScriptStepID = 400, kFMsp_PortListScriptStepMin = 0, kFMsp_PortListScriptStepMax = 0 };
static const char* kFMsp_PortListScriptStepName("Port List");
static const char* kFMsp_PortListScriptStepDefinition("<PluginStep>"
														"<Parameter Type=\"target\" Label=\"Destination\" ShowInLine=\"true\"/>"
													  "</PluginStep>");
static const char* kFMsp_PortListScriptStepDescription("Get a list of COM ports");

static FMX_PROC(fmx::errcode) Do_FMsp_PortList( short, const fmx::ExprEnv& environment, const fmx::DataVect& dataVect, fmx::Data& results )
{	
	//Detect the install ports by enumerating the values contained at the registry key HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM

	fmx::errcode errorResult(960);
	fmx::TextUniquePtr outText;
	const fmx::Locale* outLocale(&results.GetLocale());

	//list of ports
	std::vector<std::wstring> ports;

	ATL::CRegKey serialKey;

	//open the specified key
	serialKey.Open(HKEY_LOCAL_MACHINE, _T("HARDWARE\\DEVICEMAP\\SERIALCOMM"), KEY_QUERY_VALUE);

	//A pointer to a variable that receives the size of the key's longest value name, in Unicode characters.
	DWORD max = 0;

	//The size does not include the terminating null character. This parameter can be NULL.
	RegQueryInfoKey(serialKey, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, &max /*here*/, nullptr, nullptr, nullptr);

	//to take into account the terminating null character
	const DWORD maxChar = max + 1;

	//A pointer to a buffer that receives the name of the value as a null-terminated string /RegEnumValue function/
	//This buffer must be large enough to include the terminating null character.
	std::vector<TCHAR> valueName;
	valueName.resize(maxChar);

	//loop parameters
	bool bContinue = true;
	//The index of the value to be retrieved. This parameter should be zero for the first call to the RegEnumValue function and then be incremented for subsequent calls.
	DWORD index = 0;

	while (bContinue)
	{
		//A pointer to a variable that specifies the size of the buffer pointed to by the lpValueName /valueName.data()/ parameter, in characters.
		//When the function returns, the variable receives the number of characters stored in the buffer, not including the terminating null character.
		DWORD valueNameSize = maxChar;

		valueName[0] = _T('\0');

		bContinue = (RegEnumValue(serialKey, index, valueName.data(), &valueNameSize, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS);

		if (bContinue)
		{
			std::wstring sPortName;

			//First query for the size of the registry value
			//contains the size, in TCHARs, of the string retrieved, including a terminating null character
			ULONG nChars = 0;

			if (serialKey.QueryStringValue(valueName.data(), nullptr, &nChars) == ERROR_SUCCESS)
			{
				//Allocate enough bytes for the return value
				//+2 is to allow us to null terminate the data if required
				sPortName.resize(static_cast<size_t>(nChars) + 2);
				const DWORD allocatedSize = ((nChars + 2) * sizeof(TCHAR));

				//A pointer to a variable that receives a code indicating the type of data stored in the specified value.
				DWORD type = 0;

				//A pointer to a variable that specifies the size of the buffer pointed to by the lpData /here (LPBYTE)(sValue.data()) / parameter, in bytes.
				//When the function returns, this variable contains the size of the data copied to lpData.
				ULONG bytes = allocatedSize;


				//valueName - the name of the registry value.
				if (RegQueryValueEx(serialKey, valueName.data(), nullptr, &type, (LPBYTE)(sPortName.data()), &bytes) == ERROR_SUCCESS)
				{
					// Forcibly null terminate the data ourselves
					if (sPortName[(bytes / sizeof(TCHAR)) - 1] != _T('\0'))
					{
						sPortName[bytes / sizeof(TCHAR)] = _T('\0');
					}

					ports.push_back(sPortName);
				}
			}
			++index;
		}
	}

	for (int i = 0; i < ports.size(); i++) {

		fmx::TextUniquePtr portName;
		portName->AssignWide(ports[i].c_str());

		fmx::TextUniquePtr space;
		std::wstring w_space = L"\n";

		space->AssignWide(w_space.c_str());
		portName->AppendText(*space);

		outText->AppendText(*portName);
	}

	errorResult = 0;
	results.SetAsText(*outText, *outLocale);
	return errorResult;
}

//Do_PluginInit ===========================================================================

static fmx::ptrtype Do_PluginInit( fmx::int16 version )
{
	fmx::ptrtype                    result(static_cast<fmx::ptrtype>(kDoNotEnable));
	const fmx::QuadCharUniquePtr    pluginID(kFMsp[0], kFMsp[1], kFMsp[2], kFMsp[3]);
	fmx::TextUniquePtr              name;
	fmx::TextUniquePtr              definition;
	fmx::TextUniquePtr              description;
	fmx::uint32                     flags(fmx::ExprEnv::kDisplayInAllDialogs | fmx::ExprEnv::kFutureCompatible);

	//plug-in functions and script steps initialize /my assumption/
	if (version >= k150ExtnVersion)
	{
		name->Assign(kFMsp_ReadPortName, fmx::Text::kEncoding_UTF8);
		definition->Assign(kFMsp_ReadPortDefinition, fmx::Text::kEncoding_UTF8);
		description->Assign(kFMsp_ReadPortDescription, fmx::Text::kEncoding_UTF8);
		if (fmx::ExprEnv::RegisterExternalFunctionEx(*pluginID, kFMsp_ReadPortID, *name, *definition, *description, kFMsp_ReadPortMin, kFMsp_ReadPortMax, flags, Do_FMsp_ReadPort) == 0)
		{
			result = kCurrentExtnVersion;
		}
	}
	else if (version == k140ExtnVersion)
	{
		name->Assign(kFMsp_ReadPortName, fmx::Text::kEncoding_UTF8);
		definition->Assign(kFMsp_ReadPortDefinition, fmx::Text::kEncoding_UTF8);
		if (fmx::ExprEnv::RegisterExternalFunction(*pluginID, kFMsp_ReadPortID, *name, *definition, kFMsp_ReadPortMin, kFMsp_ReadPortMax, flags, Do_FMsp_ReadPort) == 0)
		{
			result = kCurrentExtnVersion;
		}
	}

	if (version >= k150ExtnVersion)
	{
		name->Assign(kFMsp_PortListName, fmx::Text::kEncoding_UTF8);
		definition->Assign(kFMsp_PortListDefinition, fmx::Text::kEncoding_UTF8);
		description->Assign(kFMsp_PortListDescription, fmx::Text::kEncoding_UTF8);
		if (fmx::ExprEnv::RegisterExternalFunctionEx(*pluginID, kFMsp_PortListID, *name, *definition, *description, kFMsp_PortListMin, kFMsp_PortListMax, flags, Do_FMsp_PortList) == 0)
		{
			result = kCurrentExtnVersion;
		}
	}
	else if (version == k140ExtnVersion)
	{
		name->Assign(kFMsp_PortListName, fmx::Text::kEncoding_UTF8);
		definition->Assign(kFMsp_PortListDefinition, fmx::Text::kEncoding_UTF8);
		if (fmx::ExprEnv::RegisterExternalFunction(*pluginID, kFMsp_PortListID, *name, *definition, kFMsp_PortListMin, kFMsp_PortListMax, flags, Do_FMsp_PortList) == 0)
		{
			result = kCurrentExtnVersion;
		}
	}

	if (version >= k160ExtnVersion)
	{
		name->Assign(kFMsp_ReadPortScriptStepName, fmx::Text::kEncoding_UTF8);
		definition->Assign(kFMsp_ReadPortScriptStepDefinition, fmx::Text::kEncoding_UTF8);
		description->Assign(kFMsp_ReadPortScriptStepDescription, fmx::Text::kEncoding_UTF8);
		if (fmx::ExprEnv::RegisterScriptStep(*pluginID, kFMsp_ReadPortScriptStepID, *name, *definition, *description, flags, Do_FMsp_ReadPort) == 0)
		{
			result = kCurrentExtnVersion;
		}
	}

	if (version >= k160ExtnVersion)
	{
		name->Assign(kFMsp_PortListScriptStepName, fmx::Text::kEncoding_UTF8);
		definition->Assign(kFMsp_PortListScriptStepDefinition, fmx::Text::kEncoding_UTF8);
		description->Assign(kFMsp_PortListScriptStepDescription, fmx::Text::kEncoding_UTF8);
		if (fmx::ExprEnv::RegisterScriptStep(*pluginID, kFMsp_PortListScriptStepID, *name, *definition, *description, flags, Do_FMsp_ReadPort) == 0)
		{
			result = kCurrentExtnVersion;
		}
	}
	return result;
}

// Do_PluginShutdown ===========================================================================

static void Do_PluginShutdown( fmx::int16 version)
{
	const fmx::QuadCharUniquePtr    pluginID(kFMsp[0], kFMsp[1], kFMsp[2], kFMsp[3]);

	if (version >= k140ExtnVersion)
	{
		static_cast<void>(fmx::ExprEnv::UnRegisterExternalFunction(*pluginID, kFMsp_ReadPortID));
	}

	if (version >= k140ExtnVersion)
	{
		static_cast<void>(fmx::ExprEnv::UnRegisterExternalFunction(*pluginID, kFMsp_PortListID));
	}

	if (version >= k160ExtnVersion)
	{
		static_cast<void>(fmx::ExprEnv::UnRegisterScriptStep(*pluginID, kFMsp_ReadPortScriptStepID));
	}

	if (version >= k160ExtnVersion)
	{
		static_cast<void>(fmx::ExprEnv::UnRegisterScriptStep(*pluginID, kFMsp_PortListScriptStepID));
	}
}

// Do_GetString ================================================================================

static void CopyUTF8StrToUnichar16Str(const char* inStr, fmx::uint32 outStrSize, fmx::unichar16* outStr)
{
	fmx::TextUniquePtr txt;
	txt->Assign(inStr, fmx::Text::kEncoding_UTF8);
	const fmx::uint32 txtSize((outStrSize <= txt->GetSize()) ? (outStrSize - 1) : txt->GetSize());
	txt->GetUnicode(outStr, 0, txtSize);
	outStr[txtSize] = 0;
}

static void Do_GetString(fmx::uint32 whichString, fmx::uint32, fmx::uint32 outBufferSize, fmx::unichar16* outBuffer)
{
	switch (whichString)
	{
	case kFMXT_NameStr:
	{
		CopyUTF8StrToUnichar16Str("SerialPort", outBufferSize, outBuffer);
		break;
	}

	case kFMXT_AppConfigStr:
	{
		CopyUTF8StrToUnichar16Str("My File-Maker plugin to read from a COM port", outBufferSize, outBuffer);
		break;
	}

	case kFMXT_OptionsStr:
	{
		// Characters 1-4 are the plug-in ID
		CopyUTF8StrToUnichar16Str(kFMsp, outBufferSize, outBuffer);

		// Character 5 is always "1"
		outBuffer[4] = '1';

		// Character 6
		// use "Y" if you want to enable the Configure button for plug-ins in the Preferences dialog box.
		// Use "n" if there is no plug-in configuration needed.
		// If the flag is set to "Y" then make sure to handle the kFMXT_DoAppPreferences message.
		outBuffer[5] = 'n';

		// Character 7  is always "n"
		outBuffer[6] = 'n';

		// Character 8
		// Set to "Y" if you want to receive kFMXT_Init/kFMXT_Shutdown messages
		// In most cases, you want to set it to 'Y' since it's the best time to register/unregister your plugin functions.
		outBuffer[7] = 'Y';

		// Character 9
		// Set to "Y" if the kFMXT_Idle message is required.
		// For simple external functions this may not be needed and can be turned off by setting the character to "n"
		outBuffer[8] = 'n';

		// Character 10
		// Set to "Y" to receive kFMXT_SessionShutdown and kFMXT_FileShutdown messages
		outBuffer[9] = 'n';

		// Character 11 is always "n"
		outBuffer[10] = 'n';

		// NULL terminator
		outBuffer[11] = 0;
		break;
	}

	case kFMXT_HelpURLStr:
	{
		CopyUTF8StrToUnichar16Str("http://httpbin.org/get?id=", outBufferSize, outBuffer);
		break;
	}

	default:
	{
		outBuffer[0] = 0;
		break;
	}
	}
}

// Do_PluginIdle ===========================================================================

static void Do_PluginIdle( FMX_IdleLevel idleLevel, fmx::ptrtype )
{
	// Check idle state
	switch (idleLevel)
	{
	case kFMXT_UserIdle:
	case kFMXT_UserNotIdle:
	case kFMXT_ScriptPaused:
	case kFMXT_ScriptRunning:
	case kFMXT_Unsafe:
	default:
	{
		break;
	}
	}
}

// Do_PluginPrefs ==========================================================================

static void Do_PluginPrefs(void)
{
}

// Do_SessionNotifications =================================================================

static void Do_SessionNotifications( fmx::ptrtype )
{
}

// Do_FileNotifications ====================================================================

static void Do_FilenNotifications( fmx::ptrtype, fmx::ptrtype )
{
}

// FMExternCallProc ========================================================================

FMX_ExternCallPtr gFMX_ExternCallPtr(nullptr);

void FMX_ENTRYPT FMExternCallProc(FMX_ExternCallPtr pb)
{
	// Setup global defined in FMXExtern.h (this will be obsoleted in a later header file)
	gFMX_ExternCallPtr = pb;

	// Message dispatcher
	switch (pb->whichCall)
	{
	case kFMXT_Init:
		pb->result = Do_PluginInit(pb->extnVersion);
		break;

	case kFMXT_Idle:
		Do_PluginIdle(pb->parm1, pb->parm2);
		break;

	case kFMXT_Shutdown:
		Do_PluginShutdown(pb->extnVersion);
		break;

	case kFMXT_DoAppPreferences:
		Do_PluginPrefs();
		break;

	case kFMXT_GetString:
		Do_GetString(static_cast<fmx::uint32>(pb->parm1), static_cast<fmx::uint32>(pb->parm2), static_cast<fmx::uint32>(pb->parm3), reinterpret_cast<fmx::unichar16*>(pb->result));
		break;

	case kFMXT_SessionShutdown:
		Do_SessionNotifications(pb->parm2);
		break;

	case kFMXT_FileShutdown:
		Do_FilenNotifications(pb->parm2, pb->parm3);
		break;

	} // switch whichCall

} // FMExternCallProc