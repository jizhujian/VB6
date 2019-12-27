#if defined CPTDLL
#  define DllExport __declspec(dllexport)
#else
#  define DllExport __declspec(dllimport)
#endif

// communication types - RS232 and IRDA are similar
#define CPT_RS232     1
#define CPT_CRADLE_IR 2
#define CPT_IRDA      3

// return statuses
#define CPT_OK                  0 // operation successfull
#define CPT_COM_PORT_NOT_OPEN  -1 // CPTOpenPort com port can not be open
#define CPT_CRADLE_IR_NOT_SET  -2 // CPTOpenPort CradleIR can not be configured
#define CPT_COM_PORT_NOT_INI   -3 // CPTOpenPort com port can not be initialized
#define CPT_TIMEOUT            -4 // operation timeout
#define CPT_NO_MORE_DATA       -5 // CODEWARE_CPTReadRecord all data are received
#define CPT_INVALID_RECORD     -6 // CODEWARE_CPTReadRecord received record is invalid and
                                  // should be received once more

DllExport int WINAPI CODEWARE_CPTOpenComPort(int iPortNumber,
                                             int iBaudRate,
                                             int iPortType);

/****************************************************************************
 * Purpose     To open the COM port.
 * Parameters  iPortNumber : COM port, can be 1..n
 *             iBaudRate   : Transmission rate, can be 9600/19200/38400/57600/115200
 *             iPortType   : Port type, CPT_RS232, CPT_IRDA, CPT_CRADLE_IR
 * Returns     If successful, it returns the CPT_OK status.
 *             If fails to open the port, the return value see return statuses.
 * Example     CODEWARE_CPTOpenComPort(1, 115200, CPT_CRADLE_IR);
 ***************************************************************************/

DllExport void WINAPI CODEWARE_CPTCloseComPort(void);
/****************************************************************************
 * Purpose     To close the COM port.
 * Parameters  None
 * Returns     None
 * Example     CODEWARE_CPTCloseComPort();
 ***************************************************************************/

DllExport void WINAPI CODEWARE_CPTWriteComPort(LPSTR szWriteString);
/****************************************************************************
 * Purpose     Sends a string to the COM port.
 * Parameters  szWriteString: A pointer to a null-terminated string.
 * Returns     None
 * Example     CODEWARE_CPTWriteComPort("READ\r");
 ***************************************************************************/

DllExport void WINAPI CODEWARE_CPTWriteComPortBin(BYTE * btWriteData,
                                                  DWORD  dwDataLen);
/****************************************************************************
 * Purpose     Sends a data stream to the COM port.
 * Parameters  btWriteData: A pointer to a data stream.
 *             dwDataLen: Data stream length.
 * Returns     None
 * Example     CODEWARE_CPTWriteComPortBin((BYTE *)"READ\r", 5);
 ***************************************************************************/

DllExport int WINAPI CODEWARE_CPTReadComPort(LPSTR szReadString);
/****************************************************************************
 * Purpose     To read data from the COM port.
 * Parameters  szReadString: A string buffer for receiving the data.
 * Returns     Number of the received characters.
 *             CPT_TIMEOUT if transmission error or time out (5 sec).
 * Example     nChar = CODEWARE_CPTReadComPort(szBuf)
 ***************************************************************************/

DllExport int WINAPI CODEWARE_CPTShowErrorMessage (int nShow);
/****************************************************************************
 * Purpose     Enable or disable the error messages in case error happens
 * Parameters  An integer value, 1 for enable, 0 for disable
 * Returns     Always return CPT_OK
 * Example     CODEWARE_CPTShowErrorMessage(0);   // disable showing error messages
 ***************************************************************************/

DllExport HANDLE WINAPI CODEWARE_CPTGetComPortHandle(void);
/****************************************************************************
 * Purpose     Returns handle (for C) of the opened COM port.
 * Parameters  None
 * Returns     Handle (for C) of the opened COM port
 * Example     hComPort = CODEWARE_CPTGetComPortHandle();
 ***************************************************************************/

DllExport int WINAPI CODEWARE_CPTStartDataUpload(int iPortNumber,
                                                 int iBaudRate,
                                                 int iPortType);
/****************************************************************************
 * Purpose     To open the COM port and start data upload.
 * Parameters  iPortNumber : COM port, can be 1..n
 *             iBaudRate   : Transmission rate, can be 9600/19200/38400/57600/115200
 *             iPortType   : Port type, CPT_RS232, CPT_IRDA, CPT_CRADLE_IR
 * Returns     If successful, it returns the CPT_OK status.
 *             CPT_TIMEOUT if transmission error or time out (5 sec).
 *             If fails to open the port, the return value see return statuses.
 * Example     CODEWARE_CPTStartDataUpload(1, 115200, CPT_CRADLE_IR);
 ***************************************************************************/

DllExport int WINAPI CODEWARE_CPTReadDataRecord(LPSTR szReadRecord);
/****************************************************************************
 * Purpose     To read data from the terminal. Data are transferred in internal
 *             format and this function checks data integrity and translate it
 *             to the pure data record.
 * Parameters  szReadRecord: A string buffer for receiving the data record.
 * Returns     Number of the received characters.
 *             CPT_TIMEOUT if transmission error or time out (5 sec).
 *             CPT_NO_MORE_DATA  all data are received
 *             CPT_INVALID_RECORD  received record is invalid and should be
 *                                 received once more
 * Example     nChar = CODEWARE_CPTReadDataRecord(szRecord)
 ***************************************************************************/

DllExport void WINAPI CODEWARE_CPTFinishDataUpload(void);
/****************************************************************************
 * Purpose     To finish data upload and close the COM port.
 * Parameters  None
 * Returns     None
 * Example     CODEWARE_CPTFinishDataUpload();
 ***************************************************************************/

DllExport int WINAPI CODEWARE_CPTStartLookupDownload(int iPortNumber,
                                                     int iBaudRate,
                                                     int iPortType);
/****************************************************************************
 * Purpose     To open the COM port and start data upload.
 * Parameters  iPortNumber : COM port, can be 1..n
 *             iBaudRate   : Transmission rate, can be 9600/19200/38400/57600/115200
 *             iPortType   : Port type, CPT_RS232, CPT_IRDA, CPT_CRADLE_IR
 * Returns     If successful, it returns the CPT_OK status.
 *             CPT_TIMEOUT if transmission error or time out (5 sec).
 *             If fails to open the port, the return value see return statuses.
 * Example     CODEWARE_CPTStartLookupDownload(1, 115200, CPT_CRADLE_IR);
 ***************************************************************************/

DllExport int WINAPI CODEWARE_CPTSendLookupRecord(LPSTR szSendRecord);
/****************************************************************************
 * Purpose     To send data to the terminal. Data are transferred in internal
 *             format and this function translate the record to the internal
 *             failsafe format.
 * Parameters  szSendRecord: Data record to be sent.
 * Returns     CPT_OK if the record is successfully sent.
 *             CPT_TIMEOUT if transmission error or time out (5 sec).
 *             CPT_INVALID_RECORD  record received by the terminal is invalid
 *                                 and should be send once more
 * Example     iStat =  CODEWARE_CPTSendLookupRecord(szRecord)
 ***************************************************************************/

DllExport void WINAPI CODEWARE_CPTFinishLookupDownload(void);
/****************************************************************************
 * Purpose     To finish lookup download and close the COM port.
 * Parameters  None
 * Returns     None
 * Example     CODEWARE_CPTFinishLookupDownload();
 ***************************************************************************/

DllExport int WINAPI CODEWARE_CPTBrowseForFile(LPSTR szResultedFile,
                                               LPSTR szStartDirectory,
                                               LPSTR szFileTypeComment,
                                               LPSTR szFileType);
/****************************************************************************
 * Purpose     To browse for a file, for example for the lookup.
 * Parameters  szResultedFile: Resulted file.
 *             szStartDirectory: Directory to start search.
 *             szFileTypeComment: Description of required files.
 *             szFileType: Required file type.
 * Returns     0 if fails otherwise length of the resulted file.
 * Example     CODEWARE_CPTBrowseForFile(szResultedFile,
                                         "c:\\lookups",
                                         "Text"
                                         "txt");
 ***************************************************************************/

/* ----------------------------------------------------------------------- */
/* Sample program #1                                                       */
/* This sample use the above functions to upload data.                     */
/*
{

  char szData[256];
  int  iReadLength;

  if(CODEWARE_CPTOpenComPort(1, 115200, CPT_CRADLE_IR))
  {

    CODEWARE_CPTWriteComPort("READ\r");
    while(1)
    {

      iReadLength = CODEWARE_CPTReadComPort(szData);
      szData[iReadLength] = '\0';

      if (lstrcmp(szData, "OVER\r") == 0)
      {
        MessageBox (hwnd, "Done", "Test", MB_OK);
        break;
      }

      if (!*szData)
      {
        MessageBox (hwnd, "Transmission Error", "Test", MB_OK);
        break;
      }

      if (lstrcmp(szData, "NAK\r") == 0)
      {
        MessageBox(hwnd, "Command Error", "Test", MB_OK);
        break;
      }

      CODEWARE_CPTWriteComPort("ACK\r");

      // to do :
      // save the received data (szData) to a file

    }

    CODEWARE_CPTCloseComPort();

  }

}
*/
/* ----------------------------------------------------------------------- */
