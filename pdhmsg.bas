Attribute VB_Name = "pdhmsg"
Option Explicit

    Public Const PDH_CSTATUS_VALID_DATA = &H0
    Public Const PDH_CSTATUS_NEW_DATA = &H1
    Public Const PDH_CSTATUS_NO_MACHINE = &H800007D0
    Public Const PDH_CSTATUS_NO_INSTANCE = &H800007D1
    Public Const PDH_MORE_DATA = &H800007D2
    Public Const PDH_CSTATUS_ITEM_NOT_VALIDATED = &H800007D3
    Public Const PDH_RETRY = &H800007D4
    Public Const PDH_NO_DATA = &H800007D5
    Public Const PDH_CALC_NEGATIVE_DENOMINATOR = &H800007D6
    Public Const PDH_CALC_NEGATIVE_TIMEBASE = &H800007D7
    Public Const PDH_CALC_NEGATIVE_VALUE = &H800007D8
    Public Const PDH_DIALOG_CANCELLED = &H800007D9
    Public Const PDH_END_OF_LOG_FILE = &H800007DA
    Public Const PDH_CSTATUS_NO_OBJECT = &HC0000BB8
    Public Const PDH_CSTATUS_NO_COUNTER = &HC0000BB9
    Public Const PDH_CSTATUS_INVALID_DATA = &HC0000BBA
    Public Const PDH_MEMORY_ALLOCATION_FAILURE = &HC0000BBB
    Public Const PDH_INVALID_HANDLE = &HC0000BBC
    Public Const PDH_INVALID_ARGUMENT = &HC0000BBD
    Public Const PDH_FUNCTION_NOT_FOUND = &HC0000BBE
    Public Const PDH_CSTATUS_NO_COUNTERNAME = &HC0000BBF
    Public Const PDH_CSTATUS_BAD_COUNTERNAME = &HC0000BC0
    Public Const PDH_INVALID_BUFFER = &HC0000BC1
    Public Const PDH_INSUFFICIENT_BUFFER = &HC0000BC2
    Public Const PDH_CANNOT_CONNECT_MACHINE = &HC0000BC3
    Public Const PDH_INVALID_PATH = &HC0000BC4
    Public Const PDH_INVALID_INSTANCE = &HC0000BC5
    Public Const PDH_INVALID_DATA = &HC0000BC6
    Public Const PDH_NO_DIALOG_DATA = &HC0000BC7
    Public Const PDH_CANNOT_READ_NAME_STRINGS = &HC0000BC8
    Public Const PDH_LOG_FILE_CREATE_ERROR = &HC0000BC9
    Public Const PDH_LOG_FILE_OPEN_ERROR = &HC0000BCA
    Public Const PDH_LOG_TYPE_NOT_FOUND = &HC0000BCB
    Public Const PDH_NO_MORE_DATA = &HC0000BCC
    Public Const PDH_ENTRY_NOT_IN_LOG_FILE = &HC0000BCD
    Public Const PDH_DATA_SOURCE_IS_LOG_FILE = &HC0000BCE
    Public Const PDH_DATA_SOURCE_IS_REAL_TIME = &HC0000BCF
    Public Const PDH_UNABLE_READ_LOG_HEADER = &HC0000BD0
    Public Const PDH_FILE_NOT_FOUND = &HC0000BD1
    Public Const PDH_FILE_ALREADY_EXISTS = &HC0000BD2
    Public Const PDH_NOT_IMPLEMENTED = &HC0000BD3
    Public Const PDH_STRING_NOT_FOUND = &HC0000BD4
    Public Const PDH_UNABLE_MAP_NAME_FILES = &H80000BD5
    Public Const PDH_UNKNOWN_LOG_FORMAT = &HC0000BD6
    Public Const PDH_UNKNOWN_LOGSVC_COMMAND = &HC0000BD7
    Public Const PDH_LOGSVC_QUERY_NOT_FOUND = &HC0000BD8
    Public Const PDH_LOGSVC_NOT_OPENED = &HC0000BD9
    Public Const PDH_WBEM_ERROR = &HC0000BDA
    Public Const PDH_ACCESS_DENIED = &HC0000BDB
    Public Const PDH_LOG_FILE_TOO_SMALL = &HC0000BDC


Public Function PdhError(lngError As Long, apiFunction As String, Optional errDescription As String, Optional errConst As String)
    Select Case lngError
        Case PDH_CSTATUS_VALID_DATA: errDescription = "The returned data is valid.": errConst = "PDH_CSTATUS_VALID_DATA"
        Case PDH_CSTATUS_NEW_DATA: errDescription = "The return data value is valid and different from the last sample.": errConst = "PDH_CSTATUS_NEW_DATA"
        Case PDH_CSTATUS_NO_MACHINE: errDescription = "Unable to connect to specified machine or machine is off line.": errConst = "PDH_CSTATUS_NO_MACHINE"
        Case PDH_CSTATUS_NO_INSTANCE: errDescription = "The specified instance is not present.": errConst = "PDH_CSTATUS_NO_INSTANCE"
        Case PDH_MORE_DATA: errDescription = "There is more data to return than would fit in the supplied buffer. Allocate a larger buffer and call the function again.": errConst = "PDH_MORE_DATA"
        Case PDH_CSTATUS_ITEM_NOT_VALIDATED: errDescription = "The data item has been added to the query, but has not been validated nor accessed. No other status information on this data item is available.": errConst = "PDH_CSTATUS_ITEM_NOT_VALIDATED"
        Case PDH_RETRY: errDescription = "The selected operation should be retried.": errConst = "PDH_RETRY"
        Case PDH_NO_DATA: errDescription = "No data to return.": errConst = "PDH_NO_DATA"
        Case PDH_CALC_NEGATIVE_DENOMINATOR: errDescription = "A counter with a negative denominator value was detected.": errConst = "PDH_CALC_NEGATIVE_DENOMINATOR"
        Case PDH_CALC_NEGATIVE_TIMEBASE: errDescription = "A counter with a negative timebase value was detected.": errConst = "PDH_CALC_NEGATIVE_TIMEBASE"
        Case PDH_CALC_NEGATIVE_VALUE: errDescription = "A counter with a negative value was detected.": errConst = "PDH_CALC_NEGATIVE_VALUE"
        Case PDH_DIALOG_CANCELLED: errDescription = "The user cancelled the dialog box.": errConst = "PDH_DIALOG_CANCELLED"
        Case PDH_END_OF_LOG_FILE: errDescription = "The end of the log file was reached.": errConst = "PDH_END_OF_LOG_FILE"
        Case PDH_CSTATUS_NO_OBJECT: errDescription = "The specified object is not found on the system.": errConst = "PDH_CSTATUS_NO_OBJECT"
        Case PDH_CSTATUS_NO_COUNTER: errDescription = "The specified counter could not be found.": errConst = "PDH_CSTATUS_NO_COUNTER"
        Case PDH_CSTATUS_INVALID_DATA: errDescription = "The returned data is not valid.": errConst = "PDH_CSTATUS_INVALID_DATA"
        Case PDH_MEMORY_ALLOCATION_FAILURE: errDescription = "A PDH function could not allocate enough temporary memory to complete the operation. Close some applications or extend the pagefile and retry the function.": errConst = "PDH_MEMORY_ALLOCATION_FAILURE"
        Case PDH_INVALID_HANDLE: errDescription = "The handle is not a valid PDH object.": errConst = "PDH_INVALID_HANDLE"
        Case PDH_INVALID_ARGUMENT: errDescription = "A required argument is missing or incorrect.": errConst = "PDH_INVALID_ARGUMENT"
        Case PDH_FUNCTION_NOT_FOUND: errDescription = "Unable to find the specified function.": errConst = "PDH_FUNCTION_NOT_FOUND"
        Case PDH_CSTATUS_NO_COUNTERNAME: errDescription = "No counter was specified.": errConst = "PDH_CSTATUS_NO_COUNTERNAME"
        Case PDH_CSTATUS_BAD_COUNTERNAME: errDescription = "Unable to parse the counter path. Check the format and syntax of the specified path.": errConst = "PDH_CSTATUS_BAD_COUNTERNAME"
        Case PDH_INVALID_BUFFER: errDescription = "The buffer passed by the caller is invalid.": errConst = "PDH_INVALID_BUFFER"
        Case PDH_INSUFFICIENT_BUFFER: errDescription = "The requested data is larger than the buffer supplied. Unable to return the requested data.": errConst = "PDH_INSUFFICIENT_BUFFER"
        Case PDH_CANNOT_CONNECT_MACHINE: errDescription = "Unable to connect to the requested machine.": errConst = "PDH_CANNOT_CONNECT_MACHINE"
        Case PDH_INVALID_PATH: errDescription = "The specified counter path could not be interpreted.": errConst = "PDH_INVALID_PATH"
        Case PDH_INVALID_INSTANCE: errDescription = "The instance name could not be read from the specified counter path.": errConst = "PDH_INVALID_INSTANCE"
        Case PDH_INVALID_DATA: errDescription = "The data is not valid.": errConst = "PDH_INVALID_DATA"
        Case PDH_NO_DIALOG_DATA: errDescription = "The dialog box data block was missing or invalid.": errConst = "PDH_NO_DIALOG_DATA"
        Case PDH_CANNOT_READ_NAME_STRINGS: errDescription = "Unable to read the counter and/or explain text from the specified machine.": errConst = "PDH_CANNOT_READ_NAME_STRINGS"
        Case PDH_LOG_FILE_CREATE_ERROR: errDescription = "Unable to create the specified log file.": errConst = "PDH_LOG_FILE_CREATE_ERROR"
        Case PDH_LOG_FILE_OPEN_ERROR: errDescription = "Unable to open the specified log file.": errConst = "PDH_LOG_FILE_OPEN_ERROR"
        Case PDH_LOG_TYPE_NOT_FOUND: errDescription = "The specified log file type has not been installed on this system.": errConst = "PDH_LOG_TYPE_NOT_FOUND"
        Case PDH_NO_MORE_DATA: errDescription = "No more data is available.": errConst = "PDH_NO_MORE_DATA"
        Case PDH_ENTRY_NOT_IN_LOG_FILE: errDescription = "The specified record was not found in the log file.": errConst = "PDH_ENTRY_NOT_IN_LOG_FILE"
        Case PDH_DATA_SOURCE_IS_LOG_FILE: errDescription = "The specified data source is a log file.": errConst = "PDH_DATA_SOURCE_IS_LOG_FILE"
        Case PDH_DATA_SOURCE_IS_REAL_TIME: errDescription = "The specified data source is the current activity.": errConst = "PDH_DATA_SOURCE_IS_REAL_TIME"
        Case PDH_UNABLE_READ_LOG_HEADER: errDescription = "The log file header could not be read.": errConst = "PDH_UNABLE_READ_LOG_HEADER"
        Case PDH_FILE_NOT_FOUND: errDescription = "Unable to find the specified file.": errConst = "PDH_FILE_NOT_FOUND"
        Case PDH_FILE_ALREADY_EXISTS: errDescription = "There is already a file with the specified file name.": errConst = "PDH_FILE_ALREADY_EXISTS"
        Case PDH_NOT_IMPLEMENTED: errDescription = "The function referenced has not been implemented.": errConst = "PDH_NOT_IMPLEMENTED"
        Case PDH_STRING_NOT_FOUND: errDescription = "Unable to find the specified string in the list of performance name and explain text strings.": errConst = "PDH_STRING_NOT_FOUND"
        Case PDH_UNABLE_MAP_NAME_FILES: errDescription = "Unable to map to the performance counter name data files. The data will be read from the registry and stored locally.": errConst = "PDH_UNABLE_MAP_NAME_FILES"
        Case PDH_UNKNOWN_LOG_FORMAT: errDescription = "The format of the specified log file is not recognized by the PDH DLL.": errConst = "PDH_UNKNOWN_LOG_FORMAT"
        Case PDH_UNKNOWN_LOGSVC_COMMAND: errDescription = "The specified Log Service command value is not recognized.": errConst = "PDH_UNKNOWN_LOGSVC_COMMAND"
        Case PDH_LOGSVC_QUERY_NOT_FOUND: errDescription = "The specified Query from the Log Service could not be found or could not be opened.": errConst = "PDH_LOGSVC_QUERY_NOT_FOUND"
        Case PDH_LOGSVC_NOT_OPENED: errDescription = "The Performance Data Log Service key could not be opened. This may be due to insufficient privilege or because the service has not been installed.": errConst = "PDH_LOGSVC_NOT_OPENED"
        Case PDH_WBEM_ERROR: errDescription = "An error occured while accessing the WBEM data store.  The WBEM error code is contained in the LastError value.": errConst = "PDH_WBEM_ERROR"
        Case PDH_ACCESS_DENIED: errDescription = "Unable to access the desired machine or service. Check the permissions and authentication of the log service or the interactive user session against those on the machine or service being monitored.": errConst = "PDH_ACCESS_DENIED"
        Case PDH_LOG_FILE_TOO_SMALL: errDescription = "The maximum log file size specified is too small to log the selected counters. No data will be recorded in this log file. Specify a smaller set of counters to log or a larger file size and retry this call.": errConst = "PDH_LOG_FILE_TOO_SMALL"
    End Select
    
    If errMsg = True Then
        MsgBox apiFunction & vbCrLf & vbCrLf & errConst & vbCrLf & vbCrLf & errDescription, vbExclamation, "Error"
    End If
End Function
