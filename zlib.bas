Attribute VB_Name = "zlib"
'zlib.h - 1.1.3

Option Explicit


'Functions - Not all

Public Declare Function adler32 Lib "zlib.dll" (ByVal adler As Long, ByVal buf As String, ByVal buf_len As Long) As Long
Public Declare Function crc32 Lib "zlib.dll" (ByVal crc As Long, ByVal buf As String, ByVal buf_len As Long) As Long
Public Declare Function compress Lib "zlib.dll" (ByVal dest As String, destLen As Long, ByVal source As String, ByVal sourceLen As Long) As Long
Public Declare Function compress2 Lib "zlib.dll" (ByVal dest As String, destLen As Long, ByVal source As String, ByVal sourceLen As Long, ByVal level As Integer) As Long
Public Declare Function gzclose Lib "zlib.dll" (ByVal file As Long) As Long
Public Declare Function gzflush Lib "zlib.dll" (ByVal file As Long, ByVal Flush As Long) As Long
Public Declare Function gzopen Lib "zlib.dll" (ByVal path As String, ByVal Mode As String) As Long
Public Declare Function gzread Lib "zlib.dll" (ByVal file As Long, ByVal buf As String, ByVal buf_len As Long) As Long
Public Declare Function gzwrite Lib "zlib.dll" (ByVal file As Long, ByVal buf As String, ByVal buf_len As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (ByVal dest As String, destLen As Long, ByVal source As String, ByVal sourceLen As Long) As Long
Public Declare Function zlibVersion Lib "zlib.dll" () As Long


    'Constants - All

    Public Const Z_NO_FLUSH = 0
    Public Const Z_PARTIAL_FLUSH = 1 'will be removed, use Z_SYNC_FLUSH instead
    Public Const Z_SYNC_FLUSH = 2
    Public Const Z_FULL_FLUSH = 3
    Public Const Z_FINISH = 4
    'Allowed flush values; see deflate() below for details
    
    Public Const Z_OK = 0
    Public Const Z_STREAM_END = 1
    Public Const Z_NEED_DICT = 2
    Public Const Z_ERRNO = (-1)
    Public Const Z_STREAM_ERROR = (-2)
    Public Const Z_DATA_ERROR = (-3)
    Public Const Z_MEM_ERROR = (-4)
    Public Const Z_BUF_ERROR = (-5)
    Public Const Z_VERSION_ERROR = (-6)
    'Return codes for the compression/decompression functions. Negative
    'values are errors, positive values are used for special but normal events.
    
    Public Const Z_NO_COMPRESSION = 0
    Public Const Z_BEST_SPEED = 1
    Public Const Z_BEST_COMPRESSION = 9
    Public Const Z_DEFAULT_COMPRESSION = (-1)
    'compression levels
    
    Public Const Z_FILTERED = 1
    Public Const Z_HUFFMAN_ONLY = 2
    Public Const Z_DEFAULT_STRATEGY = 0
    'compression strategy; see deflateInit2() below for details
    
    Public Const Z_BINARY = 0
    Public Const Z_ASCII = 1
    Public Const Z_UNKNOWN = 2
    'Possible values of the data_type field
    
    Public Const Z_DEFLATED = 8 'The deflate compression method (the only one supported in this version)
    
    Public Const Z_NULL = 0 'for initializing zalloc, zfree, opaque

