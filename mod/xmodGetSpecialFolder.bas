Attribute VB_Name = "xmodGetSpecialFolder"
Option Explicit

'Author is Microsoft Source is MSDN

Public Enum CSIDLS
  CSIDL_DESKTOP = &H0&  ' Desktop
  CSIDL_INTERNET = &H1& ' Internet
  CSIDL_PROGRAMS = &H2& ' Startmenü: Programme
  CSIDL_CONTROLS = &H3& ' Systemsteuerung
  CSIDL_PRINTERS = &H4& ' Drucker
  CSIDL_PERSONAL = &H5& ' Eigene Dateien
  CSIDL_FAVORITES = &H6&   ' IE: Favoriten
  CSIDL_STARTUP = &H7&  ' Autostart
  CSIDL_RECENT = &H8&   ' Zuletzt benutzte Dokumente
  CSIDL_SENDTO = &H9&   ' Senden an / SendTo
  CSIDL_BITBUCKET = &HA&   ' Papierkorb
  CSIDL_STARTMENU = &HB&   ' Startmenü
  CSIDL_MYMUSIC = &HD   ' Eigene Musik
  CSIDL_MYVIDEO = &HE   ' Eigene Videos
  CSIDL_DESKTOPDIRECTORY = &H10& ' Desktopverzeichnis
  CSIDL_DRIVES = &H11&  ' Mein Computer
  CSIDL_NETWORK = &H12& ' Netzwerk
  CSIDL_NETHOOD = &H13& ' Netzwerkumgebung
  CSIDL_FONTS = &H14&   ' Windows\Fonts
  CSIDL_TEMPLATES = &H15&  ' Vorlagen
  CSIDL_COMMON_STARTMENU = &H16&  ' "All Users" - Startmenü
  CSIDL_COMMON_PROGRAMS = &H17&   ' "All Users" - Programme
  CSIDL_COMMON_STARTUP = &H18& ' "All Users" - Autostart
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19& ' "All Users" - Desktop
  CSIDL_APPDATA = &H1A&  ' Anwendungsdaten
  CSIDL_PRINTHOOD = &H1B&   ' Druckumgebung
  CSIDL_LOCAL_APPDATA = &H1C&  ' Lokale Einstellungen\Anwendungsdaten
  CSIDL_COMMON_FAVORITES = &H1F&  ' "All Users" - Favoriten
  CSIDL_INTERNET_CACHE = &H20& ' IE: Temporäre Internetdateien
  CSIDL_COOKIES = &H21&  ' IE: Cookies
  CSIDL_HISTORY = &H22&  ' IE: Verlauf
  CSIDL_COMMON_APPDATA = &H23& ' "All Users" - Anwendungsdaten
  CSIDL_WINDOWS = &H24&  ' Windows
  CSIDL_SYSTEM = &H25&   ' Windows\System32
  CSIDL_PROGRAM_FILES = &H26&  ' C:\Programme
  CSIDL_MYPICTURES = &H27&  ' Eigene Bilder
  CSIDL_PROFILE = &H28&  ' Anwenderprofil (Benutzername)
  CSIDL_SYSTEMX86 = &H29&   ' Windows\System32
  CSIDL_PROGRAM_FILES_COMMON = &H2B& ' Gemeinsame Dateien
  CSIDL_COMMON_TEMPLATES = &H2D&  ' "All Users" - Vorlagen
  CSIDL_COMMON_DOCUMENTS = &H2E&  ' "All Users" - Dokumente
  CSIDL_COMMON_ADMINTOOLS = &H2F& ' "All Users" - Verwaltung
  CSIDL_ADMINTOOLS = &H30&  ' Startmenü\Programme\Verwaltung
End Enum
Private Const CSIDL_FLAG_DONT_VERIFY As Long = &H4000&
Private Const CSIDL_FLAG_CREATE   As Long = &H8000&
Private Const CSIDL_FLAG_MASK  As Long = &HFF00&
Private Const SHGFP_TYPE_CURRENT As Long = 0&
Private Const SHGFP_TYPE_DEFAULT As Long = 1&
Private Const MAX_PATH As Long = 260&
Private Const S_OK   As Long = 0&
Private Const S_FALSE   As Long = 1&
Private Const E_INVALIDARG As Long = &H80070057
Private Declare Function SHGetFolderPath _
 Lib "shfolder" Alias "SHGetFolderPathA" ( _
 ByVal hWndOwner As Long, _
 ByVal Folder As Long, _
 ByVal hToken As Long, _
 ByVal Flags As Long, _
 ByVal strPath As String _
 ) As Long
Public Function GetSpecialFolder(ByVal CSIDL As CSIDLS, _
   Optional ByVal Create As Boolean = False, _
   Optional ByVal Verify As Boolean = False _
   ) As String
' Liefert den Pfad zu einem speziellen Verzeichnis zurück. Im Fehlerfall
' wird ein leerer String returniert. Wird Create zu True gesetzt, so wird
' ein abgefragtes Verzeichnis bei Bedarf automatisch angelegt. Setzen Sie
' Verify zu True, um vor der Rückgabe eines Pfades eine Prüfung durchzu-
' führen, dass der Pfad tatsächlich existiert.
Dim sPath As String ' Zu ermittelnder Pfad
Dim RetVal As Long  ' Rückgabewert
Dim lFlags As Long  ' Eigenschaften
  ' Stringbuffer füllen
  sPath = Space$(MAX_PATH)
  ' Flags-Parameter zusammenstellen
  lFlags = CSIDL
  If Create Then ' Bei Bedarf automatisch erzeugen
 lFlags = lFlags Or CSIDL_FLAG_CREATE
  End If
  If Not Verify Then ' Existenz nicht überprüfen
 lFlags = lFlags Or CSIDL_FLAG_DONT_VERIFY
  End If
  ' Pfad zum Verzeichnis ermitteln
  RetVal = SHGetFolderPath(0, lFlags, 0, SHGFP_TYPE_CURRENT, sPath)
  ' Erfolgskontrolle und Rückgabe des Ergebnisses
  Select Case RetVal
 Case S_OK
   ' Gültiges Verzeichnis gefunden
   GetSpecialFolder = Left$(sPath, InStr(1, sPath, vbNullChar) - 1)
 Case S_FALSE ' S_FALSE
   ' lCSIDL ist gültig, aber das Verzeichnis existiert nicht
 Case E_INVALIDARG
   ' Ungültiges Verzeichnis
  End Select
End Function


