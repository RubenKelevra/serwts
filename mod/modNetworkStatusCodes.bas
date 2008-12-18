Attribute VB_Name = "modNetworkStatusCodes"
Option Explicit

Public Enum POP3Stat
    awaitingFirstOK = 0
    awaitingUserOK = 1
    awaitingPassOK = 2
    awaitingStatOK = 3
    awaitingListOK = 4
    awaitingRetrOK = 5
    awaitingDeleOK = 6
    awaitingNoopOK = 7
    awaitingRsetOK = 8
    awaitingQuitOK = 9
    awaitingTopOK = 10
    awaitingUidlOK = 11
    awaitingApopOK = 12
    closed = 13
End Enum

Public Enum POP3TaskCode
    NoTask = 0
    checkAPOPCapability = 1
    getEmailHeaders = 2 'fetch all emailheaders
    
    'email is definied in global var
    getOneEmail_Wuid_Wdelete = 3
    getOneEmail_Wuid_WOdelete = 4
    deleteEmail_Wuid = 5
    getAllEmails_Wuid_deleteALL = 6 'get missing mails decided via uid and delete fetched and not fetched mails
    getAllEmails_Wuid_NOdelete = 7
    getAllEmails_Wuid_deleteFETCHED = 8
    getNewEmailHeaders_Wuid = 9
    
    getOneEmail_WOuid_Wdelete = 10
    getOneEmail_WOuid_WOdelete = 11
    deleteEmail_WOuid = 12
    getAllEmails_WOuid_deleteALL = 13 'get missing mails and delete fetched and not fetched mails
    getAllEmails_WOuid_NOdelete = 14
    getAllEMails_WOuid_deleteFETCHED = 15
    getNewEmailHeaders_WOuid = 16
End Enum
    
    
    
