Attribute VB_Name = "modGlobalErrorCodes"
Option Explicit

Public Enum Error
    No = 0
    False_ = 0
    Unsuccess = 0
    Yes = 1
    True_ = 1
    Success = 1
    NotConfigured = 2
    WskConf = 3
    WskConnect = 4
    WskClosingTout = 5
    WskClosingError = 6
    WskConnectingTout = 7
    WskConnectionRefused = 8
    WskConnectionReset = 9
    WskHostNotFound = 10
    WskNetworkUnreachable = 11
    WskNetworkSubsystemFailed = 12
    WskConnectingFailed = 13
    WskSendDataError = 14
    PopCommunicationError = 15
    WskCommunicationTout = 16
    PopUserNotFound = 17
    PopAPOPNotSupported = 18
    PopPasswordNotCorrect = 19
    PopAPOPisSupported = 20
    PopErrWhileQuitting = 21
End Enum
