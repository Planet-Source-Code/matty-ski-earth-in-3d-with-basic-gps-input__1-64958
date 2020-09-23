Attribute VB_Name = "modGPS"
Option Explicit

Global StatsRunning As Byte     ' 1 = Running
Global LoggerBuffer() As String ' Bit wasteful on the old memory
Global SaveRawData As Byte
