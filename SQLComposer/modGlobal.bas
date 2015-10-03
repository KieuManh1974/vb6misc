Attribute VB_Name = "modGlobal"
Option Explicit

Public goCon As Connection
Public goConEBS As Connection
Public goParseExpression As IParseObject
Public goParseTable As IParseObject
Public goParseFunction As IParseObject
Public goParseIdentifier As IParseObject

Public goDatabases As New clsDatabases
Public goTables As New clsTables
Public goFields As New clsFields
Public goSQL As New clsSQL
Public goExpressions As New clsExpressions
