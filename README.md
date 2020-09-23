<div align="center">

## Copy a database table


</div>

### Description

How to copy a database table. This may require some tweaking....

"Bill Pearson" <billp@dnai.com>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |3.8 (23 globes from 6 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-copy-a-database-table__1-540/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
  Dim dbFrom As Database
  Dim dbTo  As Database
  Set dbFrom = workspaces(0).opendatabase("c:\vb4\biblio.mdb")
  Set dbTo = workspaces(0).opendatabase("c:\vb4\biblio.mdb")
  db_Copy_Tabledef dbFrom, dbTo, "Authors", "CopyOfAuthors"
  dbFrom.Close
  dbTo.Close
End Sub
Public Function db_Copy_Tabledef(dbFrom As Database, dbTo As Database,
TableNameFrom As String, TableNameTo As String) As Boolean
  Dim tdFrom    As TableDef
  Dim tdTo     As TableDef
  Dim fldFrom   As Field
  Dim fldTo    As Field
  Dim ndxFrom   As Index
  Dim ndxTo    As Index
  Dim FunctionName As String
  Dim Found    As Boolean
  On Error Resume Next
  For Each tdFrom In dbFrom.TableDefs
    '-----------------------------
    'Loop until find the table def
    '-----------------------------
    If LCase$(tdFrom.Name) = LCase$(TableNameFrom) Then
      Found = True
     '----------------------
     'Create Table defintion
     '----------------------
      Set tdTo = dbTo.CreateTableDef(TableNameTo)
     '------------------------------
     'Copy each field and attributes
     '------------------------------
      For Each fldFrom In dbFrom.TableDefs(tdFrom.Name).Fields
        Set fldTo = tdTo.CreateField(fldFrom.Name)
        fldTo.Type = fldFrom.Type
        fldTo.DefaultValue = fldFrom.DefaultValue
        fldTo.Required = fldFrom.Required
        Select Case fldFrom.Type
         Case dbText
           fldTo.Size = fldFrom.Size
           fldTo.Attributes = fldFrom.Attributes
           fldTo.AllowZeroLength = fldTo.AllowZeroLength
         Case dbMemo
           fldTo.AllowZeroLength = fldTo.AllowZeroLength
         Case Else
        End Select
        tdTo.Fields.Append fldTo
        If Err.Number > 0 Then
         MsgBox "Error adding field to table " & TableNameTo &
".", vbCritical, FunctionName
         Exit Function
        End If
      Next
     '-----------------------
     'Copy Index defintion(s)
     '-----------------------
      For Each ndxFrom In dbFrom.TableDefs(tdFrom.Name).Indexes
        Set ndxTo = tdTo.CreateIndex(ndxFrom.Name)
        ndxTo.Required = ndxFrom.Required
        ndxTo.IgnoreNulls = ndxFrom.IgnoreNulls
        ndxTo.Primary = ndxFrom.Primary
        ndxTo.Clustered = ndxFrom.Clustered
        ndxTo.Unique = ndxFrom.Unique
       '---------------------
       'Copy each index field
       '---------------------
        For Each fldFrom In
dbFrom.TableDefs(tdFrom.Name).Indexes(ndxFrom.Name).Fields
          Set fldTo = ndxTo.CreateField(fldFrom.Name)
          ndxTo.Fields.Append fldTo
          If Err.Number > 0 Then
           MsgBox "Error adding field to index in table " &
TableNameTo & ".", vbCritical, FunctionName
           Exit Function
          End If
        Next
        tdTo.Indexes.Append ndxTo
        If Err.Number > 0 Then
         MsgBox "Error adding index to table " & TableNameTo &
".", vbCritical, FunctionName
         Exit Function
        End If
      Next
      dbTo.TableDefs.Append tdTo
      If Err.Number > 0 Then
       MsgBox "Error adding table " & TableNameTo & ".", vbCritical,
FunctionName
       Exit Function
      End If
      Exit For
    End If
  Next
  If Found Then
    db_Copy_Tabledef = True
  Else
    MsgBox "Table " & TableNameFrom & " not found.", vbExclamation,
FunctionName
  End If
  On Error GoTo 0
End Function
```

