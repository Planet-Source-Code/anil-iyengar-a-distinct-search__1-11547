<div align="center">

## A Distinct Search


</div>

### Description

This a Search program that searches for a

'specific record in a database.

This is a kind of DISTINCT search.

Where u just have to enter the first letter of the data u want and it gives u an output in the grid.
 
### More Info
 
'Im using a MSHFlexGrid with a ADO DataControl.

'Connect the ADODC to the Biblio.mdb

'Set Recordsource to Publisers .

im using the Mid function.

Get ADODC and connect to Biblio.mdb.

A Specific Row of data.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Anil Iyengar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/anil-iyengar.md)
**Level**          |Beginner
**User Rating**    |3.8 (73 globes from 19 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/anil-iyengar-a-distinct-search__1-11547/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
Unload Me 'Unload the program
End
End Sub
Private Sub Form_Load()
'set the Colwidth of the grid
fg.ColWidth(0) = 550
fg.ColWidth(1) = 3000
fg.ColWidth(2) = 3000
End Sub
Private Sub Text1_Change()
Adodc1.RecordSource = "select PubID,Name,[Company Name] from publishers where ucase(mid(pubid,1," & Len(Text1.Text) & "))= '" & Text1.Text & "' and ucase(mid(name,1," & Len(Text2.Text) & "))= '" & Text2.Text & "'"
Adodc1.Refresh
fg.SelectionMode = flexSelectionByRow
'The mid function checkes the records according
'to the info typed in the textbox.
'It queries the ADODC with every letter typed
'in the textbox,making it a bit more refined
'search on the records.
End Sub
Private Sub Text2_Change()
Adodc1.RecordSource = "select PubID,Name,[Company Name] from publishers where ucase(mid(name,1," & Len(Text2.Text) & "))= '" & Text2.Text & "'"
Adodc1.Refresh
fg.SelectionMode = flexSelectionByRow
End Sub
'Just use the mid function as i have and
'you can query any database for the record.
'This kind of search is useful if the u have to
'go thru a large database.
'PLEASE VOTE
```

