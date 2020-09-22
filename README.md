<div align="center">

## a \- Navigate And Wait a Web Browser Control


</div>

### Description

Navigate a web browser control and wait for it to finish loading before going on to the next lines of code!
 
### More Info
 
You simply need to have a Web Browser control on your form and add this sample code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brogan Scott Houston McIntyre](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brogan-scott-houston-mcintyre.md)
**Level**          |Beginner
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brogan-scott-houston-mcintyre-a-navigate-and-wait-a-web-browser-control__1-45358/archive/master.zip)





### Source Code

```
'This is the NavigateAndWait sub.
'rURL is the hyperlink that you want the web browser control to navigate to.
'rWebBrowser is the web browser control that you would like to use.
Public Sub NavigateAndWait(rURL As String, rWebBrowser As Control)
rWebBrowser.Navigate rURL 'Navigate to any URL.
Do Until rWebBrowser.ReadyState = READYSTATE_COMPLETE 'Loop until the web browser control is ready.
  DoEvents 'This will stop the program from locking up while waiting, and allow other code to continue processing.
Loop
End Sub
Private Sub Command1_Click() 'This is an example button.
NavigateAndWait "http://www.Planet-Source-Code.com", Me.WebBrowser1 'This is an example on how to call the NavigateAndWait sub.
MsgBox "Loaded!" 'This is a message box to alert you that the page has been loaded. This line is not required.
End Sub
```

