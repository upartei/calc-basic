sub Wikitab
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dim args1(3) as new com.sun.star.beans.PropertyValue
args1(0).Name = "FileName"
args1(0).Value = "https://de.wikipedia.org/wiki/Liste_der_Millionenst%C3%A4dte"
args1(1).Name = "FilterName"
args1(1).Value = "calc_HTML_WebQuery"
args1(2).Name = "Options"
args1(2).Value = "0 0"
args1(3).Name = "Source"
args1(3).Value = "HTML_4"

dispatcher.executeDispatch(document, ".uno:InsertExternalDataSource", "", 0, args1())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:InsertExternalDataSource", "", 0, Array())

rem ----------------------------------------------------------------------
dim args3(0) as new com.sun.star.beans.PropertyValue
args3(0).Name = "Nr"
args3(0).Value = 1

dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args3())

rem ----------------------------------------------------------------------
dim args4(0) as new com.sun.star.beans.PropertyValue
args4(0).Name = "Nr"
args4(0).Value = 2

dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args4())

rem ----------------------------------------------------------------------
dim args5(0) as new com.sun.star.beans.PropertyValue
args5(0).Name = "Nr"
args5(0).Value = 1

dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args5())

rem ----------------------------------------------------------------------
dim args6(0) as new com.sun.star.beans.PropertyValue
args6(0).Name = "Nr"
args6(0).Value = 2

dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args6())

rem ----------------------------------------------------------------------
dim args7(0) as new com.sun.star.beans.PropertyValue
args7(0).Name = "Nr"
args7(0).Value = 1

dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args7())

rem ----------------------------------------------------------------------
dim args8(0) as new com.sun.star.beans.PropertyValue
args8(0).Name = "Nr"
args8(0).Value = 2

dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args8())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:Insert", "", 0, Array())


end sub
