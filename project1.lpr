program project1;

{$mode objfpc}{$H+}

uses
  Classes, sysutils, variants, comobj;

procedure InsertIntoCell(const CellName, theText: WideString; Table: Variant);
var
  xTableText, xTableTextCursor: Variant;
begin
  xTableText := Table.getCellByName(CellName);
  xTableTextCursor := xTableText.createTextCursor();
  xTableTextCursor.setPropertyValue('CharColor', 16777215);
  xTableText.setString( theText );
end;

function CreateStruct(const Reflection: Variant; const strTypeName: WideString): Variant;
var
  IdlClass: Variant;
begin
  IdlClass := Reflection.forName(strTypeName);
  // https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1reflection_1_1XIdlClass.html
  IdlClass.createObject(Result);
end;

const
  ServiceName = 'com.sun.star.ServiceManager';

var
  ServiceManager, Desktop, Document, CoreReflection, LoadParams: Variant;
  xText, xTextCursor: Variant;
  xTable, xTableRows, xRow: Variant;
  xFrame, xFrameText, xFrameTextCursor, xSize: Variant;

begin
  if Assigned(InitProc) then
    TProcedure(InitProc);

  try
    ServiceManager := CreateOleObject(ServiceName);
  except
    WriteLn('Unable to start OO/LO.');
    Exit;
  end;

  // https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1frame_1_1Desktop.html
  Desktop := ServiceManager.CreateInstance('com.sun.star.frame.Desktop');

  CoreReflection := ServiceManager.CreateInstance('com.sun.star.reflection.CoreReflection');

  LoadParams := VarArrayCreate([0, -1], varVariant);

  // Desktop implements XComponentLoader https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XComponentLoader.html
  // Document is a TextDocument https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextDocument.html
  Document := Desktop.LoadComponentFromURL('private:factory/swriter', '_blank', 0, LoadParams);

  // insert some text.
  // For this purpose get the Text-Object of the document an create the
  // cursor. Now it is possible to insert a text at the cursor-position
  // via insertString

  //getting the text object
  xText := Document.getText();

  //create a cursor object
  xTextCursor := xText.createTextCursor();

  //inserting some Text
  xText.insertString( xTextCursor, 'The first line in the newly created text document.' + #10, False);

  //inserting a second line
  xText.insertString( xTextCursor, 'Now we''re in the second line' + #10, False);

  // insert a text table.
  // For this purpose create an
  // instance of com.sun.star.text.TextTable and initialize it. Now it can
  // be inserted at the cursor position via insertTextContent.
  // After that some properties are changed and some data is inserted.

  xTable := Document.createInstance('com.sun.star.text.TextTable');

  //initialize the text table with 4 columns an 4 rows
  xTable.initialize(4, 4);

  //insert the table
  xText.insertTextContent(xTextCursor, xTable, False);

  // get first Row
  xTableRows := xTable.getRows();
  xRow := xTableRows.getByIndex(0);

  // Change the BackColor
  xTable.setPropertyValue('BackTransparent', False);
  xTable.setPropertyValue('BackColor', 13421823);
  xRow.setPropertyValue('BackTransparent', False);
  xRow.setPropertyValue('BackColor', 6710932);

  InsertIntoCell('A1','FirstColumn', xTable);
  InsertIntoCell('B1','SecondColumn', xTable) ;
  InsertIntoCell('C1','ThirdColumn', xTable) ;
  InsertIntoCell('D1','SUM', xTable) ;


  //Insert Something in the text table
  xTable.getCellByName('A2').setValue(22.5);
  xTable.getCellByName('B2').setValue(5615.3);
  xTable.getCellByName('C2').setValue(-2315.7);
  xTable.getCellByName('D2').setFormula('sum <A2:C2>');

  xTable.getCellByName('A3').setValue(21.5);
  xTable.getCellByName('B3').setValue(615.3);
  xTable.getCellByName('C3').setValue(-315.7);
  xTable.getCellByName('D3').setFormula('sum <A3:C3>');

  xTable.getCellByName('A4').setValue(121.5);
  xTable.getCellByName('B4').setValue(-615.3);
  xTable.getCellByName('C4').setValue(415.7);
  xTable.getCellByName('D4').setFormula('sum <A4:C4>');

  // insert a colored text.
  // Get the propertySet of the cursor, change the CharColor and add a
  // shadow. Then insert the Text via InsertString at the cursor position.

  // Change the CharColor and add a Shadow
  xTextCursor.setPropertyValue('CharColor', Integer(255));
  xTextCursor.setPropertyValue('CharShadowed', True);

  // create a paragraph break
  // https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1text_1_1ControlCharacter.html
  xText.insertControlCharacter(xTextCursor, 0 {PARAGRAPH_BREAK}, False);

  //inserting colored Text
  xText.insertString(xTextCursor, ' This is a colored Text - blue with shadow' + #10, False);
  xText.insertControlCharacter(xTextCursor, 0, false);

  // insert a text frame.
  // create an instance of com.sun.star.text.TextFrame using the MSF of the
  // document. Change some properties an insert it.
  // Now get the text-Object of the frame an the corresponding cursor.
  // Insert some text via insertString.

  // Create a TextFrame
  xFrame := Document.createInstance('com.sun.star.text.TextFrame');

  // Set size
  xSize := CreateStruct(CoreReflection, 'com.sun.star.awt.Size');
  xSize.Height := 400;
  xSize.Width := 15000;
  xFrame.setSize(xSize);

  // alternatively
  {
  xFrame.Height := 400;
  xFrame.Width := 15000;
  }

  // get the property set of the text frame
  // https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1text.html#a470b1caeda4ff15fee438c8ff9e3d834acb5ff4ea4718bf38762e2da8c553f924
  xFrame.setPropertyValue('AnchorType', 1 {AS_CHARACTER});

  xText.insertTextContent(xTextCursor, xFrame, False);

  //getting the text object of Frame
  xFrameText := xFrame.getText();

  //create a cursor object
  xFrameTextCursor := xFrameText.createTextCursor();

  //inserting some Text
  xFrameText.insertString(xFrameTextCursor,
           'The first line in the newly created text frame.', False);

  xFrameText.insertString(xFrameTextCursor,
           #10 + 'With this second line the height of the frame raises.', false);

  //insert a paragraph break
  xText.insertControlCharacter(xTextCursor, 0, false );

  xTextCursor.setPropertyValue('CharColor', 65536);
  xTextCursor.setPropertyValue('CharShadowed', False);

  xText.insertString(xTextCursor, ' That''s all for now !!', false);
end.


