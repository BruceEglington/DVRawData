unit DVRD_ShtIm2;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles,
  Grids, DBGrids, DBCtrls, AxCtrls, DB,
  System.UITypes,
  VCL.FlexCel.Core, FlexCel.Render, FlexCel.Preview,
  FlexCel.XLSAdapter, Vcl.Tabs, System.ImageList, Vcl.ImgList,
  Vcl.VirtualImageList;

type
  TfmSheetImport2 = class(TForm)
    pControl: TPanel;
    sbSheet: TStatusBar;
    bbOpenSheet: TBitBtn;
    OpenDialogSprdSheet: TOpenDialog;
    gbDefineFields: TGroupBox;
    bbCancel: TBitBtn;
    gbDefineRows: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    eFromRow: TEdit;
    eToRow: TEdit;
    bbImport: TBitBtn;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    eImportSpecNameCol: TEdit;
    ePositionCol: TEdit;
    eCalledCol: TEdit;
    Label10: TLabel;
    sbFindLastRow: TSpeedButton;
    pSpreadSheet: TPanel;
    Panel3: TPanel;
    gbDefineTabSheet: TGroupBox;
    cbSheetName: TComboBox;
    Label21: TLabel;
    Label22: TLabel;
    Label12: TLabel;
    eColumnCol: TEdit;
    Splitter1: TSplitter;
    pDefinitions: TPanel;
    gbDefaults: TGroupBox;
    Splitter2: TSplitter;
    Label1: TLabel;
    eDefaultMinimum: TEdit;
    Memo1: TMemo;
    TabControl: TTabControl;
    Tabs: TTabSet;
    SheetData: TStringGrid;
    lFilePath: TLabel;
    VirtualImageList1: TVirtualImageList;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure sbFindLastRowClick(Sender: TObject);
    procedure cbSheetNameChange(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
    procedure SheetDataSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure TabsChange(Sender: TObject; NewTab: Integer;
      var AllowChange: Boolean);
  private
    { Private declarations }
    Xls : TXlsFile;
    function ConvertCol2Int(AnyString : string) : integer;
    procedure FillTabs;
    procedure ClearGrid;
    procedure FillGrid(const Formatted: boolean);
    function GetStringFromCell(iRow,iCol : integer) : string;
  public
    { Public declarations }
  end;

var
  fmSheetImport2: TfmSheetImport2;

implementation

uses allsorts, DVRD_varb, DVRD_dm;

{$R *.DFM}

var
  iRec, iRecCount      : integer;

procedure TfmSheetImport2.bbOpenSheetClick(Sender: TObject);
var
  tmpStr    : string;
  i : integer;
begin
  TabControl.Tabs.Clear;
  cbSheetname.Items.Clear;
  OpenDialogSprdSheet.InitialDir := DataPath;
  if OpenDialogSprdSheet.Execute then
  begin
    DataPath := ExtractFilePath(OpenDialogSprdSheet.FileName);
    lFilePath.Caption := OpenDialogSprdSheet.FileName;
    //Open the Excel file.
    if Xls = nil then Xls := TXlsFile.Create(false);
    xls.Open(OpenDialogSprdSheet.FileName);
    FillTabs;
    Tabs.TabIndex := Xls.ActiveSheet - 1;
    cbSheetName.ItemIndex := Xls.ActiveSheet - 1;
    FillGrid(true);
    {
    FlexCelImport1.OpenFile(OpenDialogSprdSheet.FileName);
    for i := 1 to FlexCelImport1.SheetCount do
    begin
      FlexCelImport1.ActiveSheet:=i;
      TabControl.Tabs.Add(FlexCelImport1.ActiveSheetName);
      cbSheetname.Items.Add(FlexCelImport1.ActiveSheetName);
    end;
    FlexCelImport1.ActiveSheet:=1;
    TabControl.TabIndex:=FlexCelImport1.ActiveSheet-1;
    cbSheetName.ItemIndex := FlexCelImport1.ActiveSheet-1;
    Data.LoadSheet;
    Data.Zoom := 70;
    pDefinitions.Visible := true;
    Splitter1.Visible := true;
    TabControl.Visible := true;
    }
    bbImport.Visible := true;
    SheetData.Row := 1;
    SheetData.Col := 1;
    pDefinitions.Visible := true;
    sbFindLastRowClick(Sender);
  end;
end;

procedure TfmSheetImport2.FillTabs;
var
  s: integer;
begin
  Tabs.Tabs.Clear;
  cbSheetname.Items.Clear;
  for s := 1 to Xls.SheetCount do
  begin
    Tabs.Tabs.Add(Xls.GetSheetName(s));
    cbSheetname.Items.Add(Xls.GetSheetName(s));
  end;
end;

procedure TfmSheetImport2.ClearGrid;
var
  r: integer;
begin
  for r := 1 to SheetData.RowCount do SheetData.Rows[r].Clear;
end;

procedure TfmSheetImport2.FillGrid(const Formatted: boolean);
var
  r, c, cIndex: Integer;
  v: TCellValue;
begin
  if Xls = nil then exit;

  if (Tabs.TabIndex + 1 <= Xls.SheetCount) and (Tabs.TabIndex >= 0) then Xls.ActiveSheet := Tabs.TabIndex + 1 else Xls.ActiveSheet := 1;
  //Clear data in previous grid
  ClearGrid;
  SheetData.RowCount := 1;
  SheetData.ColCount := 1;
  //FmtBox.Text := '';

  SheetData.RowCount := Xls.RowCount + 1; //Include fixed row
  SheetData.ColCount := Xls.ColCount + 1; //Include fixed col. NOTE THAT COLCOUNT IS SLOW. We use it here because we really need it. See the Performance.pdf doc.

  if (SheetData.ColCount > 1) then SheetData.FixedCols := 1; //it is deleted when we set the width to 1.
  if (SheetData.RowCount > 1) then SheetData.FixedRows := 1;

  for r := 1 to Xls.RowCount do
  begin
    //Instead of looping in all the columns, we will just loop in the ones that have data. This is much faster.
    for cIndex := 1 to Xls.ColCountInRow(r) do
    begin
      c := Xls.ColFromIndex(r, cIndex); //The real column.
      if Formatted then
      begin
        SheetData.Cells[c, r] := Xls.GetStringFromCell(r, c);
      end
      else
      begin
        v := Xls.GetCellValue(r, c);
        SheetData.Cells[c, r] := v.ToString;
      end;
    end;
  end;

  //Fill the row headers
  for r := 1 to SheetData.RowCount - 1 do
  begin
    SheetData.Cells[0, r] := IntToStr(r);
    SheetData.RowHeights[r] := Round(Xls.GetRowHeight(r) / TExcelMetrics.RowMultDisplay(Xls));
  end;

  //Fill the column headers
  for c := 1 to SheetData.ColCount - 1 do
  begin
    SheetData.Cells[c, 0] := TCellAddress.EncodeColumn(c);
    SheetData.ColWidths[c] := Round(Xls.GetColWidth(c) / TExcelMetrics.ColMult(Xls));
  end;

  //SelectedCell(1,1);

end;


procedure TfmSheetImport2.TabControlChange(Sender: TObject);
begin
  {
  Data.ApplySheet;
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  }
  cbSheetname.ItemIndex := TabControl.TabIndex;
  {
  Data.Zoom := 70;
  Data.LoadSheet;
  }
  //sbFindLastRowClick(Sender);
end;

procedure TfmSheetImport2.TabsChange(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
begin

end;

//procedure TfmSheetImport2.TabsClick(Sender: TObject);
//begin
//  FillGrid(true);
//end;

function TfmSheetImport2.ConvertCol2Int(AnyString : string) : integer;
var
  itmp    : integer;
  tmpStr  : string;
  tmpChar : char;
begin
    AnyString := UpperCase(AnyString);
    tmpStr := AnyString;
    ClearNull(tmpStr);
    Result := 0;
    if (length(tmpStr) = 2) then
    begin
      tmpChar := tmpStr[1];
      itmp := (ord(tmpChar)-64)*26;
      tmpChar := tmpStr[2];
      Result := itmp+(ord(tmpChar)-64);
    end else
    begin
      tmpChar := tmpStr[1];
      Result := (ord(tmpChar)-64);
    end;
end;

function TfmSheetImport2.GetStringFromCell(iRow,iCol : integer) : string;
begin
  Result := Xls.GetStringFromCell(iRow,iCol);
end;

procedure TfmSheetImport2.SheetDataSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  //SelectedCell(aCol, aRow);
  CanSelect := true;
end;

procedure TfmSheetImport2.cbSheetNameChange(Sender: TObject);
begin
  //ImportSheetNumber := cbSheetName.ItemIndex+1;

  Tabs.TabIndex := cbSheetname.ItemIndex;
  ClearGrid;
  FillGrid(true);
  {
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  Data.ApplySheet;
  Data.Zoom := 70;
  Data.LoadSheet;
  }
end;

procedure TfmSheetImport2.bbImportClick(Sender: TObject);
var
  j      : integer;
  iCode  : integer;
  i      : integer;
  tmpStr : string;
  tmpImportSpecName, tmpVariableID,
  tmpIsoSystem, tmpNormalisingStandard,
  tmpTechAbr,
  tmpMaterialAbr,
  tmpColumnLetter : string;
  tmpPos, tmpColumnNo : integer;
  tmpStandardValue, tmpNormalisingFactor : double;
  WasSuccessful : boolean;
  tVariableShouldBe, tIsoSystemShouldBe : string;
begin
  ImportSheetNumber := cbSheetName.ItemIndex + 1;
  SheetData.Row := 1;
  SheetData.Col := 1;
  FromRowValueString := UpperCase(eFromRow.Text);
  ToRowValueString := UpperCase(eToRow.Text);
  eImportSpecNameCol.Text := UpperCase(eImportSpecNameCol.Text);
  ePositionCol.Text := UpperCase(ePositionCol.Text);
  eCalledCol.Text := UpperCase(eCalledCol.Text);
  eColumnCol.Text := UpperCase(eColumnCol.Text);
  ImportSpecNameColStr := eImportSpecNameCol.Text;
  PositionColStr := ePositionCol.Text;
  CalledColStr := eCalledCol.Text;
  ColumnColStr := eColumnCol.Text;
  //check row variables
  iCode := 1;
  repeat
    //From Row
    tmpStr := eFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    //To Row
    if (iCode = 0) then
    begin
      tmpStr := eToRow.Text;
      Val(tmpStr, ToRow, iCode);
    end else
    begin
      ShowMessage('Incorrect value entered for From row');
      Exit;
    end;
    if (iCode = 0) then
    begin
      if (ToRow >= FromRow) then iCode := 0
                            else iCode := -1;
    end else
    begin
      ShowMessage('Incorrect value entered for To row');
      Exit;
    end;
    if (iCode <> 0)
      then begin
        ShowMessage('Incorrect values entered for rows to import');
        Exit;
      end;
  until (iCode = 0);
  //convert input columns for variables to numeric
  ImportSpecNameCol := ConvertCol2Int(eImportSpecNameCol.Text);
  //ShowMessage(IntToStr(ImportSpecNameCol)+'$'+eImportSpecNameCol.Text);
  PositionCol := ConvertCol2Int(ePositionCol.Text);
  //ShowMessage(IntToStr(PositionCol)+'$'+ePositionCol.Text);
  CalledCol := ConvertCol2Int(eCalledCol.Text);
  //ShowMessage(IntToStr(CalledCol)+'$'+eCalledCol.Text);
  ColumnCol := ConvertCol2Int(eColumnCol.Text);
  //ShowMessage(IntToStr(ColumnCol)+'$'+eColumnCol.Text);
  IsoSystemCol := ColumnCol+1;
  //TechAbrCol := ColumnCol+2;
  NormStandardCol := ColumnCol+2;
  StandardValueCol := ColumnCol+3;
  NormFactorCol := ColumnCol+4;
  //MaterialAbrCol := ColumnCol+6;
  dmDVRD.cdsImportSpecVariables.Open;
  dmDVRD.cdsImportSpecVariables.Last;
  if not (dmDVRD.cdsImportSpecVariables.Bof  and dmDVRD.cdsImportSpecVariables.Eof) then
  begin
    sbSheet.SimpleText := 'Clearing existing definitions';
    repeat
      dmDVRD.cdsImportSpecVariables.Delete;
    until dmDVRD.cdsImportSpecVariables.Bof;
    //try
      dmDVRD.cdsImportSpecVariables.ApplyUpdates(0);
    //except
    //end;
  end;
  sbSheet.SimpleText := 'Appending new definitions';
  for i := FromRow to ToRow do
  begin
    try
      //dmDVRD.cdsImportSpecVariables.Append;
      j := ImportSpecNameCol;
      tmpImportSpecName := Xls.GetStringFromCell(i,j);
      //tmpImportSpecName := FlexCelImport1.CellValue[i,j];
      //dmDVRD.cdsImportSpecVariablesImportSpecName.AsString := tmpStr;
      j := PositionCol;
      tmpStr := Xls.GetStringFromCell(i,j);
      //tmpStr := FlexCelImport1.CellValue[i,j];
      Val(tmpStr,tmpPos,iCode);
      //dmDVRD.cdsImportSpecVariablesPos.AsString := tmpStr;
      j := CalledCol;
      tmpVariableID := Xls.GetStringFromCell(i,j);
      //tmpVariableID := FlexCelImport1.CellValue[i,j];
      //dmDVRD.cdsImportSpecVariablesVariableID.AsString := tmpStr;
      //dmDVRD.cdsVarVar.Locate('VariableID',tmpVariableID,[loPartialKey]);
      //tVariableShouldBe := TRIM(dmDVRD.cdsVarVarVARIABLEID.AsString);
      //if (tVariableShouldBe <> tmpVariableID) then
      //begin
      //  MessageDlg('Unrecognised VariableID '+tmpVariableID,mtWarning,[mbOK],0);
      //end;
      j := ColumnCol;
      tmpColumnLetter := Xls.GetStringFromCell(i,j);
      tmpColumnLetter := UpperCase(tmpColumnLetter);
      //tmpColumnLetter := FlexCelImport1.CellValue[i,j];
      //ShowMessage(IntToStr(tmpPos)+'  '+tmpVariableID);
      //dmDVRD.cdsImportSpecVariablesColumnLetter.AsString := tmpStr;
      tmpColumnNo := ConvertCol2Int(tmpColumnLetter);
      //dmDVRD.cdsImportSpecVariablesColumnNo.AsString := IntToStr(ConvertCol2Int(tmpStr));
      //ShowMessage(tmpImportSpecName+'$'+tmpStr+'$'+tmpVariableID+'$'+tmpColumnLetter);
      j := NormStandardCol;
      tmpNormalisingStandard := Xls.GetStringFromCell(i,j);
      //tmpNormalisingStandard := FlexCelImport1.CellValue[i,j];
      //dmDVRD.cdsImportSpecVariablesNormalisingStandard.AsString := tmpStr;
      j := IsoSystemCol;
      tmpIsoSystem := Xls.GetStringFromCell(i,j);
      //tmpIsoSystem := FlexCelImport1.CellValue[i,j];
      //dmDVRD.cdsImportSpecVariablesIsoSystem.AsString := tmpStr;
      //dmDVRD.cdsVarVar.Locate('IsoSystem',tmpIsoSystem,[]);
      //tIsoSystemShouldBe := TRIM(dmDVRD.cdsVarVarISOSYSTEM.AsString);
      //if (tIsoSystemShouldBe <> tmpIsoSystem) then
      //begin
      //  MessageDlg('Unrecognised Isotope System '+tmpIsoSystem,mtWarning,[mbOK],0);
      //end;
      j := StandardValueCol;
      tmpStr := Xls.GetStringFromCell(i,j);
      //tmpStr := FlexCelImport1.CellValue[i,j];
      Val(tmpStr,tmpStandardValue,iCode);
      if (iCode > 0) then tmpStandardValue := 1.0;

      //dmDVRD.cdsImportSpecVariablesStandardValue.AsString := tmpStr;
      j := NormFactorCol;
      tmpStr := Xls.GetStringFromCell(i,j);
      //tmpStr := FlexCelImport1.CellValue[i,j];
      Val(tmpStr,tmpNormalisingFactor,iCode);
      //dmDVRD.cdsImportSpecVariablesNormalisingFactor.AsString := tmpStr;
      //j := TechAbrCol;
      //tmpTechAbr := FlexCelImport1.CellValue[i,j];
      //dmDVRD.cdsImportSpecVariables.Post;
      //j := MaterialAbrCol;
      //tmpMaterialAbr := FlexCelImport1.CellValue[i,j];
      dmDVRD.AddNewImportSpec(tmpImportSpecName,tmpPos,tmpVariableID,
                           tmpColumnLetter,tmpColumnNo,tmpIsoSystem,
                           tmpNormalisingStandard,tmpStandardValue,
                           tmpNormalisingFactor,WasSuccessful);
      if (not WasSuccessful) then
      begin
        MessageDlg('Unrecognised VariableID '+tmpVariableID+' '+tmpIsoSystem,mtWarning,[mbOK],0);
      end;
    except
      //iCode := 1;
      MessageDlg('Error importing data definitions at row '+IntToStr(i),mtWarning,[mbOK],0);
    end;
    try
      //dmDVRD.cdsImportSpecVariables.ApplyUpdates(0);
    except
    end;
  end;
  dmDVRD.cdsImportSpecVariables.First;
  //dmDVRD.cdsImportSpecVariables.Close;
  //dmDVRD.cdsImportSpecVariables.Open;
  if (iCode = 0) then
  begin
    ModalResult := mrOK;
  end else
  begin
    ModalResult := mrNone;
  end;
end;

procedure TfmSheetImport2.FormShow(Sender: TObject);
var
  i, j : integer;
begin
  bbImport.Visible := false;
  TabControl.Visible := false;
  Splitter1.Visible := false;
  eFromRow.Text := FromRowValueString;
  eToRow.Text := ToRowValueString;
  eImportSpecNameCol.Text := ImportSpecNameColStr;
  ePositionCol.Text := PositionColStr;
  eCalledCol.Text := CalledColStr;
  eColumnCol.Text := ColumnColStr;
  pDefinitions.Visible := false;
  bbOpenSheetClick(Sender);
end;


procedure TfmSheetImport2.bbCancelClick(Sender: TObject);
begin
  ModalResult := mrNone;
  Close;
end;


procedure TfmSheetImport2.sbFindLastRowClick(Sender: TObject);
var
  iCode : integer;
  tmpStr : string;
  i,j : integer;
  v : TCellValue;
begin
  ImportSheetNumber := cbSheetName.ItemIndex;
  eToRow.Text := '';
  ToRow := 0;
  iCode := 1;
  repeat
    tmpStr := eFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    if (iCode = 0) then
    begin
    end else
    begin
      ShowMessage('Incorrect value entered for From row');
      Exit;
    end;
  until (iCode = 0);
  try
    i := FromRow;
    j := ConvertCol2Int(eImportSpecNameCol.Text);
    ToRow := 0;
    repeat
      i := i + 1;
      if (i > ToRow) then ToRow := i-1;
      eToRow.Text := IntToStr(ToRow);
      v := Xls.GetCellValue(i,j);
      tmpStr := v.ToString;
      //tmpStr := FlexCelImport1.CellValue[i,j];
    until (tmpStr = '');
    eToRow.Text := IntToStr(ToRow);
    RowCount[ImportSheetNumber] := ToRow + 1;
  except
    //MessageDlg('Error reading data for main variable',mtwarning,[mbOK],0);
  end;
end;

end.
