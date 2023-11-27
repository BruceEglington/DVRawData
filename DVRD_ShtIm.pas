unit DVRD_ShtIm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles,
  Grids, DBGrids, DBCtrls, AxCtrls,
  VCL.FlexCel.Core, FlexCel.Render, FlexCel.Preview,
  FlexCel.XlsAdapter,
  Vcl.Tabs, Data.DB, System.ImageList, Vcl.ImgList, Vcl.VirtualImageList;

type
  TfmSheetImport = class(TForm)
    pControl: TPanel;
    sbSheet: TStatusBar;
    bbOpenSheet: TBitBtn;
    OpenDialogSprdSheet: TOpenDialog;
    gbDefineFields: TGroupBox;
    bbCancel: TBitBtn;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    gbDefineRows: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    meFromRow: TEdit;
    meToRow: TEdit;
    bbImport: TBitBtn;
    Memo1: TMemo;
    Label4: TLabel;
    Label5: TLabel;
    Panel2: TPanel;
    Splitter1: TSplitter;
    gbDefineTabSheet: TGroupBox;
    cbSheetName: TComboBox;
    sbFindLastRow: TSpeedButton;
    pDefinitions: TPanel;
    Panel1: TPanel;
    TabControl: TTabControl;
    cbImportSpec: TComboBox;
    SheetData: TStringGrid;
    Tabs: TTabSet;
    pSpreadSheet: TPanel;
    lFilePath: TLabel;
    VirtualImageList1: TVirtualImageList;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure sbFindLastRowClick(Sender: TObject);
    procedure cbSheetNameChange(Sender: TObject);
    procedure dblcbImportSpecCloseUp(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
    procedure SheetDataSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
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
  fmSheetImport: TfmSheetImport;

implementation

{$R *.DFM}

uses
  DVRD_varb, allsorts, DVRD_dm;

var
  iRec, iRecCount      : integer;

function TfmSheetImport.ConvertCol2Int(AnyString : string) : integer;
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

function TfmSheetImport.GetStringFromCell(iRow,iCol : integer) : string;
begin
  Result := Xls.GetStringFromCell(iRow,iCol);
end;

procedure TfmSheetImport.SheetDataSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  //SelectedCell(aCol, aRow);
  CanSelect := true;
end;

procedure TfmSheetImport.bbOpenSheetClick(Sender: TObject);
var
  tmpStr    : string;
  i         : integer;
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
    }
    SheetData.Row := 1;
    SheetData.Col := 1;
    bbImport.Visible := true;
    pDefinitions.Visible := true;
    Splitter1.Visible := true;
    TabControl.Visible := true;
    //dblcbImportSpec.KeyValue := dmDVRD.ElemNamesImportSpecName.AsString;
    try
      sbFindLastRowClick(Sender);
    except
    end;
  end;
end;

procedure TfmSheetImport.bbImportClick(Sender: TObject);
var
  j, k     : integer;
  iCode : integer;
  i : integer;
  FromRow, ToRow : integer;
  tmpStr : string;
  RefnumCol,
  SampleCol, FracCol, ZoneCol,
  DataCol, VariableCol,
  NormStdCol,
  MaterialAbrCol,
  TechAbrCol : integer;
  tmpSampleNo, tmpFrac, tmpZoneID,
  tmpRefNum, tmpTechAbr,
  tmpMaterialAbr,
  tmpIsoSystem, tmpNormalisingStandard,
  tmpVariableID, tmpDataValue : string;
  tmpStandardValue, tmpNormalisingFactor : double;
  WasSuccessful : boolean;
  AreVariablesCorrect : boolean;
  tmpVarNamestr : string;
  tRefNumStr, tSampleNoStr, tFracStr, tZoneIDstr,
  tTechAbrStr, tMaterialAbrStr : string;
begin
  tRefNumStr := 'RefNum';
  tSampleNoStr := 'SampleNo';
  tFracStr := 'Frac';
  tZoneIDstr := 'ZoneID';
  tTechAbrStr := 'TechAbr';
  tMaterialAbrStr := 'MaterialAbr';
  ImportSheetNumber := cbSheetName.ItemIndex + 1;
  iCode := 1;
  repeat
    tmpStr := meFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    if (iCode = 0) then
    begin
      tmpStr := meToRow.Text;
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
  dmDVRD.qElemNames.Close;
  dmDVRD.qElemNames.ParamByName('ImportSpecName').Value := cbImportSpec.Text;
  dmDVRD.cdsElemNames.Close;
  dmDVRD.cdsElemNames.Open;
  dmDVRD.cdsElemNames.First;
  j := 1;
  repeat
    //ElementPos[j] := dmDVRD.cdsElemNamesColumnNo.AsInteger;
    j := j + 1;
    dmDVRD.cdsElemNames.Next;
  until dmDVRD.cdsElemNames.Eof;
  dmDVRD.cdsElemNames.First;
  Nox := j - 1;
  dmDVRD.cdsElemNames.First;
  //TechAbrCol := dmDVRD.cdsElemNamesColumnNo.AsInteger;
  //dmDVRD.cdsElemNames.Next;
  //NormStdCol := dmDVRD.cdsElemNamesColumnNo.AsInteger;
  //dmDVRD.cdsElemNames.Next;
  AreVariablesCorrect := true;
  RefNumCol := dmDVRD.cdsElemNamesColumnNo.AsInteger;
  tmpVarNamestr := dmDVRD.cdsElemNamesVARIABLEID.AsString;
  if (tmpVarNamestr <> tRefNumStr) then AreVariablesCorrect := false;
  dmDVRD.cdsElemNames.Next;
  SampleCol := dmDVRD.cdsElemNamesColumnNo.AsInteger;
  tmpVarNamestr := dmDVRD.cdsElemNamesVARIABLEID.AsString;
  if (tmpVarNamestr <> tSampleNoStr) then AreVariablesCorrect := false;
  dmDVRD.cdsElemNames.Next;
  FracCol := dmDVRD.cdsElemNamesColumnNo.AsInteger;
  tmpVarNamestr := dmDVRD.cdsElemNamesVARIABLEID.AsString;
  if (tmpVarNamestr <> tFracStr) then AreVariablesCorrect := false;
  dmDVRD.cdsElemNames.Next;
  ZoneCol := dmDVRD.cdsElemNamesColumnNo.AsInteger;
  tmpVarNamestr := dmDVRD.cdsElemNamesVARIABLEID.AsString;
  if (tmpVarNamestr <> tZoneIDstr) then AreVariablesCorrect := false;
  dmDVRD.cdsElemNames.Next;
  TechAbrCol := dmDVRD.cdsElemNamesColumnNo.AsInteger;
  tmpVarNamestr := dmDVRD.cdsElemNamesVARIABLEID.AsString;
  if (tmpVarNamestr <> tTechAbrStr) then AreVariablesCorrect := false;
  dmDVRD.cdsElemNames.Next;
  MaterialAbrCol := dmDVRD.cdsElemNamesColumnNo.AsInteger;
  tmpVarNamestr := dmDVRD.cdsElemNamesVARIABLEID.AsString;
  if (tmpVarNamestr <> tMaterialAbrStr) then AreVariablesCorrect := false;
  dmDVRD.cdsElemNames.Next;
  //ShowMessage(IntToStr(RefNumCol)+'  '+IntToStr(SampleCol)+'  '+IntToStr(FracCol)+'  '+IntToStr(TechAbrCol)+'  '+IntToStr(MaterialAbrCol));
  if not AreVariablesCorrect then
  begin
    ShowMessage('One or more of the variables is (are) incorrectly defined!!!');
    Exit;
  end;
  with dmDVRD do
  begin
    {import data into SmpData}
    sbSheet.Panels[1].Text := 'Importing raw data for '+IntToStr(Nox)+' variables';
    sbSheet.Refresh;
    {
    do for all rows in data spreadsheet
    for each variable
    store variableid, isosystem, normstd, stdval, normfac
    read sampleno and frac
    repeat through variables
      read data value
      append and store values
    }
    tmpStr := '';
    repeat
      tmpVariableID := dmDVRD.cdsElemNamesVariableID.AsString;
      tmpIsoSystem := dmDVRD.cdsElemNamesIsoSystem.AsString;
      tmpNormalisingStandard := dmDVRD.cdsElemNamesNormalisingStandard.AsString;
      tmpStandardValue := dmDVRD.cdsElemNamesStandardValue.AsFloat;
      tmpNormalisingFactor := dmDVRD.cdsElemNamesNormalisingFactor.AsFloat;
      //tmpTechAbr := dmDVRD.cdsElemNamesTECHABR.AsString;
      //tmpMaterialAbr := dmDVRD.cdsElemNamesMATERIALABR.AsString;
      DataCol := dmDVRD.cdsElemNamesColumnNo.AsInteger;
      //ShowMessage(tmpVariableID+'$'+tmpIsoSystem+'$'+tmpNormalisingStandard+'$'+FormatFloat('####0.0000',tmpStandardValue)+'$'+FormatFloat('####0.0000',tmpNormalisingFactor)+'$'+tmpTechAbr);
      //ShowMessage(IntToStr(DataCol));
      for i := FromRow to ToRow do
      begin
        tmpSampleNo := Trim(Xls.GetStringFromCell(i,SampleCol));
        tmpFrac := Trim(Xls.GetStringFromCell(i,FracCol));
        tmpZoneID := Trim(Xls.GetStringFromCell(i,ZoneCol));
        tmpTechAbr := Trim(Xls.GetStringFromCell(i,TechAbrCol));
        tmpMaterialAbr := Trim(Xls.GetStringFromCell(i,MaterialAbrCol));
        tmpRefNum := Trim(Xls.GetStringFromCell(i,RefNumCol));
        tmpDataValue := Trim(Xls.GetStringFromCell(i,DataCol));
        //ShowMessage(tmpSampleNo+'$'+tmpFrac+'$'+tmpTechAbr+'$'+tmpMaterialAbr+'$'+tmpRefNum+'$'+tmpDataValue+'$'+IntToStr(i));
        //ShowMessage(tmpDataValue+'$'+IntToStr(i));
        {
        if not UseDefaultTechnique then
          tmpTechAbr := Trim(FlexCelImport1.CellValue[i,TechAbrCol]);
        if not UseDefaultNormalisingStandard then
          tmpNormalisingStandard := Trim(FlexCelImport1.CellValue[i,NormStdCol]);
        }
        try
          //cdsSmpData.Append;
          //dmDVRD.cdsSmpDataSampleNo.AsString := tmpSampleNo;
          //dmDVRD.cdsSmpDataFrac.AsString := tmpFrac;
          //dmDVRD.cdsSmpDataIsoSystem.AsString := tmpIsoSystem;
          //dmDVRD.cdsSmpDataVariableID.AsString := tmpVariableID;
          //dmDVRD.cdsSmpDataNormalisingStandard.AsString := tmpNormalisingStandard;
          //dmDVRD.cdsSmpDataStandardValue.AsString := tmpStandardValue;
          //dmDVRD.cdsSmpDataNormalisingFactor.AsString := tmpNormalisingFactor;
          //dmDVRD.cdsSmpDataRefNum.AsString := tmpRefNum;
          if (tmpDataValue <> '') then
          begin
            try
              //dmDVRD.cdsSmpDataDataValue.AsString := tmpStr;
              dmDVRD.AddNewSmpData(tmpSampleNo,tmpFrac,tmpZoneID,tmpIsoSystem,
                           tmpVariableID,tmpDataValue,
                           tmpNormalisingStandard,tmpStandardValue,
                           tmpNormalisingFactor,tmpRefNum,
                           tmpTechAbr,tmpMaterialAbr,WasSuccessful);
            except
              //dmDVRD.cdsSmpDataDataValue.AsString := '';
            end;
          end else
          begin
            //dmDVRD.cdsSmpDataDataValue.AsString := '';
          end;
          //dmDVRD.cdsSmpData.Post;
        except
        end;
        try
          //dmDVRD.cdsSmpData.ApplyUpdates(0);
        except
        end;
        if (i mod 100 = 0) then
        begin
          sbSheet.Panels[0].Text := IntToStr(i);
          sbSheet.Refresh;
        end;
      end;
      Application.ProcessMessages;
      dmDVRD.cdsElemNames.Next;
    until dmDVRD.cdsElemNames.Eof;
    cdsSmpData.First;
  end;
  sbSheet.Panels[1].Text := 'Finished importing data for '+IntToStr(Nox)+' variables';
  dmDVRD.cdsSmpData.EnableControls;
  dmDVRD.cdsElemNames.EnableControls;
  sbSheet.Panels[1].Text := 'Finished importing all data';
  sbSheet.Refresh;
end;

procedure TfmSheetImport.FormShow(Sender: TObject);
var
  i, j : integer;
begin
  TabControl.Visible := false;
  Splitter1.Visible := false;
  pDefinitions.Visible := false;
  meFromRow.Text := '2';
  meToRow.Text := '3';
  dmDVRD.cdsSmpData.Open;
  dmDVRD.cdsImportGroup.Open;
  cbImportSpec.Items.Clear;
  repeat
    cbImportSpec.Items.Add(dmDVRD.cdsImportGroupImportSpecName.AsString);
    dmDVRD.cdsImportGroup.Next;
  until dmDVRD.cdsImportGroup.Eof;
  bbImport.Enabled := true;
  bbOpenSheetClick(Sender);
end;

procedure TfmSheetImport.bbCancelClick(Sender: TObject);
begin
  ModalResult := mrNone;
  Close;
end;

procedure TfmSheetImport.sbFindLastRowClick(Sender: TObject);
var
  iCode : integer;
  tmpStr : string;
  i,j : integer;
  v : TCellValue;
begin
  ImportSheetNumber := cbSheetName.ItemIndex;
  meToRow.Text := '';
  ToRow := 0;
  dmDVRD.cdsImportSpecVariables.First;
  iCode := 1;
  repeat
    tmpStr := meFromRow.Text;
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
    j := ConvertCol2Int(dmDVRD.cdsImportSpecVariablesCOLUMNLETTER.AsString);
    ToRow := 0;
    repeat
      i := i + 1;
      if (i > ToRow) then ToRow := i-1;
      meToRow.Text := IntToStr(ToRow);
      try
        v := Xls.GetCellValue(i,j);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,j];
      except
        tmpStr := '0.0';
      end;
    until (tmpStr = '');
  except
    //MessageDlg('Error reading data in column '+IntToStr(Data.Col),mtwarning,[mbOK],0);
  end;
  meToRow.Text := IntToStr(ToRow);
  RowCount[ImportSheetNumber] := ToRow + 1;
  SheetData.Row := 1;
end;

procedure TfmSheetImport.cbSheetNameChange(Sender: TObject);
begin
  Tabs.TabIndex := cbSheetname.ItemIndex;
  ClearGrid;
  FillGrid(true);
  {
  ImportSheetNumber := cbSheetName.ItemIndex+1;
  TabControl.TabIndex := cbSheetname.ItemIndex;
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  Data.ApplySheet;
  Data.Zoom := 70;
  Data.LoadSheet;
  }
  //sbFindLastRowClick(Sender);
end;

procedure TfmSheetImport.dblcbImportSpecCloseUp(Sender: TObject);
begin
  //ImportSheetNumber := cbSheetName.ItemIndex+1;
  Tabs.TabIndex := cbSheetname.ItemIndex;
  //FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  //Data.ApplySheet;
  //Data.Zoom := 70;
  //Data.LoadSheet;
  dmDVRD.qElemNames.Close;
  dmDVRD.qElemNames.ParamByName('ImportSpecName').Value := cbImportSpec.Text;
  dmDVRD.cdsElemNames.Close;
  dmDVRD.cdsElemNames.Open;
  //sbFindLastRowClick(Sender);
  bbImport.Enabled := true;
end;

procedure TfmSheetImport.TabControlChange(Sender: TObject);
begin
  //Data.ApplySheet;
  //FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  //cbSheetname.ItemIndex := TabControl.TabIndex;
  //Data.Zoom := 70;
  //Data.LoadSheet;
  //sbFindLastRowClick(Sender);
end;

procedure TfmSheetImport.FillTabs;
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

procedure TfmSheetImport.ClearGrid;
var
  r: integer;
begin
  for r := 1 to SheetData.RowCount do SheetData.Rows[r].Clear;
end;

procedure TfmSheetImport.FillGrid(const Formatted: boolean);
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

end.
