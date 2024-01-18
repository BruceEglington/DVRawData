unit DVRD_ShtImSmpReg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles,
  Grids, DBGrids, DBCtrls, AxCtrls, UFlexCelImport,
  UFlexCelGrid, UExcelAdapter, XLSAdapter;

type
  TfmSheetImportSmpReg = class(TForm)
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
    XLSAdapter1: TXLSAdapter;
    FlexCelImport1: TFlexCelImport;
    TabControl: TTabControl;
    Data: TFlexCelGrid;
    cbImportSpec: TComboBox;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure sbFindLastRowClick(Sender: TObject);
    procedure cbSheetNameChange(Sender: TObject);
    procedure dblcbImportSpecCloseUp(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmSheetImportSmpReg: TfmSheetImportSmpReg;

implementation

{$R *.DFM}

uses
  AllSorts, DVRD_varb, DVRD_dm_acs;

var
  iRec, iRecCount      : integer;

procedure TfmSheetImportSmpReg.bbOpenSheetClick(Sender: TObject);
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
    //dblcbImportSpec.KeyValue := dmDVRD.ElemNamesImportSpecName.AsString;
    try
      sbFindLastRowClick(Sender);
    except
    end;
  end;
end;

procedure TfmSheetImportSmpReg.Button1Click(Sender: TObject);
begin
  dmDVRD.ImportGroup.Open;
end;

procedure TfmSheetImportSmpReg.bbImportClick(Sender: TObject);
var
  j, k     : integer;
  iCode : integer;
  i : integer;
  FromRow, ToRow : integer;
  tmpStr : string;
  RecordIDCol,IncludedCol,
  SampleCol, FracCol,
  DataCol, VariableCol : integer;
  tmpSampleNo, tmpFrac,
  tmpRecordID,tmpIncluded,
  tmpIsoSystem, tmpNormalisingStandard,
  tmpStandardValue, tmpNormalisingFactor,
  tmpVariableID : string;
begin
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
  dmDVRD.ElemNames.Close;
  dmDVRD.ElemNames.Parameters.ParamByName('ImportSpecName').Value := cbImportSpec.Text;
  dmDVRD.ElemNames.Open;
  dmDVRD.ElemNames.First;
  RecordIDCol := dmDVRD.ElemNamesColumnNo.AsInteger;
  dmDVRD.ElemNames.Next;
  IncludedCol := dmDVRD.ElemNamesColumnNo.AsInteger;
  dmDVRD.ElemNames.Next;
  SampleCol := dmDVRD.ElemNamesColumnNo.AsInteger;
  dmDVRD.ElemNames.Next;
  FracCol := dmDVRD.ElemNamesColumnNo.AsInteger;
  dmDVRD.ElemNames.Next;
  with dmDVRD do
  begin
    SmpReg.Open;
    {import data into SmpReg}
    sbSheet.SimpleText := 'Importing SmpReg information';
    sbSheet.Refresh;
    repeat
      tmpIsoSystem := dmDVRD.ElemNamesIsoSystem.AsString;
      DataCol := dmDVRD.ElemNamesColumnNo.AsInteger;
      for i := FromRow to ToRow do
      begin
        tmpSampleNo := FlexCelImport1.CellValue[i,SampleCol];
        tmpFrac := FlexCelImport1.CellValue[i,FracCol];
        tmpRecordID := FlexCelImport1.CellValue[i,RecordIDCol];
        tmpIncluded := FlexCelImport1.CellValue[i,IncludedCol];
        tmpStr := FlexCelImport1.CellValue[i,DataCol];
        try
          SmpReg.Append;
          dmDVRD.SmpRegSampleNo.AsString := tmpSampleNo;
          dmDVRD.SmpRegFrac.AsString := tmpFrac;
          dmDVRD.SmpRegIsoSystem.AsString := tmpIsoSystem;
          dmDVRD.SmpRegIncluded.AsString := tmpIncluded;
          dmDVRD.SmpRegRecordID.AsString := tmpRecordID;
          dmDVRD.SmpReg.Post;
        except
        end;
      end;
      dmDVRD.ElemNames.Next;
    until dmDVRD.ElemNames.Eof;
    SmpReg.First;
    SmpReg.Close;
  end;
  dmDVRD.SmpData.EnableControls;
  dmDVRD.ElemNames.EnableControls;
  sbSheet.SimpleText := 'Finished importing all data';
  sbSheet.Refresh;
end;

procedure TfmSheetImportSmpReg.FormShow(Sender: TObject);
var
  i, j : integer;
begin
  TabControl.Visible := false;
  Splitter1.Visible := false;
  pDefinitions.Visible := false;
  meFromRow.Text := '2';
  meToRow.Text := '3';
  dmDVRD.SmpData.Open;
  dmDVRD.ImportGroup.Open;
  cbImportSpec.Items.Clear;
  repeat
    cbImportSpec.Items.Add(dmDVRD.ImportGroupImportSpecName.AsString);
    dmDVRD.ImportGroup.Next;
  until dmDVRD.ImportGroup.Eof;
  bbOpenSheetClick(Sender);
end;

procedure TfmSheetImportSmpReg.bbCancelClick(Sender: TObject);
begin
  ModalResult := mrNone;
  Close;
end;

procedure TfmSheetImportSmpReg.sbFindLastRowClick(Sender: TObject);
var
  iCode : integer;
  tmpStr : string;
  i,j : integer;
begin
  ImportSheetNumber := cbSheetName.ItemIndex+1;
  meToRow.Text := '';
  ToRow := 0;
  dmDVRD.Variables.First;
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
    j := ConvertCol2Int(dmDVRD.VariablesColumnLetter.AsString);
    ToRow := 0;
    repeat
      i := i + 1;
      if (i > ToRow) then ToRow := i-1;
      meToRow.Text := IntToStr(ToRow);
      try
        tmpStr := FlexCelImport1.CellValue[i,j];
      except
        tmpStr := '0.0';
      end;
    until (tmpStr = '');
  except
    //MessageDlg('Error reading data in column '+IntToStr(Data.Col),mtwarning,[mbOK],0);
  end;
  meToRow.Text := IntToStr(ToRow);
  RowCount[ImportSheetNumber] := ToRow + 1;
  Data.Row := 1;
end;

procedure TfmSheetImportSmpReg.cbSheetNameChange(Sender: TObject);
begin
  ImportSheetNumber := cbSheetName.ItemIndex+1;
  TabControl.TabIndex := cbSheetname.ItemIndex;
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  Data.ApplySheet;
  Data.Zoom := 70;
  Data.LoadSheet;
  sbFindLastRowClick(Sender);
end;

procedure TfmSheetImportSmpReg.dblcbImportSpecCloseUp(Sender: TObject);
begin
  ImportSheetNumber := cbSheetName.ItemIndex+1;
  TabControl.TabIndex := cbSheetname.ItemIndex;
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  Data.ApplySheet;
  Data.Zoom := 70;
  Data.LoadSheet;
  dmDVRD.ElemNames.Close;
  dmDVRD.ElemNames.Parameters.ParamByName('ImportSpecName').Value := cbImportSpec.Text;
  dmDVRD.ElemNames.Open;
  sbFindLastRowClick(Sender);
end;

procedure TfmSheetImportSmpReg.TabControlChange(Sender: TObject);
begin
  Data.ApplySheet;
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  cbSheetname.ItemIndex := TabControl.TabIndex;
  Data.Zoom := 70;
  Data.LoadSheet;
  sbFindLastRowClick(Sender);
end;

end.
