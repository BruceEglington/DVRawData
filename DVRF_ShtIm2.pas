unit DVRF_ShtIm2;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles,
  Grids, DBGrids, DBCtrls, AxCtrls, UFlexCelImport, UExcelAdapter,
  XLSAdapter, UFlexCelGrid;

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
    Data: TFlexCelGrid;
    XLSAdapter1: TXLSAdapter;
    FlexCelImport1: TFlexCelImport;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure sbFindLastRowClick(Sender: TObject);
    procedure cbSheetNameChange(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmSheetImport2: TfmSheetImport2;

implementation

uses DVRD_varb, DVRD_dm_acs;

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

    bbImport.Visible := true;
    Data.Row := 1;
    Data.Col := 1;
    pDefinitions.Visible := true;
    sbFindLastRowClick(Sender);
  end;
end;


procedure TfmSheetImport2.bbImportClick(Sender: TObject);
var
  j      : integer;
  iCode  : integer;
  i      : integer;
  tmpStr : string;
begin
  ImportSheetNumber := cbSheetName.ItemIndex + 1;
  Data.Row := 1;
  Data.Col := 1;
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
  {check row variables}
  iCode := 1;
  repeat
    {From Row}
    tmpStr := eFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    {To Row}
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
  {check Default values}
  iCode := 1;
  {convert input columns for variables to numeric}
  ImportSpecNameCol := ConvertCol2Int(eImportSpecNameCol.Text);
  PositionCol := ConvertCol2Int(ePositionCol.Text);
  CalledCol := ConvertCol2Int(eCalledCol.Text);
  ColumnCol := ConvertCol2Int(eColumnCol.Text);
  dmDVRD.Variables.Open;
  dmDVRD.Variables.Last;
  if not (dmDVRD.Variables.Bof  and dmDVRD.Variables.Eof) then
  begin
    sbSheet.SimpleText := 'Clearing existing definitions';
    repeat
      dmDVRD.Variables.Delete;
    until dmDVRD.Variables.Bof;
  end;
  sbSheet.SimpleText := 'Appending new definitions';
  for i := FromRow to ToRow do
  begin
    try
      dmDVRD.Variables.Append;
      //Data.Row := i;
      //Data.Col := ImportSpecNameCol;
      j := ImportSpecNameCol;
      tmpStr := FlexCelImport1.CellValue[i,j];
      dmDVRD.VariablesImportSpecName.AsString := tmpStr;
      j := PositionCol;
      tmpStr := FlexCelImport1.CellValue[i,j];
      dmDVRD.VariablesPosition.AsString := tmpStr;
      j := CalledCol;
      tmpStr := FlexCelImport1.CellValue[i,j];
      dmDVRD.VariablesVariableID.AsString := tmpStr;
      j := ColumnCol;
      tmpStr := FlexCelImport1.CellValue[i,j];
      dmDVRD.CoranFacAllCOLUMN.AsString := tmpStr;
      j := TakeLogCol;
      tmpStr := FlexCelImport1.CellValue[i,j];
      dmDVRD.CoranFacAllTakeLog.AsString := tmpStr;
      j := WSumFacCol;
      tmpStr := FlexCelImport1.CellValue[i,j];
      dmDVRD.CoranFacAllWSumFac.AsString := tmpStr;
      dmDVRD.CoranFacAll.Post;
    except
    end;
  end;
  dmDVRD.CoranFacAll.First;
  repeat
    if (dmDVRD.CoranFacAllPOS.AsInteger > MM) then dmDVRD.CoranFacAll.Delete
                                             else dmDVRD.CoranFacAll.Next;
  until dmDVRD.CoranFacAll.Eof;
  dmDVRD.CoranFacAll.Close;
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
begin
  ImportSheetNumber := cbSheetName.ItemIndex+1;
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
    //Data.Row := i;
    //Data.Col := j;
    ToRow := 0;
    repeat
      //if (Data.Row > 48) then ShowMessage('repeat '+IntToStr(Data.Row)+'   '+FlexCelImport1.CellValue[Data.Row,Data.Col]);
      i := i + 1;
      //Data.Row := i;
      //Data.Col := j;
      if (i > ToRow) then ToRow := i-1;
      eToRow.Text := IntToStr(ToRow);
      tmpStr := FlexCelImport1.CellValue[i,j];
    until (tmpStr = '');
    eToRow.Text := IntToStr(ToRow);
    RowCount[ImportSheetNumber] := ToRow + 1;
  except
    //MessageDlg('Error reading data for main variable',mtwarning,[mbOK],0);
  end;
end;

procedure TfmSheetImport2.cbSheetNameChange(Sender: TObject);
begin
  ImportSheetNumber := cbSheetName.ItemIndex+1;
  TabControl.TabIndex := cbSheetname.ItemIndex;
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  Data.ApplySheet;
  Data.Zoom := 70;
  Data.LoadSheet;
  sbFindLastRowClick(Sender);
end;

procedure TfmSheetImport2.TabControlChange(Sender: TObject);
begin
  Data.ApplySheet;
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  cbSheetname.ItemIndex := TabControl.TabIndex;
  Data.Zoom := 70;
  Data.LoadSheet;
  sbFindLastRowClick(Sender);
end;

end.
