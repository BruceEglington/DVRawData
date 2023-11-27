unit DVRD_mn;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, ToolWin, ComCtrls,
  Printers, Menus, Mask, DBCtrls, Db, IniFiles,
  ActnList,MidasLib,
  DVRD_Varb, FlexCel.XlsAdapter,
  VCL.FlexCel.Core, FlexCel.Render, FlexCel.Preview,
  ActnMan, XPStyleActnCtrls, System.Actions, VclTee.TeeGDIPlus, VCLTee.Series,
  VCLTee.TeEngine, VCLTee.TeeProcs, VCLTee.Chart, VCLTee.TeeTools,
  System.ImageList, Vcl.ImgList, Vcl.VirtualImageList, VCL.Themes;

type
  TfmDVRDMain = class(TForm)
    ToolBar1: TToolBar;
    sbMain: TStatusBar;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    pc1: TPageControl;
    tsControl: TTabSheet;
    N1: TMenuItem;
    ImportRawData1: TMenuItem;
    Help: TMenuItem;
    About1: TMenuItem;
    PrinterSetupDialog1: TPrinterSetupDialog;
    PrintDialog1: TPrintDialog;
    Panel4: TPanel;
    SaveDialogSprdSheet: TSaveDialog;
    tsCheck: TTabSheet;
    DBNavigator8: TDBNavigator;
    DBGrid10: TDBGrid;
    Button1: TButton;
    DBGrid11: TDBGrid;
    DBGrid9: TDBGrid;
    DBNavigator7: TDBNavigator;
    DBNavigator13: TDBNavigator;
    DBGrid19: TDBGrid;
    DBGrid15: TDBGrid;
    DBGrid22: TDBGrid;
    DBGrid23: TDBGrid;
    Panel6: TPanel;
    Panel3: TPanel;
    dbgVariables: TDBGrid;
    Panel16: TPanel;
    Panel2: TPanel;
    Panel25: TPanel;
    dbgRefs: TDBGrid;
    Panel26: TPanel;
    dbgSmpData: TDBGrid;
    Panel24: TPanel;
    bbEmptySmpData: TBitBtn;
    dbnSmpData: TDBNavigator;
    Panel27: TPanel;
    dbnVariables: TDBNavigator;
    ActionManager1: TActionManager;
    ImportDataDefinitions1: TMenuItem;
    Splitter5: TSplitter;
    Splitter6: TSplitter;
    SaveDialogJPEG: TSaveDialog;
    Panel1: TPanel;
    dbnRefs: TDBNavigator;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    DBGrid33: TDBGrid;
    DBNavigator35: TDBNavigator;
    DBGrid37: TDBGrid;
    Button2: TButton;
    bCloseAll: TButton;
    boOpenAll: TButton;
    Button3: TButton;
    N5: TMenuItem;
    EmptySmpDataDataTables1: TMenuItem;
    Button4: TButton;
    bDelete: TButton;
    bDim4Smp: TButton;
    Options1: TMenuItem;
    UseDefaultTechnique1: TMenuItem;
    UseDefaultNormalisingStandard1: TMenuItem;
    EmptySmpListdatatable1: TMenuItem;
    EmptySmpFracdatatable1: TMenuItem;
    N2: TMenuItem;
    CopySamplestoSmpListdatatable1: TMenuItem;
    CopySamplesFractoSmpFracdatatable1: TMenuItem;
    Update1: TMenuItem;
    UpdateVariables1: TMenuItem;
    Updatematerialsanalysed1: TMenuItem;
    Updatetechniques1: TMenuItem;
    Updateisotopesystems1: TMenuItem;
    Updateall1: TMenuItem;
    Check1: TMenuItem;
    CheckVariables1: TMenuItem;
    Panel5: TPanel;
    Splitter1: TSplitter;
    DBGrid1: TDBGrid;
    tsGraph: TTabSheet;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Splitter2: TSplitter;
    Splitter3: TSplitter;
    Splitter4: TSplitter;
    Panel11: TPanel;
    Panel12: TPanel;
    Panel13: TPanel;
    Panel14: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    ChartWetherill: TChart;
    Series1: TLineSeries;
    Series2: TPointSeries;
    Series3: TPointSeries;
    Series4: TPointSeries;
    ChartTeraWasserburg: TChart;
    Series7: TLineSeries;
    Series8: TPointSeries;
    Series9: TPointSeries;
    Series10: TPointSeries;
    ChartAgeHfInitial: TChart;
    LineSeries1: TLineSeries;
    PointSeries1: TPointSeries;
    PointSeries2: TPointSeries;
    PointSeries3: TPointSeries;
    ChartAgeHfEpsilon: TChart;
    LineSeries2: TLineSeries;
    PointSeries4: TPointSeries;
    PointSeries5: TPointSeries;
    PointSeries6: TPointSeries;
    DBlcbIsotopeSystem: TDBLookupComboBox;
    DBlcbSample: TDBLookupComboBox;
    bUpdateGraphs: TButton;
    CheckGraphs1: TMenuItem;
    ChartTool1: TMarksTipTool;
    ChartTool2: TMarksTipTool;
    ChartTool3: TMarksTipTool;
    ChartTool4: TMarksTipTool;
    ChartTool5: TMarksTipTool;
    ChartTool6: TMarksTipTool;
    ChartTool7: TMarksTipTool;
    ChartTool8: TMarksTipTool;
    Button5: TButton;
    lCurveAges: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    eCurveAgeMinimum: TEdit;
    eCurveAgeMaximum: TEdit;
    CheckRatios1: TMenuItem;
    Splitter7: TSplitter;
    Panel15: TPanel;
    ChartErrorCorrelation: TChart;
    LineSeries3: TLineSeries;
    PointSeries7: TPointSeries;
    PointSeries8: TPointSeries;
    PointSeries9: TPointSeries;
    MarksTipTool1: TMarksTipTool;
    MarksTipTool2: TMarksTipTool;
    Splitter8: TSplitter;
    Chart2DM: TChart;
    LineSeries4: TLineSeries;
    PointSeries10: TPointSeries;
    PointSeries11: TPointSeries;
    PointSeries12: TPointSeries;
    MarksTipTool3: TMarksTipTool;
    MarksTipTool4: TMarksTipTool;
    Panel17: TPanel;
    dblcbSampleOnly: TDBLookupComboBox;
    bSampleOnly: TButton;
    bShowAll: TButton;
    Splitter9: TSplitter;
    Panel18: TPanel;
    ChartAgeDiscordance: TChart;
    LineSeries5: TLineSeries;
    PointSeries13: TPointSeries;
    PointSeries14: TPointSeries;
    PointSeries15: TPointSeries;
    MarksTipTool5: TMarksTipTool;
    MarksTipTool6: TMarksTipTool;
    VirtualImageList1: TVirtualImageList;
    bExit: TButton;
    Styles1: TMenuItem;
    procedure ImportRawData1Click(Sender: TObject);
    procedure bbEmptySmpDataClick(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ImportDataDefinitions1Click(Sender: TObject);
    procedure bbExitClick(Sender: TObject);
    procedure EmptySmpListdatatable1Click(Sender: TObject);
    procedure EmptySmpFracdatatable1Click(Sender: TObject);
    procedure CopySamplestoSmpListdatatable1Click(Sender: TObject);
    procedure CopySamplesFractoSmpFracdatatable1Click(Sender: TObject);
    procedure UseDefaultTechnique1Click(Sender: TObject);
    procedure UseDefaultNormalisingStandard1Click(Sender: TObject);
    procedure UpdateVariables1Click(Sender: TObject);
    procedure Updateisotopesystems1Click(Sender: TObject);
    procedure Updatematerialsanalysed1Click(Sender: TObject);
    procedure Updatetechniques1Click(Sender: TObject);
    procedure Updateall1Click(Sender: TObject);
    procedure CheckVariables1Click(Sender: TObject);
    procedure bUpdateGraphsClick(Sender: TObject);
    procedure CheckGraphs1Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure CheckRatios1Click(Sender: TObject);
    procedure Options1Click(Sender: TObject);
    procedure bSampleOnlyClick(Sender: TObject);
    procedure bShowAllClick(Sender: TObject);
    procedure dblcbSampleOnlyClick(Sender: TObject);
    procedure StyleClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    procedure DisableEnableControls(DisableEnable : string);
    procedure ReplotConcordia(GraphAgeFrom,GraphAgeTo : double);
    procedure RefreshConcordiaGraphs(AgeFrom,AgeTo : double);
    procedure GetConcordiaSampleData(ConcordanceFrom,ConcordanceTo : double);
    procedure PopulateConcordiaGraphs(ConcordanceFrom,ConcordanceTo : double);
    procedure ReplotDM(GraphAgeFrom,GraphAgeTo : double);
  public
    { Public declarations }
    procedure GetIniFile;
    procedure SetIniFile;
  end;

var
  fmDVRDMain: TfmDVRDMain;

implementation

uses
  System.IOUtils,
  DVRD_ShtIm, AllSorts, DVRD_ShtIm2,
  DVRD_About, DVRD_dm;

{$R *.DFM}
{$D+}
var
  ImportForm : TfmSheetImport;
  ImportForm2 : TfmSheetImport2;
  AboutForm : TAboutBox;

procedure TfmDVRDMain.ImportRawData1Click(Sender: TObject);
var
  i, ii : integer;
  DataImported : boolean;
begin
  sbMain.Panels[1].Text := 'Importing';
  sbMain.Refresh;
  dmDVRD.cdsSmpData.DisableControls;
  dbgSmpData.DataSource := nil;
  dbnSmpData.DataSource := nil;
  try
    try
      dmDVRD.cdsImportGroup.Open;
      dmDVRD.cdsImportGroup.First;
      dmDVRD.cdsImportSpecVariables.Open;
    except
    end;
    try
      ImportForm := TfmSheetImport.Create(Self);
      ImportForm.OpenDialogSprdSheet.FileName := 'DateViewRawData';
      //ImportForm.FillData;
      if (ImportForm.ShowModal = mrOK) then DataImported := true
                                       else DataImported := false;
    finally
      //ImportForm.Xls.CloseFile;
      ImportForm.Free;
    end; //finally
  finally
    try
      dmDVRD.cdsImportGroup.Close;
      //dmDVRD.Variables.Close;
    except
    end;
  end; //finally
  //ShowMessage('before close and open. RecordCount = '+IntToStr(dmDVRD.cdsSmpData.RecordCount));
  if DataImported then
  begin
    fmDVRDMain.Refresh;
    //ShowMessage('before close and open. RecordCount = '+IntToStr(dmDVRD.cdsSmpData.RecordCount));
    dmDVRD.cdsSmpData.Close;
    dmDVRD.cdsSmpData.Open;
    //ShowMessage('before close and open. RecordCount = '+IntToStr(dmDVRD.cdsSmpData.RecordCount));
    dmDVRD.cdsDataRefs.Close;
    dmDVRD.cdsDataRefs.Open;
    dbgSmpData.DataSource := dmDVRD.dsSmpData;
    dbnSmpData.DataSource := dmDVRD.dsSmpData;
    sbMain.Panels[1].Text := 'Finished';
    sbMain.Refresh;
    dmDVRD.cdsElemNames.First;
    try
      dmDVRD.cdsSmpData.EnableControls;
    except
      ShowMessage('Problem enablecontrols SmpData');
    end;
    sbMain.Panels[1].Text := 'New data imported';
    sbMain.Refresh;
  end else
  begin
    sbMain.Panels[1].Text := 'Import cancelled';
    sbMain.Refresh;
  end;
  //ShowMessage('Importing complete');
end;

procedure TfmDVRDMain.Options1Click(Sender: TObject);
begin
  MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
end;

procedure TfmDVRDMain.bbEmptySmpDataClick(Sender: TObject);
var
  WasSuccessful : boolean;
begin
  dmDVRD.EmptySmpData(WasSuccessful);
  if (WasSuccessful) then
  begin
    dmDVRD.cdsSmpData.Close;
    dmDVRD.cdsSmpData.Open;
  end else
  begin
    //
  end;
end;

procedure TfmDVRDMain.About1Click(Sender: TObject);
begin
  AboutForm := TAboutBox.Create(self);
  try
    AboutForm.ShowModal;
  finally
    AboutForm.Free;
  end;
end;

procedure TfmDVRDMain.GetIniFile;
var
  PublicPath : string;
  ProgramPath : string;
  AppIni   : TIniFile;
  tmpStr   : string;
  iCode    : integer;
  DBMonitor,
  DriverName,
  LibraryName, VendorLib, GetDriverFunc,
  DBUserName, DBPassword,DBSpecific,DBSQLDialectStr,DBCharSet : string;
begin
  DriverName := 'DevartFirebird';
  LibraryName := 'dbexpida41.dll';
  VendorLib := 'fbclient.dll';
  GetDriverFunc := 'getSQLDriverFirebird';
  PublicPath := TPath.GetHomePath;
  ProgramPath := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName));
  CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  IniFilename := CommonFilePath + 'DateViewRawData.ini';
  AppIni := TIniFile.Create(IniFileName);
  try
    GlobalChosenStyle := 'Windows';
    ExportPath := AppIni.ReadString('Paths','Spreadsheet exports path',CommonFilePath);
    FlexTemplatePath := AppIni.ReadString('Paths','Spreadsheet template path',CommonFilePath);
    DBPath := AppIni.ReadString('Paths','DB path',CommonFilePath+'DVrawdata\Data\Firebird\DVRAWDATAV30.FDB');
    GlobalChosenStyle := AppIni.ReadString('Styles','Chosen style','Windows');
    if (GlobalChosenStyle = '') then GlobalChosenStyle := 'Windows';
    dmDVRD.ChosenStyle := GlobalChosenStyle;
    ImportSpecNameColStr := AppIni.ReadString('ColumnDefinitions','ImportSpecNameColStr','A');
    PositionColStr := AppIni.ReadString('ColumnDefinitions','PositionColStr','B');
    CalledColStr := AppIni.ReadString('ColumnDefinitions','CalledColStr','C');
    ColumnColStr := AppIni.ReadString('ColumnDefinitions','ColumnColStr','D');
    tmpStr := AppIni.ReadString('Defaults','DefaultMinimum','1.0e-6');
    Val(tmpStr,DefaultMinimum,iCode);
    if (iCode > 0) then DefaultMinimum := 1.0e-6;
    LibraryName := AppIni.ReadString('Parameters','LibraryName','dbexpida41.dll');
    DBVendorLib := AppIni.ReadString('Parameters','VendorLib',ProgramPath+'FBCLIENT.DLL');
    GetDriverFunc := AppIni.ReadString('Parameters','GetDriverFunc','getSQLDriverFirebird');
    DriverName := AppIni.ReadString('Parameters','DriverName','DevartFirebird');
    DriverName := AppIni.ReadString('Parameters','DriverName','DevartFirebird');
    DBUserName := AppIni.ReadString('Parameters','User_Name','SYSDBA');
    DBPassword := AppIni.ReadString('Parameters','Password','masterkey');
    DBSQLDialectStr := AppIni.ReadString('Parameters','SQLDialect','3');
    DBCharSet := AppIni.ReadString('Parameters','Charset','ASCII');
    DBMonitor := AppIni.ReadString('Monitor','DBMonitor','Inactive');
  finally
    AppIni.Free;
  end;
  //define connection parameters for DateViewRawData connection
  dmDVRD.DateViewRawData.Connected := false;
  dmDVRD.DateViewRawData.Params.Clear;
  dmDVRD.DateViewRawData.Params.Append('DriverName='+DriverName);
  dmDVRD.DateViewRawData.Params.Append('Database='+DBPath);
  dmDVRD.DateViewRawData.Params.Append('LibraryName='+LibraryName);
  dmDVRD.DateViewRawData.Params.Append('GetDriverFunc='+GetDriverFunc);
  dmDVRD.DateViewRawData.Params.Append('VendorLib='+DBVendorLib);
  dmDVRD.DateViewRawData.Params.Append('User_Name='+DBUserName);
  dmDVRD.DateViewRawData.Params.Append('Password='+DBPassword);
  dmDVRD.DateViewRawData.Params.Append('SQLDialect='+DBSQLDialectStr);
  dmDVRD.DateViewRawData.Params.Append('Charset='+DBCharSet);
  dmDVRD.DateViewRawData.Params.Append('LocaleCode=0000');
  dmDVRD.DateViewRawData.Params.Append('DevartInterBase TransIsolation=ReadCommitted');
  dmDVRD.DateViewRawData.Params.Append('WaitOnLocks=True');
  dmDVRD.DateViewRawData.Params.Append('CharLength=1');
  dmDVRD.DateViewRawData.Params.Append('EnableBCD=True');
  dmDVRD.DateViewRawData.Params.Append('OptimizedNumerics=False');
  dmDVRD.DateViewRawData.Params.Append('LongStrings=True');
  dmDVRD.DateViewRawData.Params.Append('UseQuoteChar=False');
  dmDVRD.DateViewRawData.Params.Append('FetchAll=False');
  dmDVRD.DateViewRawData.Params.Append('UseUnicode=False');
  if (DBMonitor = 'Active') then
  begin
    dmDVRD.SQLMonitor1.Active := true;
  end else
  begin
    dmDVRD.SQLMonitor1.Active := false;
  end;
end;

procedure TfmDVRDMain.SetIniFile;
var
  PublicPath : string;
  ProgramPath : string;
  DBVendorLibStr, DBPathStr : string;
  AppIni   : TIniFile;
begin
  PublicPath := TPath.GetHomePath;
  ProgramPath := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName));
  CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  IniFilename := CommonFilePath + 'DateViewRawData.ini';
  //DBVendorLibStr := ProgramPath+'FBCLIENT.DLL';
  //DBPathStr := CommonFilePath+'DVrawdata\Data\Firebird\DVRAWDATAV25.FDB';
  DBVendorLibStr := DBVendorLib;
  DBPathStr := DBPath;
  AppIni := TIniFile.Create(IniFileName);
  try
    //AppIni.WriteString('Paths','Data path',DataPath);
    AppIni.WriteString('Paths','Spreadsheet exports path',ExportPath);
    AppIni.WriteString('Paths','Spreadsheet template path',FlexTemplatePath);
    AppIni.WriteString('Styles','Chosen style',GlobalChosenStyle);
    AppIni.WriteString('ColumnDefinitions','ImportSpecNameColStr',ImportSpecNameColStr);
    AppIni.WriteString('ColumnDefinitions','PositionColStr',PositionColStr);
    AppIni.WriteString('ColumnDefinitions','CalledColStr',CalledColStr);
    AppIni.WriteString('ColumnDefinitions','ColumnColStr',ColumnColStr);
    AppIni.WriteString('Defaults','DefaultMinimum',FormatFloat('##0.0000e-00',DefaultMinimum));

    AppIni.WriteString('Paths','DB path',DBPathStr);
    AppIni.WriteString('Parameters','VendorLib',DBVendorLibStr);
  finally
    AppIni.Free;
  end;
end;

procedure TfmDVRDMain.StyleClick(Sender: TObject);
var
  StyleName : String;
  i : integer;
begin
  //get style name
  StyleName := TMenuItem(Sender).Caption;
  StyleName := StringReplace(StyleName, '&', '',
    [rfReplaceAll,rfIgnoreCase]);
  GlobalChosenStyle := StyleName;
  dmDVRD.ChosenStyle := GlobalChosenStyle;
  //set active style
  Application.ProcessMessages;
  TStyleManager.SetStyle(GlobalChosenStyle);
  dmDVRD.ChosenStyle := GlobalChosenStyle;
  Application.ProcessMessages;
  //check the currently selected menu item
  (Sender as TMenuItem).Checked := true;
  //uncheck all other style menu items
  for i := 0 to Styles1.Count-1 do
  begin
    if not Styles1.Items[i].Equals(Sender) then
      Styles1.Items[i].Checked := false;
  end;
  for i := 0 to Styles1.Count-1 do
  begin
    if Styles1.Items[i].Checked then GlobalChosenStyle := StringReplace(Styles1.Items[i].Caption, '&', '',
    [rfReplaceAll,rfIgnoreCase]);
  end;
  TStyleManager.SetStyle(GlobalChosenStyle);
  try
    dmDVRD.ChosenStyle := GlobalChosenStyle;
  finally
    dmDVRD.ChosenStyle := GlobalChosenStyle;
  end;
end;

procedure TfmDVRDMain.Updateall1Click(Sender: TObject);
begin
  MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
  //UpdateIsotopeSystems1Click(Sender);
  //UpdateMaterialsAnalysed1Click(Sender);
  //UpdateTechniques1Click(Sender);
  //UpdateVariables1Click(Sender);
end;

procedure TfmDVRDMain.Updateisotopesystems1Click(Sender: TObject);
begin
  MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
end;

procedure TfmDVRDMain.Updatematerialsanalysed1Click(Sender: TObject);
begin
  MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
end;

procedure TfmDVRDMain.Updatetechniques1Click(Sender: TObject);
begin
  MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
end;

procedure TfmDVRDMain.UpdateVariables1Click(Sender: TObject);
begin
  MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
end;

procedure TfmDVRDMain.UseDefaultNormalisingStandard1Click(Sender: TObject);
begin
  //
  MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
end;

procedure TfmDVRDMain.UseDefaultTechnique1Click(Sender: TObject);
begin
  //
  MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
end;

procedure TfmDVRDMain.FormShow(Sender: TObject);
var
  ii, j : integer;
begin
  tsCheck.TabVisible := false;
  tsGraph.TabVisible := false;
  pc1.ActivePage := tsControl;
  DefaultMinimum := 1.0e-6;
  GetIniFile;
  with dmDVRD do
  begin
    try
      DateViewRawData.Connected := false;
    except
    end;
    //DateViewRawData.ConnectionString := 'FILE NAME='+ADODataLinkFile;
    //DateViewRawData.Provider := ADODataLinkFile;
    //DateViewRawData.Open('admin','');
    try
      DateViewRawData.Connected := true;
    except
    end;
  end;
  with dmDVRD do
  begin
    cdsSmpData.Open;
    cdsImportSpecVariables.Open;
    dmDVRD.cdsDataRefs.Open;
  end;
  FromRowValueString := '2';
  ToRowValueString := '2';
  dmDVRD.cdsSamples.Open;
  //cbSampleOnly.Checked := true;
end;

procedure TfmDVRDMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SetIniFile;
end;

procedure TfmDVRDMain.FormCreate(Sender: TObject);
var
  Style: String;
  Item: TMenuItem;
begin
  //MainPanel.Align := alClient;
  { position the form at the top of display }
  //Left := 0;
  //Top := 0;
  //Application.OnHint := true;
  GetIniFile;
  TStyleManager.TrySetStyle(GlobalChosenStyle);
  //Add child menu items based on available styles.
  for Style in TStyleManager.StyleNames do
  begin
    Item := TMenuItem.Create(Styles1);
    Item.Caption := Style;
    Item.OnClick := StyleClick;
    if TStyleManager.ActiveStyle.Name = Style then
      Item.Checked := true;
    Styles1.Add(Item);
  end;
end;

procedure TfmDVRDMain.ImportDataDefinitions1Click(Sender: TObject);
var
  tVariableID, tSystemShouldBe, tSystemDeclared : string;
begin
  try
    ImportForm2 := TfmSheetImport2.Create(Self);
    ImportForm2.OpenDialogSprdSheet.FileName := 'DateViewRawDataDefinitions';
    if (ImportForm2.ShowModal = mrOK) then
    begin
      sbMain.Panels[1].Text := 'New definitions imported';
      sbMain.Refresh;
    end else
    begin
      sbMain.Panels[1].Text := 'Cancelled import of new definitions';
      sbMain.Refresh;
    end;
  finally
    ImportForm2.Free;
  end;
  dmDVRD.cdsImportSpecVariables.Close;
  dmDVRD.cdsImportSpecVariables.Open;
  Application.ProcessMessages;
  dbgVariables.DataSource := dmDVRD.dsImportSpecVariables;
  // Check definition variable associations
  if (sbMain.Panels[1].Text = ' definitions imported') then
  begin
    // 1. Check that variables match the associated isotope system
    dmDVRD.cdsVarVar.Open;
    dmDVRD.cdsImportSpecVariables.First;
    repeat
      tVariableID := dmDVRD.cdsImportSpecVariablesVARIABLEID.AsString;
      tSystemDeclared := TRIM(dmDVRD.cdsImportSpecVariablesISOSYSTEM.AsString);
      dmDVRD.cdsVarVar.Locate('VariableID',tVariableID,[]);
      tSystemShouldBe := TRIM(dmDVRD.cdsVarVarISOSYSTEM.AsString);
      if ((tSystemShouldBe <> 'na') AND (tSystemDeclared <> tSystemShouldbe)) then
      begin
        dmDVRD.cdsImportSpecVariables.Edit;
        dmDVRD.cdsImportSpecVariablesISOSYSTEM.AsString := tSystemShouldBe;
        dmDVRD.cdsImportSpecVariables.Post;
        MessageDlg('Incorrect sytem association for '+tVariableID,mtWarning,[mbOK],0);
      end;
      dmDVRD.cdsImportSpecVariables.Next;
    until dmDVRD.cdsImportSpecVariables.Eof;
    // 2. Check variables which may be asociated with multiple systems
  end;
  dmDVRD.cdsVarVar.Close;
end;

procedure TfmDVRDMain.bbExitClick(Sender: TObject);
begin
  SetIniFile;
  dmDVRD.DateViewRawData.Connected := false;
  Close;
end;

procedure TfmDVRDMain.bSampleOnlyClick(Sender: TObject);
begin
  dmDVRD.cdsSmpData.Filter := 'SAMPLENO = '+QuotedStr(dblcbSampleOnly.KeyValue);
  dmDVRD.cdsSmpData.Filtered := true;
end;

procedure TfmDVRDMain.bShowAllClick(Sender: TObject);
begin
  dmDVRD.cdsSmpData.Filtered := false;
end;

procedure TfmDVRDMain.bUpdateGraphsClick(Sender: TObject);
const
  MaxDiscordance  : double = 10.0;
var
  iCode : integer;
  tXTW, tYTW,
  tX, tY, tAge,
  tR, tZ,
  MinX, MaxX,
  MinY, MaxY,
  MinXTW, MaxXTW,
  MinYTW, MaxYTW,
  MinXEC, MaxXEC,
  MinYEC, MaxYEC,
  MinX2DM, MaxX2DM,
  MinY2DM, MaxY2DM,
  MinDiscord, MaxDiscord,
  MinAge, MaxAge : double;
  tInclude : integer;
  WasSuccessful : boolean;
  i : integer;
  tFracStr : string;
  CurveAgeMinimum, CurveAgeMaximum : double;
  tDiscordance : double;
begin
  try
    dmDVRD.CopySmpList(WasSuccessful);
  except
  end;
  dmDVRD.qRawSmp.Close;
  dmDVRD.cdsRawSmp.Close;
  dmDVRD.qRawSmp.SQL.Clear;
  dmDVRD.qRawSmp.SQL.Add('SELECT DISTINCT SMPDATA.SAMPLENO,SMPDATA.FRAC,');
  dmDVRD.qRawSmp.SQL.Add(' SMPDATA.ZONEID,');
  dmDVRD.qRawSmp.SQL.Add('  VARREGASSOC.REGASSOCID');
  dmDVRD.qRawSmp.SQL.Add('FROM SMPDATA,SMPLIST,VARREGASSOC,SMPFRAC');
  dmDVRD.qRawSmp.SQL.Add('WHERE SMPDATA.SAMPLENO=SMPLIST.SAMPLENO');
  dmDVRD.qRawSmp.SQL.Add('AND VARREGASSOC.REGASSOCID=:RegAssocID');
  dmDVRD.qRawSmp.SQL.Add('AND SMPDATA.SAMPLENO=SMPFRAC.SAMPLENO');
  dmDVRD.qRawSmp.SQL.Add('AND SMPDATA.FRAC=SMPFRAC.FRAC');
  dmDVRD.qRawSmp.SQL.Add('AND ( SMPDATA.SAMPLENO = ');
  dmDVRD.qRawSmp.SQL.Add(QuotedStr(DBlcbSample.KeyValue));
  dmDVRD.qRawSmp.SQL.Add(')');
  for i := 0 to 3 do
  begin
    ChartWetherill.Series[i].Clear;
    ChartTeraWasserburg.Series[i].Clear;
    ChartErrorCorrelation.Series[i].Clear;
    ChartAgeHfInitial.Series[i].Clear;
    ChartAgeHfEpsilon.Series[i].Clear;
    Chart2DM.Series[i].Clear;
    ChartAgeDiscordance.Series[i].Clear;
  end;
  MinX := 9.0e9;
  MaxX := -9.0e9;
  MinY := 9.0e9;
  MaxY := -9.0e9;
  MinXTW := 9.0e9;
  MaxXTW := -9.0e9;
  MinYTW := 9.0e9;
  MaxYTW := -9.0e9;
  MinAge := 9.0e9;
  MaxAge := -9.0e9;
  MinXEC := 9.0e9;
  MaxXEC := -9.0e9;
  MinYEC := 9.0e9;
  MaxYEC := -9.0e9;
  MinX2DM := 9.0e9;
  MaxX2DM := -9.0e9;
  MinY2DM := 9.0e9;
  MaxY2DM := -9.0e9;
  MinDiscord := 9.0e9;
  MaxDiscord := -9.0e9;
  //plot the U-Pb data
  //ShowMessage('UPb 1');
  dmDVRD.qRawSmp.ParamByName('RegAssocID').AsString := 'UPb';
  dmDVRD.cdsRawSmp.Open;
  repeat
    //if (DBlcbIsotopeSystem.KeyValue = 'UPb') then
    //begin
      tFracStr := dmDVRD.cdsRawDataXFRAC.AsString;
      tDiscordance := dmDVRD.cdsRawDiscordanceDATAVALUE.AsFloat;
      tX := dmDVRD.cdsRawDataXDATAVALUE.AsFloat;
      tY := dmDVRD.cdsRawDataYDATAVALUE.AsFloat;
      tR := dmDVRD.cdsRawDataRDATAVALUE.AsFloat;
      tZ := dmDVRD.cdsRawDataZDATAVALUE.AsFloat;
      tXTW := 0.0;
      if (tY > 0.0) then tXTW := 1.0/tY;
      tYTW := dmDVRD.cdsRawDataZDATAVALUE.AsFloat;
      tAge := dmDVRD.cdsRawAgePrefDATAVALUE.AsFloat;
      if (dmDVRD.cdsRawDataIncludeDATAVALUE.AsFloat > 0.5) then tInclude := 1
                                                           else tInclude := 0;
      if (tInclude = 1) then
      begin
        if (Abs(tDiscordance) > MaxDiscordance) then tInclude := 0;
      end;
      if (tInclude=1) then
      begin
        if (tX > MaxX) then MaxX := tX;
        if (tX < MinX) then MinX := tX;
        if (tY > MaxY) then MaxY := tY;
        if (tY < MinY) then MinY := tY;
        if (tXTW > MaxXTW) then MaxXTW := tXTW;
        if (tXTW < MinXTW) then MinXTW := tXTW;
        if (tYTW > MaxYTW) then MaxYTW := tYTW;
        if (tYTW < MinYTW) then MinYTW := tYTW;
        if (tAge > MaxAge) then MaxAge := tAge;
        if (tAge < MinAge) then MinAge := tAge;
        if (tR > MaxXEC) then MaxXEC := tR;
        if (tR < MinXEC) then MinXEC := tR;
        if (tZ > MaxYEC) then MaxYEC := tZ;
        if (tZ < MinYEC) then MinYEC := tZ;
        if (tDiscordance > MaxDiscord) then MaxDiscord := tDiscordance;
        if (tDiscordance < MinDiscord) then MinDiscord := tDiscordance;
        //ChartWetherill.Series[2].AddXY(tX,tY);
        ChartWetherill.Series[2].AddXY(tX,tY,tFracStr);
        if (tXTW > 0.0) then ChartTeraWasserburg.Series[2].AddXY(tXTW,tYTW,tFracStr);
        ChartErrorCorrelation.Series[2].AddXY(tR,tZ,tFracStr);
        ChartAgeDiscordance.Series[2].AddXY(tDiscordance,tAge,tFracStr);
      end else
      begin
        ChartWetherill.Series[3].AddXY(tX,tY,tFracStr);
        if (tXTW > 0.0) then ChartTeraWasserburg.Series[3].AddXY(tXTW,tYTW,tFracStr);
        ChartErrorCorrelation.Series[3].AddXY(tR,tZ,tFracStr);
        ChartAgeDiscordance.Series[3].AddXY(tDiscordance,tAge,tFracStr);
      end;
    //end;
    dmDVRD.cdsRawSmp.Next;
  until dmDVRD.cdsRawSmp.Eof;
  //plot curves
  CurveAgeMinimum := 0.0;
  CurveAgeMaximum := 4570;
  Val(Trim(eCurveAgeMinimum.Text),CurveAgeMinimum,iCode);
  Val(Trim(eCurveAgeMaximum.Text),CurveAgeMaximum,iCode);
  //ReplotConcordia(0.000001,1.0*Round(1.1*MaxAge));
  ReplotConcordia(CurveAgeMinimum,CurveAgeMaximum);
  //plot tick marks
  //nothing here yet
  //ShowMessage('UPb 2');
  if (DBlcbIsotopeSystem.KeyValue <> 'UPb') then
  begin
    //plot the data
    dmDVRD.cdsRawSmp.Close;
    dmDVRD.qRawSmp.ParamByName('RegAssocID').AsString := DBlcbIsotopeSystem.KeyValue;
    dmDVRD.cdsRawSmp.Open;
    //ShowMessage('LuHf 1');
    repeat
        tFracStr := dmDVRD.cdsRawDataXFRAC.AsString;
        tDiscordance := dmDVRD.cdsRawDiscordanceDATAVALUE.AsFloat;
        //tX := dmDVRD.cdsRawAGEPREFDATAVALUE.AsFloat;
        tY := dmDVRD.cdsRawDataInitDATAVALUE.AsFloat;
        tYTW := dmDVRD.cdsRawDataEpsDATAVALUE.AsFloat;
        tAge := dmDVRD.cdsRawAgePrefDATAVALUE.AsFloat;
        tX := tAge;
        tXTW := tAge;
        tZ := dmDVRD.cdsRawDataDMDATAVALUE.AsFloat;
        if (dmDVRD.cdsRawDataIncludeDATAVALUE.AsFloat > 0.5) then tInclude := 1
                                                             else tInclude := 0;
        if (tInclude = 1) then
        begin
          if (Abs(tDiscordance) > MaxDiscordance) then tInclude := 0;
        end;
        //ShowMessage(tFracStr+'  tX = '+FormatFloat('###0.00',tX)+'  tY = '+FormatFloat('###0.000000',tY)+'  tYTW = '+FormatFloat('###0.000',tYTW));
        if (tInclude=1) then
        begin
          if (tX > MaxX) then MaxX := tX;
          if (tX < MinX) then MinX := tX;
          if (tY > MaxY) then MaxY := tY;
          if (tY < MinY) then MinY := tY;
          if (tXTW > MaxXTW) then MaxXTW := tXTW;
          if (tXTW < MinXTW) then MinXTW := tXTW;
          if (tYTW > MaxYTW) then MaxYTW := tYTW;
          if (tYTW < MinYTW) then MinYTW := tYTW;
          if (tAge > MaxAge) then MaxAge := tAge;
          if (tAge < MinAge) then MinAge := tAge;
          if (tAge > MaxX2DM) then MaxX2DM := tR;
          if (tAge < MinX2DM) then MinX2DM := tR;
          if (tZ > MaxY2DM) then MaxY2DM := tZ;
          if (tZ < MinY2DM) then MinY2DM := tZ;
          if (tY > 0.0) then ChartAgeHfInitial.Series[2].AddXY(tX,tY,tFracStr);
          if (tY > 0.0) then ChartAgeHfEpsilon.Series[2].AddXY(tXTW,tYTW,tFracStr);
          Chart2DM.Series[2].AddXY(tAge,tZ,tFracStr);
        end else
        begin
          if (tY > 0.0) then ChartAgeHfInitial.Series[3].AddXY(tX,tY,tFracStr);
          if (tY > 0.0) then ChartAgeHfEpsilon.Series[3].AddXY(tXTW,tYTW,tFracStr);
          Chart2DM.Series[3].AddXY(tAge,tZ,tFracStr);
        end;
      dmDVRD.cdsRawSmp.Next;
    until dmDVRD.cdsRawSmp.Eof;
    //ShowMessage('LuHf 2');
    //plot curves
    ReplotDM(CurveAgeMinimum,CurveAgeMaximum);
    //ReplotDM(0.000001,1.0*Round(1.1*MaxAge));
    //plot tick marks
    //nothing here yet
    //ShowMessage('LuHf 3');
  end;
end;

procedure TfmDVRDMain.Button5Click(Sender: TObject);
var
  i, iMax : integer;
begin
  //ShowMessage('PublicPath = '+'***'+PublicPath+'***');
  ShowMessage('CommonFilePath = '+'***'+CommonFilePath+'***');
  //ShowMessage('ProgramPath = '+'***'+ProgramPath+'***');
  ShowMessage('IniFileName = '+'***'+IniFileName+'***');
  iMax := dmDVRD.DateViewRawData.Params.Count;
  i := 0;
  repeat
    ShowMessage(dmDVRD.DateViewRawData.Params.Names[i]+'***'+dmDVRD.DateViewRawData.Params.ValueFromIndex[i]);
    i := i + 1;
  until (i > dmDVRD.DateViewRawData.Params.Count-1);
end;

procedure TfmDVRDMain.CheckGraphs1Click(Sender: TObject);
begin
  dmDVRD.cdsIsotopeSystems.Close;
  dmDVRD.cdsSamples.Close;
  tsGraph.TabVisible := true;
  pc1.ActivePage := tsGraph;
  dmDVRD.cdsIsotopeSystems.Open;
  dmDVRD.cdsSamples.Open;
  if (dmDVRD.cdsSamples.RecordCount > 0) then
  begin
    DBlcbIsotopeSystem.KeyValue := dmDVRD.cdsIsotopeSystemsISOSYSTEM.AsString;
    DBlcbSample.KeyValue := dmDVRD.cdsSamplesSAMPLENO.AsString;
  end else
  begin
    MessageDlg('No samples in SmpList table',mtWarning,[mbOK],0);
  end;
end;

procedure TfmDVRDMain.CheckRatios1Click(Sender: TObject);
var
  ErrorFound : boolean;
  VariablesWithErrors : string;
begin
  //first check U-Pb data
  dmDVRD.cdsSamples.Open;
  dmDVRD.cdsSamples.First;
  repeat
    dmDVRD.qRawSmp.Close;
    dmDVRD.cdsRawSmp.Close;
    dmDVRD.qRawSmp.SQL.Clear;
    dmDVRD.qRawSmp.SQL.Add('SELECT DISTINCT SMPDATA.SAMPLENO,SMPDATA.FRAC,');
    dmDVRD.qRawSmp.SQL.Add(' SMPDATA.ZONEID,');
    dmDVRD.qRawSmp.SQL.Add('  VARREGASSOC.REGASSOCID');
    dmDVRD.qRawSmp.SQL.Add('FROM SMPDATA,SMPLIST,VARREGASSOC,SMPFRAC');
    dmDVRD.qRawSmp.SQL.Add('WHERE SMPDATA.SAMPLENO=SMPLIST.SAMPLENO');
    dmDVRD.qRawSmp.SQL.Add('AND VARREGASSOC.REGASSOCID=:RegAssocID');
    dmDVRD.qRawSmp.SQL.Add('AND SMPDATA.SAMPLENO=SMPFRAC.SAMPLENO');
    dmDVRD.qRawSmp.SQL.Add('AND SMPDATA.FRAC=SMPFRAC.FRAC');
    dmDVRD.qRawSmp.SQL.Add('AND ( SMPDATA.SAMPLENO = ');
    dmDVRD.qRawSmp.SQL.Add(QuotedStr(dmDVRD.cdsSamplesSAMPLENO.AsString));
    dmDVRD.qRawSmp.SQL.Add(')');
    dmDVRD.qRawSmp.ParamByName('RegAssocID').AsString := 'UPb';
    dmDVRD.cdsRawSmp.Open;
    // check 207Pb*/235U
    VariablesWithErrors := '';
    dmDVRD.cdsRawSmp.First;
    ErrorFound := false;
    repeat
      if (dmDVRD.cdsRawDataXDATAVALUE.AsFloat < Minimum207Pb235U ) then ErrorFound := true;
      if (dmDVRD.cdsRawDataXDATAVALUE.AsFloat > Maximum207Pb235U ) then ErrorFound := true;
      dmDVRD.cdsRawSmp.Next;
    until dmDVRD.cdsRawSmp.Eof;
    if (ErrorFound) then VariablesWithErrors := VariablesWithErrors + '  207Pb*/235U';
    // check 206Pb*/238U
    dmDVRD.cdsRawSmp.First;
    ErrorFound := false;
    repeat
      if (dmDVRD.cdsRawDataYDATAVALUE.AsFloat < Minimum206Pb238U ) then ErrorFound := true;
      if (dmDVRD.cdsRawDataYDATAVALUE.AsFloat > Maximum206Pb238U ) then ErrorFound := true;
      dmDVRD.cdsRawSmp.Next;
    until dmDVRD.cdsRawSmp.Eof;
    if (ErrorFound) then VariablesWithErrors := VariablesWithErrors + '  206Pb*/238U';
    // check 207Pb*/206Pb*
    dmDVRD.cdsRawSmp.First;
    ErrorFound := false;
    repeat
      if (dmDVRD.cdsRawDataZDATAVALUE.AsFloat < Minimum207Pb206Pb ) then ErrorFound := true;
      if (dmDVRD.cdsRawDataZDATAVALUE.AsFloat > Maximum207Pb206Pb ) then ErrorFound := true;
      dmDVRD.cdsRawSmp.Next;
    until dmDVRD.cdsRawSmp.Eof;
    if (ErrorFound) then VariablesWithErrors := VariablesWithErrors + '  207Pb*/206Pb*';
    // check 207Pb*/206Pb* error correlation
    dmDVRD.cdsRawSmp.First;
    ErrorFound := false;
    repeat
      if (dmDVRD.cdsRawDataRDATAVALUE.AsFloat < Minimumr207Pb206Pb ) then ErrorFound := true;
      if (dmDVRD.cdsRawDataRDATAVALUE.AsFloat > Maximumr207Pb206Pb ) then ErrorFound := true;
      dmDVRD.cdsRawSmp.Next;
    until dmDVRD.cdsRawSmp.Eof;
    if (ErrorFound) then VariablesWithErrors := VariablesWithErrors + '  Error correlation';
    {
    // check 238U/206Pb*
    dmDVRD.cdsRawSmp.First;
    ErrorFound := false;
    repeat
      if (dmDVRD.cdsRawDataYDATAVALUE.AsFloat < Minimum238U206Pb ) then ErrorFound := true;
      if (dmDVRD.cdsRawDataYDATAVALUE.AsFloat > Maximum238U206Pb ) then ErrorFound := true;
      dmDVRD.cdsRawSmp.Next;
    until dmDVRD.cdsRawSmp.Eof;
    if (ErrorFound) then VariablesWithErrors := VariablesWithErrors + '  238U/206Pb*';
    }
    if (ErrorFound) then MessageDlg(dmDVRD.cdsSamplesSAMPLENO.AsString+' has issues for '+VariablesWithErrors,mtWarning,[mbOK],0);
    dmDVRD.cdsSamples.Next;
  until dmDVRD.cdsSamples.Eof;
  MessageDlg('Ratio checks completed',mtInformation,[mbOK],0);
end;

procedure TfmDVRDMain.CheckVariables1Click(Sender: TObject);
var
  tVariableID,
  tSystemDeclared,
  tSystemShouldBe : string;
begin
  // Check that variables are consistent with expected isotope systems
  //
  // First check that 'Age_preferred' is associated with Isotope System 'UPb'
  // for cases where MaterialAbr=zr
  dmDVRD.qCheckVariables.SQL.Clear;
  dmDVRD.qCheckVariables.SQL.Add('UPDATE SMPDATA');
  dmDVRD.qCheckVariables.SQL.Add('SET SMPDATA.ISOSYSTEM='+''''+'UPb'+'''');
  dmDVRD.qCheckVariables.SQL.Add('WHERE SMPDATA.MATERIALABR='+''''+'zr'+'''');
  dmDVRD.qCheckVariables.SQL.Add('AND SMPDATA.VARIABLEID='+''''+'Age_preferred'+'''');
  dmDVRD.qCheckVariables.SQL.Add('AND SMPDATA.ISOSYSTEM <> '+''''+'UPb'+'''');
  //ShowMessage(dmDVRD.qCheckVariables.SQL.Text);
  try
    dmDVRD.qCheckVariables.ExecSQL(false);
  except
  end;
  // Now check that 'sAge_preferred' is associated with Isotope System 'UPb'
  // for cases where MaterialAbr=zr
  dmDVRD.qCheckVariables.SQL.Clear;
  dmDVRD.qCheckVariables.SQL.Add('UPDATE SMPDATA');
  dmDVRD.qCheckVariables.SQL.Add('SET SMPDATA.ISOSYSTEM='+''''+'UPb'+'''');
  dmDVRD.qCheckVariables.SQL.Add('WHERE SMPDATA.MATERIALABR='+''''+'zr'+'''');
  dmDVRD.qCheckVariables.SQL.Add('AND SMPDATA.VARIABLEID='+''''+'sAge_preferred'+'''');
  dmDVRD.qCheckVariables.SQL.Add('AND SMPDATA.ISOSYSTEM <> '+''''+'UPb'+'''');
  try
    dmDVRD.qCheckVariables.ExecSQL(false);
  except
  end;
  // Then check that 'eAge_preferred' is associated with Isotope System 'UPb'
  // for cases where MaterialAbr=zr
  dmDVRD.qCheckVariables.SQL.Clear;
  dmDVRD.qCheckVariables.SQL.Add('UPDATE SMPDATA');
  dmDVRD.qCheckVariables.SQL.Add('SET SMPDATA.ISOSYSTEM='+''''+'UPb'+'''');
  dmDVRD.qCheckVariables.SQL.Add('WHERE SMPDATA.MATERIALABR='+''''+'zr'+'''');
  dmDVRD.qCheckVariables.SQL.Add('AND SMPDATA.VARIABLEID='+''''+'eAge_preferred'+'''');
  dmDVRD.qCheckVariables.SQL.Add('AND SMPDATA.ISOSYSTEM <> '+''''+'UPb'+'''');
  try
    dmDVRD.qCheckVariables.ExecSQL(false);
  except
  end;
  dmDVRD.cdsSmpData.Refresh;
  // 1. Check that variables match the associated isotope system
  // as defined in the VarVar table and correct if necessary
  dmDVRD.cdsVarVar.Open;
  dmDVRD.cdsSmpData.First;
  repeat
    tVariableID := dmDVRD.cdsSmpDataVARIABLEID.AsString;
    tSystemDeclared := TRIM(dmDVRD.cdsSmpDataISOSYSTEM.AsString);
    dmDVRD.cdsVarVar.Locate('VariableID',tVariableID,[]);
    tSystemShouldBe := TRIM(dmDVRD.cdsVarVarISOSYSTEM.AsString);
    if ((tSystemShouldBe <> 'na') AND (tSystemDeclared <> tSystemShouldbe)) then
    begin
      try
        dmDVRD.cdsSmpData.Edit;
        dmDVRD.cdsSmpDataISOSYSTEM.AsString := tSystemShouldBe;
        dmDVRD.cdsSmpData.Post;
        //MessageDlg('Incorrect sytem association for '+tVariableID,mtWarning,[mbOK],0);
      except
      end;
    end;
    dmDVRD.cdsSmpData.Next;
  until dmDVRD.cdsSmpData.Eof;
  // 2. Check variables which may be asociated with multiple systems
  dmDVRD.cdsVarVar.Close;
end;

procedure TfmDVRDMain.CopySamplesFractoSmpFracdatatable1Click(Sender: TObject);
var
  WasSuccessful : boolean;
begin
  dmDVRD.EmptySmpFrac(WasSuccessful);
  dmDVRD.CopySmpFrac(WasSuccessful);
end;

procedure TfmDVRDMain.CopySamplestoSmpListdatatable1Click(Sender: TObject);
var
  WasSuccessful : boolean;
begin
  dmDVRD.cdsSmpData.Filtered := false;
  dmDVRD.EmptySmpList(WasSuccessful);
  dmDVRD.CopySmpList(WasSuccessful);
  dmDVRD.cdsSamples.Close;
  dmDVRD.cdsSamples.Open;
end;

procedure TfmDVRDMain.dblcbSampleOnlyClick(Sender: TObject);
begin
  dmDVRD.cdsSamples.Close;
  dmDVRD.cdsSamples.Open;
end;

procedure TfmDVRDMain.DisableEnableControls(DisableEnable : string);
begin
  if (DisableEnable = 'Disable') then
  begin
    dmDVRD.cdsSmpData.DisableControls;
  end else
  begin
    dmDVRD.cdsSmpData.EnableControls;
  end;
end;

procedure TfmDVRDMain.EmptySmpFracdatatable1Click(Sender: TObject);
var
  WasSuccessful : boolean;
begin
  dmDVRD.EmptySmpFrac(WasSuccessful);
  if (WasSuccessful) then
  begin
    //dmDVRD.cdsSmpData.Close;
    //dmDVRD.cdsSmpData.Open;
  end else
  begin
    //
  end;
end;

procedure TfmDVRDMain.EmptySmpListdatatable1Click(Sender: TObject);
var
  WasSuccessful : boolean;
begin
  dmDVRD.EmptySmpList(WasSuccessful);
  if (WasSuccessful) then
  begin
    dmDVRD.EmptySmpFrac(WasSuccessful);
    //dmDVRD.cdsSmpList.Close;
    //dmDVRD.cdsSmpList.Open;
  end else
  begin
    //
  end;
  dmDVRD.cdsSamples.Close;
  dmDVRD.cdsSamples.Open;
end;

procedure TfmDVRDMain.ReplotConcordia(GraphAgeFrom,GraphAgeTo : double);
var
  i : integer;
  iMaxConcordiaAge, iConcordiaAgeIncrement, iLabelConcordiaAgeIncrement : integer;
  t207235, t206238, t207206, t238206 : double;
  tAge : double;
  iCode : integer;
  MinimumUncertainty,
  ConcordanceFrom, ConcordanceTo,
  DiscordanceFrom, DiscordanceTo : double;
begin
  iMaxConcordiaAge := 4500;
  iConcordiaAgeIncrement := 20;
  iLabelConcordiaAgeIncrement := 500;
  tAge := GraphAgeTo;
  iMaxConcordiaAge := Trunc(tAge);
  ChartWetherill.Series[0].Clear;
  ChartWetherill.Series[1].Clear;
  ChartTeraWasserburg.Series[0].Clear;
  ChartTeraWasserburg.Series[1].Clear;
  i := 0;
  tAge := GraphAgeFrom;
  i := Trunc(tAge);
  repeat
    //ShowMessage('i = '+IntToStr(i));
    dmDVRD.CalculateConcordiaForAge(1.0e6*i,t207235,t206238,t207206,t238206);
    ChartWetherill.Series[0].AddXY(t207235,t206238);
    if ((i > 0) and (t238206 < Maximum238U206Pb)) then ChartTeraWasserburg.Series[0].AddXY(t238206,t207206);
    if ((i mod iLabelConcordiaAgeIncrement) = 0) then
    begin
      //ShowMessage('Tick mark at '+IntToStr(i));
      ChartWetherill.Series[1].AddXY(t207235,t206238);
      if ((i > 0)  and (t238206 < Maximum238U206Pb)) then ChartTeraWasserburg.Series[1].AddXY(t238206,t207206);
    end;
    i := i + iConcordiaAgeIncrement;
  until (i >= iMaxConcordiaAge);
end;

procedure TfmDVRDMain.ReplotDM(GraphAgeFrom,GraphAgeTo : double);
const
  zero = 0.0;
var
  i : integer;
  iMaxModelAge, iModelAgeIncrement, iLabelModelAgeIncrement : integer;
  tAge, tUR, tDM, tEpsDM : double;
  DC,      // decay constant
  CRPDP,   // crustal average parent-daughter ratio (present day) for T2DM calculations
  URPDP,   // uniform reservoir parent-daughter present-day ratio
  URDP,    // uniform reservoir daughter present-day ratio
  DMPDP,   // depleted mantle reservoir parent-daughter present-day ratio
  DMDP : double;  // depleted mantle reservoir daughter present-day ratio
  iCode : integer;
  MinimumUncertainty,
  ConcordanceFrom, ConcordanceTo,
  DiscordanceFrom, DiscordanceTo : double;
begin
  iMaxModelAge := Round(GraphAgeTo);
  iModelAgeIncrement := 20;
  iLabelModelAgeIncrement := 500;
  tAge := 0.0;
  if (dblcbIsotopeSystem.KeyValue = 'LuHf') then
  begin
    URPDP := 0.0336; //Bouvier et al (2008)
    URDP := 0.282785; //Bouvier et al (2008)
    DC := 1.867e-11; //Scherer et al (2001) and Soderlund et al (2004)
    DMPDP := 0.0384; //Bouvier et al (2008)
    DMDP := 0.283250; //Bouvier et al (2008)
    CRPDP := 0.015;
    tUR := URDP - URPDP*(exp(tAge*1.0e6*DC)-1.0);
    tDM := DMDP - DMPDP*(exp(tAge*1.0e6*DC)-1.0);
    i := 0;
    repeat
      tAge := 1.0e6*i;
      tUR := URDP - URPDP*(exp(tAge*DC)-1.0);
      tDM := DMDP - DMPDP*(exp(tAge*DC)-1.0);
      tEpsDM := -999.0;
      if (tUR <> 0.0) then tEpsDM := 10000.0*(tDM/tUR - 1.0);
      tAge := 1.0*i;
      ChartAgeHfInitial.Series[0].AddXY(tAge,tDM);
      ChartAgeHfEpsilon.Series[0].AddXY(tAge,tEpsDM);
      //ChartAgeHfInitial.Series[0].AddXY(tAge,tDM);
      Chart2DM.Series[0].AddXY(tAge,tAge);
      if ((i mod iLabelModelAgeIncrement) = 0) then
      begin
        ChartAgeHfInitial.Series[1].AddXY(tAge,tDM);
        ChartAgeHfEpsilon.Series[1].AddXY(tAge,tEpsDM);
        //ChartAgeHfInitial.Series[1].AddXY(tAge,tDM);
        Chart2DM.Series[1].AddXY(tAge,tAge);
      end;
      i := i + iModelAgeIncrement;
    until (i >= iMaxModelAge);
  end;
  if (dblcbIsotopeSystem.KeyValue = 'SmNd') then
  begin
    URPDP := 0.1967; //
    URDP := 0.51264; //
    DC := 6.54e-12; //
    DMPDP := 0.2136; //
    DMDP := 0.513073536; //
    CRPDP := 0.11;
    tUR := URDP - URPDP*(exp(tAge*1.0e6*DC)-1.0);
    tDM := DMDP - DMPDP*(exp(tAge*1.0e6*DC)-1.0);
    i := 0;
    repeat
      tAge := 1.0e6*i;
      tUR := URDP - URPDP*(exp(tAge*DC)-1.0);
      tDM := DMDP - DMPDP*(exp(tAge*DC)-1.0);
      tEpsDM := -999.0;
      if (tUR <> 0.0) then tEpsDM := 10000.0*(tDM/tUR - 1.0);
      tAge := 1.0*i;
      ChartAgeHfInitial.Series[0].AddXY(tAge,tDM);
      ChartAgeHfEpsilon.Series[0].AddXY(tAge,tEpsDM);
      //ChartAgeHfInitial.Series[0].AddXY(tAge,tDM);
      Chart2DM.Series[0].AddXY(tAge,tAge);
      if ((i mod iLabelModelAgeIncrement) = 0) then
      begin
        ChartAgeHfInitial.Series[1].AddXY(tAge,tDM);
        ChartAgeHfEpsilon.Series[1].AddXY(tAge,tEpsDM);
        //ChartAgeHfInitial.Series[1].AddXY(tAge,tDM);
        Chart2DM.Series[1].AddXY(tAge,tAge);
      end;
      i := i + iModelAgeIncrement;
    until (i >= iMaxModelAge);
  end;
  if (dblcbIsotopeSystem.KeyValue = 'RbSr') then
  begin
    URPDP := 0.0847; //
    URDP := 0.7047; //
    DC := 1.42e-11; //
    DMPDP := 0.05; //  need to check and change
    DMDP := 0.70273029; //
    CRPDP := 0.1;  // need to check and change
    tUR := URDP - URPDP*(exp(tAge*1.0e6*DC)-1.0);
    tDM := DMDP - DMPDP*(exp(tAge*1.0e6*DC)-1.0);
    i := 0;
    repeat
      tAge := 1.0e6*i;
      tUR := URDP - URPDP*(exp(tAge*DC)-1.0);
      tDM := DMDP - DMPDP*(exp(tAge*DC)-1.0);
      tEpsDM := -999.0;
      if (tUR <> 0.0) then tEpsDM := 10000.0*(tDM/tUR - 1.0);
      tAge := 1.0*i;
      ChartAgeHfInitial.Series[0].AddXY(tAge,tDM);
      ChartAgeHfEpsilon.Series[0].AddXY(tAge,tEpsDM);
      //ChartAgeHfInitial.Series[0].AddXY(tAge,tDM);
      Chart2DM.Series[0].AddXY(tAge,tAge);
      if ((i mod iLabelModelAgeIncrement) = 0) then
      begin
        ChartAgeHfInitial.Series[1].AddXY(tAge,tDM);
        ChartAgeHfEpsilon.Series[1].AddXY(tAge,tEpsDM);
        //ChartAgeHfInitial.Series[1].AddXY(tAge,tDM);
        Chart2DM.Series[1].AddXY(tAge,tAge);
      end;
      i := i + iModelAgeIncrement;
    until (i >= iMaxModelAge);
  end;
end;

procedure TfmDVRDMain.RefreshConcordiaGraphs(AgeFrom,AgeTo : double);
var
  i : integer;
  iMaxConcordiaAge, iConcordiaAgeIncrement, iLabelConcordiaAgeIncrement : integer;
  t207235, t206238, t207206, t238206 : double;
  tAge : double;
  iCode : integer;
  MinimumUncertainty,
  ConcordanceFrom, ConcordanceTo,
  DiscordanceFrom, DiscordanceTo : double;
begin
  iMaxConcordiaAge := 4500;
  iConcordiaAgeIncrement := 100;
  iLabelConcordiaAgeIncrement := 400;
  ConcordanceFrom := 90.0;
  ConcordanceTo := 110.0;
  tAge := AgeTo;
  iMaxConcordiaAge := Trunc(tAge);
  ChartWetherill.Series[0].Clear;
  ChartWetherill.Series[1].Clear;
  ChartWetherill.Series[2].Clear;
  ChartWetherill.Series[3].Clear;
  ChartTeraWasserburg.Series[0].Clear;
  ChartTeraWasserburg.Series[1].Clear;
  ChartTeraWasserburg.Series[2].Clear;
  ChartTeraWasserburg.Series[3].Clear;
  i := 0;
  tAge := AgeFrom;
  i := Trunc(tAge);
  repeat
    dmDVRD.CalculateConcordiaForAge(1.0e6*i,t207235,t206238,t207206,t238206);
    ChartWetherill.Series[0].AddXY(t207235,t206238);
    if (i > 0) then ChartTeraWasserburg.Series[0].AddXY(t238206,t207206);
    if ((i mod iLabelConcordiaAgeIncrement) = 0) then
    begin
      ChartWetherill.Series[1].AddXY(t207235,t206238);
      if (i > 0) then ChartTeraWasserburg.Series[1].AddXY(t238206,t207206);
    end;
    i := i + iConcordiaAgeIncrement;
  until (i >= iMaxConcordiaAge);
  DiscordanceFrom := 100.0 - ConcordanceTo;
  DiscordanceTo := 100.0 - ConcordanceFrom;
  GetConcordiaSampleData(ConcordanceFrom,ConcordanceTo);
end;

procedure TfmDVRDMain.GetConcordiaSampleData(ConcordanceFrom,ConcordanceTo : double);
var
  iCnt, iCntIncluded, i : integer;
  tIncludeYN : string;
  iCode : integer;
  tConcordance : double;
  iErrTypX, iErrTypY, iErrTypZ : integer;
  tErrValX, tDataValX,
  tErrValY, tDataValY,
  tErrValZ, tDataValZ : double;
begin
  tIncludeYN := 'Y';
  iCnt := 0;
  iCntIncluded := 0;
  //GetListBoxValues(iwlSampleZones,dmDV.cdsSampleZones,'ZoneType','ZoneID',UserSession.SampleZoneValues);
  //if ((UserSession.IncludeSampleZoneValues) and (UserSession.SampleZoneValues.Count < 1)) then UserSession.IncludeSampleZoneValues := false;
  dmDVRD.ConstructRawDataSampleQuery;
  dmDVRD.qRawSmp.Close;
  dmDVRD.cdsRawSmp.Close;
  dmDVRD.qRawSmp.ParamByName('RegAssocID').AsString := 'UPb';
  dmDVRD.cdsRawSmp.Open;
  dmDVRD.cdsData.Open;
  try
    dmDVRD.cdsData.EmptyDataSet;
  except
  end;
  dmDVRD.cdsRawSmp.First;
  i := 1;
  repeat
    tConcordance := (100.0 - dmDVRD.cdsRawDiscordanceDATAVALUE.AsFloat);
    if ((tConcordance >= (ConcordanceFrom)) and (tConcordance <= (ConcordanceTo))) then tIncludeYN := 'Y'
                                                                               else tIncludeYN := 'N';
    iErrTypX := Trunc(dmDVRD.cdsRawErrTypeXDATAVALUE.AsFloat);
    tErrValX := dmDVRD.cdsRawDataXerrDATAVALUE.AsFloat;
    tDataValX := dmDVRD.cdsRawSmpqRawDataX.AsFloat;
    iErrTypY := Trunc(dmDVRD.cdsRawErrTypeYDATAVALUE.AsFloat);
    tErrValY := dmDVRD.cdsRawDataYerrDATAVALUE.AsFloat;
    tDataValY := dmDVRD.cdsRawSmpqRawDataY.AsFloat;
    iErrTypZ := Trunc(dmDVRD.cdsRawErrTypeZDATAVALUE.AsFloat);
    tErrValZ := dmDVRD.cdsRawDataZerrDATAVALUE.AsFloat;
    tDataValZ := dmDVRD.cdsRawSmpqRawDataZ.AsFloat;
      if ((tDataValX < 0.00001) or (tDataValY < 0.00001) or (tDataValZ < 0.00001)) then
      begin
        tIncludeYN := 'N';
        tConcordance := 0.0;
      end;
    if (dmDVRD.cdsRawDataIncludeDATAVALUE.AsFloat < 0.5) then
    begin
      tIncludeYN := 'N';
    end;
    if (iErrTypX = 0) then tErrValX := tErrValX*tDataValX/100.0;     //% uncertainties
    if (iErrTypY = 0) then tErrValY := tErrValY*tDataValY/100.0;     //
    if (iErrTypZ = 0) then tErrValZ := tErrValZ*tDataValZ/100.0;     //
    try
      dmDVRD.cdsData.Append;
      dmDVRD.cdsDatatRec.AsInteger := i;
      dmDVRD.cdsDataSampleNo.AsString := dmDVRD.cdsRawSmpSAMPLENO.AsString;
      dmDVRD.cdsDataFrac.AsString := dmDVRD.cdsRawSmpFRAC.AsString;
      dmDVRD.cdsDataZoneID.AsString := dmDVRD.cdsRawSmpZONEID.AsString;
      dmDVRD.cdsDataPb207U235.AsFloat := tDataValX;
      dmDVRD.cdsDataPb207U235Sigma.AsFloat := tErrValX;
      dmDVRD.cdsDataPb206U238.AsFloat := tDataValY;
      dmDVRD.cdsDataPb206U238Sigma.AsFloat := tErrValY;
      if (tDataValY > 0.0) then
      begin
        dmDVRD.cdsDataU238Pb206.AsFloat := 1.0/tDataValY;
        dmDVRD.cdsDataU238Pb206Sigma.AsFloat := (1.0/tDataValY)*tErrValY/tDataValY;
      end else
      begin
        dmDVRD.cdsDataU238Pb206.AsFloat := 0.0;
        dmDVRD.cdsDataU238Pb206Sigma.AsFloat := 1000.0;
        tIncludeYN := 'N';
      end;
      dmDVRD.cdsDataPb207Pb206.AsFloat := tDataValZ;
      dmDVRD.cdsDataPb207Pb206Sigma.AsFloat := tErrValZ;
      dmDVRD.cdsDataIncludeYN.AsString := tIncludeYN;
      dmDVRD.cdsDataPercentConcordance.AsFloat := tConcordance;
      dmDVRD.cdsDataPreferredAge.AsFloat := dmDVRD.cdsRawAgePrefDATAVALUE.AsFloat;
      dmDVRD.cdsDataPreferredAgeSigma.AsFloat := dmDVRD.cdsRawDataAgeerrDATAVALUE.AsFloat;
      dmDVRD.cdsData.Post;
    except
    end;
    i := i + 1;
    if (tIncludeYN = 'Y') then iCntIncluded := iCntIncluded + 1;
    iCnt := iCnt + 1;
    dmDVRD.cdsRawSmp.Next;
  until dmDVRD.cdsRawSmp.Eof;
  dmDVRD.cdsData.First;
  dmDVRD.cdsRawSmp.Close;
  PopulateConcordiaGraphs(ConcordanceFrom,ConcordanceTo);
end;

procedure TfmDVRDMain.PopulateConcordiaGraphs(ConcordanceFrom,ConcordanceTo : double);
var
  iCode : integer;
  i : integer;
  t207235, t206238, t207206, t238206 : double;
  tIncludeYN : string;
  tConcordance : double;
begin
  dmDVRD.cdsData.First;
  ChartWetherill.Series[2].Clear;
  ChartWetherill.Series[3].Clear;
  ChartTeraWasserburg.Series[2].Clear;
  ChartTeraWasserburg.Series[3].Clear;
  i := 0;
  repeat
    tIncludeYN := dmDVRD.cdsDataIncludeYN.AsString;
    tConcordance := dmDVRD.cdsDataPercentConcordance.AsFloat;
    t207235 := dmDVRD.cdsDataPb207U235.AsFloat;
    t206238 := dmDVRD.cdsDataPb206U238.AsFloat;
    t238206 := dmDVRD.cdsDataU238Pb206.AsFloat;
    t207206 := dmDVRD.cdsDataPb207Pb206.AsFloat;
    if (tIncludeYN = 'Y') then
    begin
      if ((t207235 > 0.0) and (t206238 > 0.0)) then
      begin
        ChartWetherill.Series[2].AddXY(t207235,t206238);
      end;
      if ((t238206 > 0.0) and (t207206 > 0.0)) then
      begin
        ChartTeraWasserburg.Series[2].AddXY(t238206,t207206);
      end;
    end else
    begin
      if ((t207235 > 0.0) and (t206238 > 0.0)) then
      begin
        ChartWetherill.Series[3].AddXY(t207235,t206238);
      end;
      if ((t238206 > 0.0) and (t207206 > 0.0)) then
      begin
        ChartTeraWasserburg.Series[3].AddXY(t238206,t207206);
      end;
    end;
    dmDVRD.cdsData.Next;
  until dmDVRD.cdsData.Eof;
  dmDVRD.cdsData.First;
end;


end.

