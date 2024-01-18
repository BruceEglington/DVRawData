unit DVRD_dm_acs;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, ADODB, DBClient, Provider, UExcelAdapter, XLSAdapter,
  UFlxMemTable;

type
  TdmDVRD = class(TDataModule)
    DateViewRawData: TADOConnection;
    dsVariables: TDataSource;
    ImportGroup: TADOQuery;
    dsImportGroup: TDataSource;
    ImportGroupImportSpecName: TWideStringField;
    SmpData: TADOQuery;
    dsSmpData: TDataSource;
    SmpDataSampleNo: TWideStringField;
    SmpDataFrac: TWideStringField;
    SmpDataIsoSystem: TWideStringField;
    SmpDataVariableID: TWideStringField;
    SmpDataDataValue: TFloatField;
    SmpDataNormalisingStandard: TWideStringField;
    SmpDataStandardValue: TFloatField;
    SmpDataNormalisingFactor: TFloatField;
    ElemNames: TADOQuery;
    dsElemNames: TDataSource;
    Variables: TADOTable;
    VariablesImportSpecName: TWideStringField;
    VariablesVariableID: TWideStringField;
    VariablesColumnLetter: TWideStringField;
    VariablesColumnNo: TSmallintField;
    VariablesIsoSystem: TWideStringField;
    VariablesNormalisingStandard: TWideStringField;
    VariablesStandardValue: TFloatField;
    VariablesNormalisingFactor: TFloatField;
    ElemNamesImportSpecName: TWideStringField;
    ElemNamesVariableID: TWideStringField;
    ElemNamesColumnLetter: TWideStringField;
    ElemNamesColumnNo: TSmallintField;
    ElemNamesIsoSystem: TWideStringField;
    ElemNamesNormalisingStandard: TWideStringField;
    ElemNamesStandardValue: TFloatField;
    ElemNamesNormalisingFactor: TFloatField;
    VariablesPos: TSmallintField;
    ElemNamesPos: TSmallintField;
    SmpDataRefNum: TIntegerField;
    procedure VariablesPostError(DataSet: TDataSet; E: EDatabaseError;
      var Action: TDataAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  dmDVRD: TdmDVRD;

implementation

{$R *.DFM}

procedure TdmDVRD.VariablesPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
begin
  MessageDlg('Key violation - duplicate combination of ImportSpec and Position',
            mtWarning,[mbOK],0);
  dmDVRD.Variables.CancelUpdates;
end;

end.
