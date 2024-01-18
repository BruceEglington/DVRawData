unit DVRD_dm_flex;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, ADODB, DBClient, Provider, UExcelAdapter, XLSAdapter,
  UFlxMemTable, XlsBaseTemplateStore, UXlsDBTemplateStore,
  UCustomFlexCelReport, UFlexCelReport;

type
  TdmFlex = class(TDataModule)
    CoranChem: TADOQuery;
    CoranChemGROUPNAME: TWideStringField;
    CoranChemPlotGroupName: TWideStringField;
    CoranChemSAMPLENUM: TWideStringField;
    CoranChemParam1: TFloatField;
    CoranChemParam2: TFloatField;
    CoranChemParam3: TFloatField;
    CoranChemParam4: TFloatField;
    CoranChemParam5: TFloatField;
    CoranChemParam6: TFloatField;
    CoranChemParam7: TFloatField;
    CoranChemParam8: TFloatField;
    CoranChemParam9: TFloatField;
    CoranChemParam10: TFloatField;
    CoranChemParam11: TFloatField;
    CoranChemParam12: TFloatField;
    CoranChemParam13: TFloatField;
    CoranChemParam14: TFloatField;
    CoranChemParam15: TFloatField;
    CoranChemParam16: TFloatField;
    CoranChemParam17: TFloatField;
    CoranChemParam18: TFloatField;
    CoranChemParam19: TFloatField;
    CoranChemParam20: TFloatField;
    CoranChemParam21: TFloatField;
    CoranChemParam22: TFloatField;
    CoranChemParam23: TFloatField;
    CoranChemParam24: TFloatField;
    CoranChemParam25: TFloatField;
    CoranChemParam26: TFloatField;
    CoranChemParam27: TFloatField;
    CoranChemParam28: TFloatField;
    CoranChemParam29: TFloatField;
    CoranChemParam30: TFloatField;
    CoranChemParam31: TFloatField;
    CoranChemParam32: TFloatField;
    CoranChemParam33: TFloatField;
    CoranChemParam34: TFloatField;
    CoranChemParam35: TFloatField;
    CoranChemParam36: TFloatField;
    CoranChemParam37: TFloatField;
    CoranChemParam38: TFloatField;
    CoranChemparam39: TFloatField;
    CoranChemParam40: TFloatField;
    CoranChemParam41: TFloatField;
    CoranChemParam42: TFloatField;
    CoranChemParam43: TFloatField;
    CoranChemParam44: TFloatField;
    CoranChemParam45: TFloatField;
    CoranChemParam46: TFloatField;
    CoranChemParam47: TFloatField;
    CoranChemParam48: TFloatField;
    CoranChemParam49: TFloatField;
    CoranChemParam50: TFloatField;
    CoranChemParam51: TFloatField;
    CoranChemParam52: TFloatField;
    CoranChemParam53: TFloatField;
    CoranChemParam54: TFloatField;
    CoranChemParam55: TFloatField;
    CoranChemParam56: TFloatField;
    CoranChemParam57: TFloatField;
    CoranChemParam58: TFloatField;
    CoranChemParam59: TFloatField;
    CoranChemParam60: TFloatField;
    CoranChemParam61: TFloatField;
    CoranChemParam62: TFloatField;
    CoranChemParam63: TFloatField;
    CoranChemParam64: TFloatField;
    CoranChemparam65: TFloatField;
    CoranChemParam66: TFloatField;
    CoranChemParam67: TFloatField;
    CoranChemParam68: TFloatField;
    CoranChemParam69: TFloatField;
    CoranChemParam70: TFloatField;
    CoranChemSequence: TIntegerField;
    FlexCoranChem: TFlexCelReport;
    XLSAdapter1: TXLSAdapter;
    SaveDialogFlex: TSaveDialog;
    procedure SmpDataPostError(DataSet: TDataSet; E: EDatabaseError;
      var Action: TDataAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  dmFlex: TdmFlex;

implementation

uses DVRD_dm;

{$R *.DFM}

procedure TdmFlex.SmpDataPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
begin
  MessageDlg('Key violation - duplicate combination of Coran type, Group and Sample',
            mtWarning,[mbOK],0);
  dmDVRD.cdsSmpData.CancelUpdates;
end;


end.
