program DVRawData;

uses
  Forms,
  DVRD_mn in 'DVRD_mn.pas' {fmDVRDMain},
  DVRD_About in 'DVRD_About.pas' {AboutBox},
  DVRD_dm in 'DVRD_dm.pas' {Empty: TDataModule},
  DVRD_varb in 'DVRD_varb.pas',
  DVRD_ShtIm2 in 'DVRD_ShtIm2.pas' {fmSheetImport2},
  DVRD_ShtIm in 'DVRD_ShtIm.pas' {fmSheetImport},
  Vcl.Themes,
  Vcl.Styles,
  Allsorts in '..\Eglington Delphi common code items\Allsorts.pas';

{$R *.res}

begin
  Application.Initialize;
  TStyleManager.TrySetStyle('Iceberg Classico');
  Application.Title := 'DateView Raw Data Importer';
  Application.CreateForm(TdmDVRD, dmDVRD);
  Application.CreateForm(TfmDVRDMain, fmDVRDMain);
  Application.Run;
end.

