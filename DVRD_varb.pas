unit DVRD_varb;

interface

uses
  SysUtils, AllSorts;

const
  NN          = 1000;
  MM          = 30;
  MMax        = 10;
  DVRDVersion = '2.2.0';
  zero         = 0;
  DefaultZeroLimit = 1.0e-14;
  Cutoff = 0.975;
  Minimum207Pb235U : double = 0.0;
  Maximum207Pb235U : double = 85.0;
  Minimum206Pb238U : double = 0.0;
  Maximum206Pb238U : double = 1.0;
  Minimum207Pb206Pb : double = 0.0;
  Maximum207Pb206Pb : double = 0.61;
  Minimumr207Pb206Pb : double = -1.0;
  Maximumr207Pb206Pb : double = 1.0;
  Minimum238U206Pb : double = 0.0;
  Maximum238U206Pb : double = 55.0;

type
  RealArrayM  = array[1..MM] of double;
  IntArray    = array[1..MM] of integer;
  String15Array100  = array[1..NN] of string[15];

var
  Component   : String15Array100;
  Filename    : string[8];
  Title       : string[40];
  tempstr     : string[10];
  //N, M        : integer;
  Total       : double;
  DefaultMinimum : double;
  GlobalChosenStyle : string;

var
   done                  : boolean;
   Lst                   : TextFile;
   AnyKey                : char;
   Toggle100             : byte;
   FilePrepared          : boolean;
   ElementPos            : array [1..MM] of integer;
   FullFileName         : string;
   CommonFilePath,
   FlexTemplatePath,
   ExportPath, DBPath,
   DBVendorLib,
   ADODataLinkFile, DataPath   : string;
   TotalRecs                   : Integer;
   RowCount             : array[1..10] of integer;
   ImportSheetNumber,
   ImportSpecNameCol,
   PositionCol, CalledCol,
   NormStandardCol, StandardValueCol,
   NormFactorCol, IsoSystemCol,
   ZoneIDCol,
   TechAbrCol,
   MaterialAbrCol,
   ColumnCol            : integer;
   ImportSpecNameColStr,
   PositionColStr, CalledColStr,
   NormStandardColStr, StandardValueColStr,
   NormFactorColStr, IsoSystemColStr,
   MaterialAbrColStr,
   ColumnColStr         : string;
   FromRowValueString, ToRowValueString : string;
   FromRow, ToRow : integer;
   Nox : integer;
   UseDefaultTechnique, UseDefaultNormalisingStandard : boolean;
   IniFileName : string;

  function ConvertCol2Int(AnyString : string) : integer;

implementation


function ConvertCol2Int(AnyString : string) : integer;
var
  itmp    : integer;
  tmpStr  : string;
  tmpChar : char;
begin
    AnyString := UpperCase(AnyString);
    tmpStr := AnyString;
    ClearNull(tmpStr);
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


end.