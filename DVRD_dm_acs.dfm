�
 TDMDVRD 0  TPF0TdmDVRDdmDVRDOldCreateOrder	Height�WidthO TADOConnectionDateViewRawDataConnectionString|Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Data\DateView contributions\Hf and O\DB_Hf_O.mdb;Persist Security Info=FalseKeepConnectionLoginPromptModecmShareDenyNoneProviderMicrosoft.Jet.OLEDB.4.0Left$Top  TDataSourcedsVariablesDataSet	VariablesLeft8TopP  	TADOQueryImportGroup
ConnectionDateViewRawData
CursorTypectStatic
Parameters SQL.Strings-select distinct ImportSpecName from Variables LeftTop�  TWideStringFieldImportGroupImportSpecName	FieldNameImportSpecName   TDataSourcedsImportGroupDataSetImportGroupLeft6Top�   	TADOQuerySmpData
ConnectionDateViewRawData
CursorTypectStatic
Parameters SQL.Stringsselect * from SmpData LeftTop�  TWideStringFieldSmpDataSampleNo	FieldNameSampleNoProviderFlags
pfInUpdate	pfInWherepfInKey   TWideStringFieldSmpDataFrac	FieldNameFracProviderFlags
pfInUpdate	pfInWherepfInKey Size  TWideStringFieldSmpDataIsoSystem	FieldName	IsoSystemProviderFlags
pfInUpdate	pfInWherepfInKey Size  TWideStringFieldSmpDataVariableID	FieldName
VariableIDProviderFlags
pfInUpdate	pfInWherepfInKey Size  TFloatFieldSmpDataDataValue	FieldName	DataValue  TWideStringFieldSmpDataNormalisingStandard	FieldNameNormalisingStandardSize
  TFloatFieldSmpDataStandardValue	FieldNameStandardValue  TFloatFieldSmpDataNormalisingFactor	FieldNameNormalisingFactor  TIntegerFieldSmpDataRefNum	FieldNameRefNum   TDataSource	dsSmpDataDataSetSmpDataLeftlTop�   	TADOQuery	ElemNames
ConnectionDateViewRawData
CursorTypectStatic
ParametersNameImportSpecName
Attributes
paNullable DataTypeftWideStringNumericScale� 	Precision� Size�Value   SQL.StringsSELECT * FROM Variables.WHERE VARIABLES.ImportSpecName=:ImportSpecNameORDER BY VARIABLES.Pos LeftTop  TWideStringFieldElemNamesImportSpecName	FieldNameImportSpecNameProviderFlags
pfInUpdate	pfInWherepfInKey   TSmallintFieldElemNamesPos	FieldNamePosProviderFlags
pfInUpdate	pfInWherepfInKey   TWideStringFieldElemNamesVariableID	FieldName
VariableIDSize  TWideStringFieldElemNamesColumnLetter	FieldNameColumnLetterSize  TSmallintFieldElemNamesColumnNo	FieldNameColumnNo  TWideStringFieldElemNamesIsoSystem	FieldName	IsoSystemSize  TWideStringFieldElemNamesNormalisingStandard	FieldNameNormalisingStandardSize
  TFloatFieldElemNamesStandardValue	FieldNameStandardValue  TFloatFieldElemNamesNormalisingFactor	FieldNameNormalisingFactor   TDataSourcedsElemNamesDataSet	ElemNamesLeft2Top   	TADOTable	Variables
ConnectionDateViewRawData
CursorTypectStatic	TableName	VariablesLeftTopP TWideStringFieldVariablesImportSpecName	FieldNameImportSpecNameProviderFlags
pfInUpdate	pfInWherepfInKey   TSmallintFieldVariablesPos	FieldNamePosProviderFlags
pfInUpdate	pfInWherepfInKey   TWideStringFieldVariablesVariableID	FieldName
VariableIDSize  TWideStringFieldVariablesColumnLetter	FieldNameColumnLetterSize  TSmallintFieldVariablesColumnNo	FieldNameColumnNo  TWideStringFieldVariablesIsoSystem	FieldName	IsoSystemSize  TWideStringFieldVariablesNormalisingStandard	FieldNameNormalisingStandardSize
  TFloatFieldVariablesStandardValue	FieldNameStandardValue  TFloatFieldVariablesNormalisingFactor	FieldNameNormalisingFactor    