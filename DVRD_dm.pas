unit DVRD_dm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBXCommon,Db, DBClient, Provider,
  WideStrings, FMTBcd, SqlExpr, MidasLib, DbxDevartInterbase,
  Vcl.BaseImageCollection, Vcl.ImageCollection, SVGIconImageCollection;

type
  TdmDVRD = class(TDataModule)
    DateViewRawData: TSQLConnection;
    qImportSpecVariables: TSQLQuery;
    dsImportSpecVariables: TDataSource;
    dspImportSpecVariables: TDataSetProvider;
    cdsImportSpecVariables: TClientDataSet;
    cdsImportSpecVariablesIMPORTSPECNAME: TStringField;
    cdsImportSpecVariablesPOS: TIntegerField;
    cdsImportSpecVariablesVARIABLEID: TStringField;
    cdsImportSpecVariablesCOLUMNLETTER: TStringField;
    cdsImportSpecVariablesCOLUMNNO: TIntegerField;
    cdsImportSpecVariablesISOSYSTEM: TStringField;
    cdsImportSpecVariablesNORMALISINGSTANDARD: TStringField;
    cdsImportSpecVariablesSTANDARDVALUE: TFloatField;
    cdsImportSpecVariablesNORMALISINGFACTOR: TFloatField;
    qImportGroup: TSQLQuery;
    dspImportGroup: TDataSetProvider;
    cdsImportGroup: TClientDataSet;
    dsImportGroup: TDataSource;
    qSmpData: TSQLQuery;
    dspSmpData: TDataSetProvider;
    cdsSmpData: TClientDataSet;
    dsSmpData: TDataSource;
    qElemNames: TSQLQuery;
    dspElemNames: TDataSetProvider;
    cdsElemNames: TClientDataSet;
    dsElemNames: TDataSource;
    InsertImportSpecVariables: TSQLQuery;
    InsertSmpData: TSQLQuery;
    qDataRefs: TSQLQuery;
    dspDataRefs: TDataSetProvider;
    cdsDataRefs: TClientDataSet;
    dsDataRefs: TDataSource;
    cdsDataRefsREFNUM: TIntegerField;
    qDeleteSmpData: TSQLQuery;
    cdsSmpDataSAMPLENO: TStringField;
    cdsSmpDataFRAC: TStringField;
    cdsSmpDataISOSYSTEM: TStringField;
    cdsSmpDataVARIABLEID: TStringField;
    cdsSmpDataTECHABR: TStringField;
    cdsSmpDataMATERIALABR: TStringField;
    cdsSmpDataDATAVALUE: TFloatField;
    cdsSmpDataNORMALISINGSTANDARD: TStringField;
    cdsSmpDataSTANDARDVALUE: TFloatField;
    cdsSmpDataNORMALISINGFACTOR: TFloatField;
    cdsSmpDataREFNUM: TIntegerField;
    cdsElemNamesIMPORTSPECNAME: TStringField;
    cdsElemNamesPOS: TIntegerField;
    cdsElemNamesVARIABLEID: TStringField;
    cdsElemNamesCOLUMNLETTER: TStringField;
    cdsElemNamesCOLUMNNO: TIntegerField;
    cdsElemNamesISOSYSTEM: TStringField;
    cdsElemNamesNORMALISINGSTANDARD: TStringField;
    cdsElemNamesSTANDARDVALUE: TFloatField;
    cdsElemNamesNORMALISINGFACTOR: TFloatField;
    cdsImportGroupIMPORTSPECNAME: TStringField;
    InsertSmpList: TSQLQuery;
    InsertSmpFrac: TSQLQuery;
    DeleteSmpList: TSQLQuery;
    DeleteSmpFrac: TSQLQuery;
    qCheckVariables: TSQLQuery;
    cdsSmpDataZONEID: TStringField;
    SQLMonitor1: TSQLMonitor;
    qSamples: TSQLQuery;
    dspSamples: TDataSetProvider;
    cdsSamples: TClientDataSet;
    dsSamples: TDataSource;
    qIsotopeSystems: TSQLQuery;
    dspIsotopeSystems: TDataSetProvider;
    cdsIsotopeSystems: TClientDataSet;
    dsIsotopeSystems: TDataSource;
    cdsSamplesSAMPLENO: TStringField;
    cdsIsotopeSystemsISOSYSTEM: TStringField;
    qRawSmp: TSQLQuery;
    qRawSmpSAMPLENO: TStringField;
    qRawSmpFRAC: TStringField;
    qRawSmpREGASSOCID: TStringField;
    qRawSmpZONEID: TStringField;
    dsqRawSmp: TDataSource;
    dspRawSmp: TDataSetProvider;
    cdsRawSmp: TClientDataSet;
    cdsRawSmpSAMPLENO: TStringField;
    cdsRawSmpFRAC: TStringField;
    cdsRawSmpZONEID: TStringField;
    cdsRawSmpREGASSOCID: TStringField;
    cdsRawSmpqRawDataXerr: TDataSetField;
    cdsRawSmpqRawDataX: TDataSetField;
    cdsRawSmpqRawDataR: TDataSetField;
    cdsRawSmpqRawErrTypeX: TDataSetField;
    cdsRawSmpqRawErrTypeY: TDataSetField;
    cdsRawSmpqRawErrTypeZ: TDataSetField;
    cdsRawSmpqRawDataPrecZ: TDataSetField;
    cdsRawSmpqRawDataZerr: TDataSetField;
    cdsRawSmpqRawDataZ: TDataSetField;
    cdsRawSmpqRawDataPrecY: TDataSetField;
    cdsRawSmpqRawDataYerr: TDataSetField;
    cdsRawSmpqRawDataY: TDataSetField;
    cdsRawSmpqRawDataPrecX: TDataSetField;
    cdsRawSmpqRawAgePref: TDataSetField;
    cdsRawSmpqRawDiscordance: TDataSetField;
    cdsRawSmpqRawDataDM: TDataSetField;
    cdsRawSmpqRawDataInclude: TDataSetField;
    dsRawSmp: TDataSource;
    qRawDataX: TSQLQuery;
    qRawDataXSAMPLENO: TStringField;
    qRawDataXFRAC: TStringField;
    qRawDataXISOSYSTEM: TStringField;
    qRawDataXVARIABLEID: TStringField;
    qRawDataXDATAVALUE: TFloatField;
    qRawDataXNORMALISINGSTANDARD: TStringField;
    qRawDataXSTANDARDVALUE: TFloatField;
    qRawDataXVARIABLENAME: TStringField;
    qRawDataXREGASSOCID: TStringField;
    dsqRawDataX: TDataSource;
    cdsRawDataX: TClientDataSet;
    cdsRawDataXSAMPLENO: TStringField;
    cdsRawDataXFRAC: TStringField;
    cdsRawDataXISOSYSTEM: TStringField;
    cdsRawDataXVARIABLEID: TStringField;
    cdsRawDataXNORMALISINGSTANDARD: TStringField;
    cdsRawDataXSTANDARDVALUE: TFloatField;
    cdsRawDataXVARIABLENAME: TStringField;
    cdsRawDataXREGASSOCID: TStringField;
    dsRawDataX: TDataSource;
    qRawDataXerr: TSQLQuery;
    qRawDataXerrSAMPLENO: TStringField;
    qRawDataXerrFRAC: TStringField;
    qRawDataXerrISOSYSTEM: TStringField;
    qRawDataXerrVARIABLEID: TStringField;
    qRawDataXerrDATAVALUE: TFloatField;
    qRawDataXerrNORMALISINGSTANDARD: TStringField;
    qRawDataXerrSTANDARDVALUE: TFloatField;
    qRawDataXerrVARIABLENAME: TStringField;
    dsqRawDataXerr: TDataSource;
    cdsRawDataXerr: TClientDataSet;
    cdsRawDataXerrSAMPLENO: TStringField;
    cdsRawDataXerrFRAC: TStringField;
    cdsRawDataXerrISOSYSTEM: TStringField;
    cdsRawDataXerrVARIABLEID: TStringField;
    cdsRawDataXerrDATAVALUE: TFloatField;
    cdsRawDataXerrNORMALISINGSTANDARD: TStringField;
    cdsRawDataXerrSTANDARDVALUE: TFloatField;
    cdsRawDataXerrVARIABLENAME: TStringField;
    dsRawDataXerr: TDataSource;
    qRawErrTypeX: TSQLQuery;
    qRawErrTypeXSAMPLENO: TStringField;
    qRawErrTypeXFRAC: TStringField;
    qRawErrTypeXISOSYSTEM: TStringField;
    qRawErrTypeXVARIABLEID: TStringField;
    qRawErrTypeXDATAVALUE: TFloatField;
    qRawErrTypeXNORMALISINGSTANDARD: TStringField;
    qRawErrTypeXSTANDARDVALUE: TFloatField;
    qRawErrTypeXVARIABLENAME: TStringField;
    cdsRawErrTypeX: TClientDataSet;
    cdsRawErrTypeXSAMPLENO: TStringField;
    cdsRawErrTypeXFRAC: TStringField;
    cdsRawErrTypeXISOSYSTEM: TStringField;
    cdsRawErrTypeXVARIABLEID: TStringField;
    cdsRawErrTypeXDATAVALUE: TFloatField;
    cdsRawErrTypeXNORMALISINGSTANDARD: TStringField;
    cdsRawErrTypeXSTANDARDVALUE: TFloatField;
    cdsRawErrTypeXVARIABLENAME: TStringField;
    dsRawErrTypeX: TDataSource;
    qRawDataPrecX: TSQLQuery;
    qRawDataPrecXSAMPLENO: TStringField;
    qRawDataPrecXFRAC: TStringField;
    qRawDataPrecXISOSYSTEM: TStringField;
    qRawDataPrecXVARIABLEID: TStringField;
    qRawDataPrecXDATAVALUE: TFloatField;
    cdsRawDataPrecX: TClientDataSet;
    cdsRawDataPrecXSAMPLENO: TStringField;
    cdsRawDataPrecXFRAC: TStringField;
    cdsRawDataPrecXISOSYSTEM: TStringField;
    cdsRawDataPrecXVARIABLEID: TStringField;
    cdsRawDataPrecXDATAVALUE: TFloatField;
    dsRawDataPrecX: TDataSource;
    qRawDataY: TSQLQuery;
    qRawDataYSAMPLENO: TStringField;
    qRawDataYFRAC: TStringField;
    qRawDataYISOSYSTEM: TStringField;
    qRawDataYVARIABLEID: TStringField;
    qRawDataYDATAVALUE: TFloatField;
    qRawDataYNORMALISINGSTANDARD: TStringField;
    qRawDataYSTANDARDVALUE: TFloatField;
    qRawDataYVARIABLENAME: TStringField;
    qRawDataYREGASSOCID: TStringField;
    dsqRawDataY: TDataSource;
    cdsRawDataY: TClientDataSet;
    cdsRawDataYSAMPLENO: TStringField;
    cdsRawDataYFRAC: TStringField;
    cdsRawDataYISOSYSTEM: TStringField;
    cdsRawDataYVARIABLEID: TStringField;
    cdsRawDataYNORMALISINGSTANDARD: TStringField;
    cdsRawDataYSTANDARDVALUE: TFloatField;
    cdsRawDataYVARIABLENAME: TStringField;
    cdsRawDataYREGASSOCID: TStringField;
    dsRawDataY: TDataSource;
    qRawDataYerr: TSQLQuery;
    qRawDataYerrSAMPLENO: TStringField;
    qRawDataYerrFRAC: TStringField;
    qRawDataYerrISOSYSTEM: TStringField;
    qRawDataYerrVARIABLEID: TStringField;
    qRawDataYerrDATAVALUE: TFloatField;
    qRawDataYerrNORMALISINGSTANDARD: TStringField;
    qRawDataYerrSTANDARDVALUE: TFloatField;
    qRawDataYerrVARIABLENAME: TStringField;
    dsqRawDataYerr: TDataSource;
    cdsRawDataYerr: TClientDataSet;
    cdsRawDataYerrSAMPLENO: TStringField;
    cdsRawDataYerrFRAC: TStringField;
    cdsRawDataYerrISOSYSTEM: TStringField;
    cdsRawDataYerrVARIABLEID: TStringField;
    cdsRawDataYerrDATAVALUE: TFloatField;
    cdsRawDataYerrNORMALISINGSTANDARD: TStringField;
    cdsRawDataYerrSTANDARDVALUE: TFloatField;
    cdsRawDataYerrVARIABLENAME: TStringField;
    dsRawDataYerr: TDataSource;
    qRawErrTypeY: TSQLQuery;
    qRawErrTypeYSAMPLENO: TStringField;
    qRawErrTypeYFRAC: TStringField;
    qRawErrTypeYISOSYSTEM: TStringField;
    qRawErrTypeYVARIABLEID: TStringField;
    qRawErrTypeYDATAVALUE: TFloatField;
    qRawErrTypeYNORMALISINGSTANDARD: TStringField;
    qRawErrTypeYSTANDARDVALUE: TFloatField;
    qRawErrTypeYVARIABLENAME: TStringField;
    cdsRawErrTypeY: TClientDataSet;
    cdsRawErrTypeYSAMPLENO: TStringField;
    cdsRawErrTypeYFRAC: TStringField;
    cdsRawErrTypeYISOSYSTEM: TStringField;
    cdsRawErrTypeYVARIABLEID: TStringField;
    cdsRawErrTypeYDATAVALUE: TFloatField;
    cdsRawErrTypeYNORMALISINGSTANDARD: TStringField;
    cdsRawErrTypeYSTANDARDVALUE: TFloatField;
    cdsRawErrTypeYVARIABLENAME: TStringField;
    dsRawErrTypeY: TDataSource;
    qRawDataPrecY: TSQLQuery;
    qRawDataPrecYSAMPLENO: TStringField;
    qRawDataPrecYFRAC: TStringField;
    qRawDataPrecYISOSYSTEM: TStringField;
    qRawDataPrecYVARIABLEID: TStringField;
    qRawDataPrecYDATAVALUE: TFloatField;
    cdsRawDataPrecY: TClientDataSet;
    cdsRawDataPrecYSAMPLENO: TStringField;
    cdsRawDataPrecYFRAC: TStringField;
    cdsRawDataPrecYISOSYSTEM: TStringField;
    cdsRawDataPrecYVARIABLEID: TStringField;
    cdsRawDataPrecYDATAVALUE: TFloatField;
    dsRawDataPrecY: TDataSource;
    qRawDataR: TSQLQuery;
    qRawDataRSAMPLENO: TStringField;
    qRawDataRFRAC: TStringField;
    qRawDataRISOSYSTEM: TStringField;
    qRawDataRVARIABLEID: TStringField;
    qRawDataRDATAVALUE: TFloatField;
    qRawDataRNORMALISINGSTANDARD: TStringField;
    qRawDataRSTANDARDVALUE: TFloatField;
    qRawDataRVARIABLENAME: TStringField;
    cdsRawDataR: TClientDataSet;
    cdsRawDataRSAMPLENO: TStringField;
    cdsRawDataRFRAC: TStringField;
    cdsRawDataRISOSYSTEM: TStringField;
    cdsRawDataRVARIABLEID: TStringField;
    cdsRawDataRDATAVALUE: TFloatField;
    cdsRawDataRNORMALISINGSTANDARD: TStringField;
    cdsRawDataRSTANDARDVALUE: TFloatField;
    cdsRawDataRVARIABLENAME: TStringField;
    dsRawDataR: TDataSource;
    qRawDataZ: TSQLQuery;
    qRawDataZSAMPLENO: TStringField;
    qRawDataZFRAC: TStringField;
    qRawDataZISOSYSTEM: TStringField;
    qRawDataZVARIABLEID: TStringField;
    qRawDataZDATAVALUE: TFloatField;
    qRawDataZNORMALISINGSTANDARD: TStringField;
    qRawDataZSTANDARDVALUE: TFloatField;
    qRawDataZVARIABLENAME: TStringField;
    qRawDataZREGASSOCID: TStringField;
    dsqRawDataZ: TDataSource;
    cdsRawDataZ: TClientDataSet;
    cdsRawDataZSAMPLENO: TStringField;
    cdsRawDataZFRAC: TStringField;
    cdsRawDataZISOSYSTEM: TStringField;
    cdsRawDataZVARIABLEID: TStringField;
    cdsRawDataZNORMALISINGSTANDARD: TStringField;
    cdsRawDataZSTANDARDVALUE: TFloatField;
    cdsRawDataZVARIABLENAME: TStringField;
    cdsRawDataZREGASSOCID: TStringField;
    dsRawDataZ: TDataSource;
    qRawDataZerr: TSQLQuery;
    qRawDataZerrSAMPLENO: TStringField;
    qRawDataZerrFRAC: TStringField;
    qRawDataZerrISOSYSTEM: TStringField;
    qRawDataZerrVARIABLEID: TStringField;
    qRawDataZerrDATAVALUE: TFloatField;
    qRawDataZerrNORMALISINGSTANDARD: TStringField;
    qRawDataZerrSTANDARDVALUE: TFloatField;
    qRawDataZerrVARIABLENAME: TStringField;
    cdsRawDataZerr: TClientDataSet;
    cdsRawDataZerrSAMPLENO: TStringField;
    cdsRawDataZerrFRAC: TStringField;
    cdsRawDataZerrISOSYSTEM: TStringField;
    cdsRawDataZerrVARIABLEID: TStringField;
    cdsRawDataZerrDATAVALUE: TFloatField;
    cdsRawDataZerrNORMALISINGSTANDARD: TStringField;
    cdsRawDataZerrSTANDARDVALUE: TFloatField;
    cdsRawDataZerrVARIABLENAME: TStringField;
    dsRawDataZerr: TDataSource;
    qRawErrTypeZ: TSQLQuery;
    qRawErrTypeZSAMPLENO: TStringField;
    qRawErrTypeZFRAC: TStringField;
    qRawErrTypeZISOSYSTEM: TStringField;
    qRawErrTypeZVARIABLEID: TStringField;
    qRawErrTypeZDATAVALUE: TFloatField;
    qRawErrTypeZNORMALISINGSTANDARD: TStringField;
    qRawErrTypeZSTANDARDVALUE: TFloatField;
    qRawErrTypeZVARIABLENAME: TStringField;
    cdsRawErrTypeZ: TClientDataSet;
    cdsRawErrTypeZSAMPLENO: TStringField;
    cdsRawErrTypeZFRAC: TStringField;
    cdsRawErrTypeZISOSYSTEM: TStringField;
    cdsRawErrTypeZVARIABLEID: TStringField;
    cdsRawErrTypeZDATAVALUE: TFloatField;
    cdsRawErrTypeZNORMALISINGSTANDARD: TStringField;
    cdsRawErrTypeZSTANDARDVALUE: TFloatField;
    cdsRawErrTypeZVARIABLENAME: TStringField;
    dsRawErrTypeZ: TDataSource;
    qRawDataPrecZ: TSQLQuery;
    qRawDataPrecZSAMPLENO: TStringField;
    qRawDataPrecZFRAC: TStringField;
    qRawDataPrecZISOSYSTEM: TStringField;
    qRawDataPrecZVARIABLEID: TStringField;
    qRawDataPrecZDATAVALUE: TFloatField;
    cdsRawDataPrecZ: TClientDataSet;
    cdsRawDataPrecZSAMPLENO: TStringField;
    cdsRawDataPrecZFRAC: TStringField;
    cdsRawDataPrecZISOSYSTEM: TStringField;
    cdsRawDataPrecZVARIABLEID: TStringField;
    cdsRawDataPrecZDATAVALUE: TFloatField;
    dsRawDataPrecZ: TDataSource;
    qRawDataDM: TSQLQuery;
    qRawDataDMSAMPLENO: TStringField;
    qRawDataDMFRAC: TStringField;
    qRawDataDMISOSYSTEM: TStringField;
    qRawDataDMVARIABLEID: TStringField;
    qRawDataDMDATAVALUE: TFloatField;
    qRawDataDMREGASSOCID: TStringField;
    dsqRawDataDM: TDataSource;
    cdsRawDataDM: TClientDataSet;
    cdsRawDataDMSAMPLENO: TStringField;
    cdsRawDataDMFRAC: TStringField;
    cdsRawDataDMISOSYSTEM: TStringField;
    cdsRawDataDMVARIABLEID: TStringField;
    cdsRawDataDMDATAVALUE: TFloatField;
    cdsRawDataDMqRawDataDMerr: TDataSetField;
    cdsRawDataDMREGASSOCID: TStringField;
    dsRawDataDM: TDataSource;
    qRawDiscordance: TSQLQuery;
    qRawDiscordanceSAMPLENO: TStringField;
    qRawDiscordanceFRAC: TStringField;
    qRawDiscordanceISOSYSTEM: TStringField;
    qRawDiscordanceVARIABLEID: TStringField;
    qRawDiscordanceDATAVALUE: TFloatField;
    cdsRawDiscordance: TClientDataSet;
    cdsRawDiscordanceSAMPLENO: TStringField;
    cdsRawDiscordanceFRAC: TStringField;
    cdsRawDiscordanceISOSYSTEM: TStringField;
    cdsRawDiscordanceVARIABLEID: TStringField;
    cdsRawDiscordanceDATAVALUE: TFloatField;
    dsRawDiscordance: TDataSource;
    qRawDataDMerr: TSQLQuery;
    qRawDataDMerrSAMPLENO: TStringField;
    qRawDataDMerrFRAC: TStringField;
    qRawDataDMerrISOSYSTEM: TStringField;
    qRawDataDMerrVARIABLEID: TStringField;
    qRawDataDMerrDATAVALUE: TFloatField;
    qRawDataDMerrNORMALISINGSTANDARD: TStringField;
    qRawDataDMerrSTANDARDVALUE: TFloatField;
    qRawDataDMerrVARIABLENAME: TStringField;
    dsqRawDataDMerr: TDataSource;
    cdsRawDataDMerr: TClientDataSet;
    cdsRawDataDMerrSAMPLENO: TStringField;
    cdsRawDataDMerrFRAC: TStringField;
    cdsRawDataDMerrISOSYSTEM: TStringField;
    cdsRawDataDMerrVARIABLEID: TStringField;
    cdsRawDataDMerrDATAVALUE: TFloatField;
    cdsRawDataDMerrNORMALISINGSTANDARD: TStringField;
    cdsRawDataDMerrSTANDARDVALUE: TFloatField;
    cdsRawDataDMerrVARIABLENAME: TStringField;
    dsRawDataDMerr: TDataSource;
    qRawAgePref: TSQLQuery;
    qRawAgePrefSAMPLENO: TStringField;
    qRawAgePrefFRAC: TStringField;
    qRawAgePrefISOSYSTEM: TStringField;
    qRawAgePrefVARIABLEID: TStringField;
    qRawAgePrefDATAVALUE: TFloatField;
    qRawAgePrefREGASSOCID: TStringField;
    dsqRawAgePref: TDataSource;
    cdsRawAgePref: TClientDataSet;
    cdsRawAgePrefSAMPLENO: TStringField;
    cdsRawAgePrefFRAC: TStringField;
    cdsRawAgePrefISOSYSTEM: TStringField;
    cdsRawAgePrefVARIABLEID: TStringField;
    cdsRawAgePrefDATAVALUE: TFloatField;
    cdsRawAgePrefqRawDataAgeerr: TDataSetField;
    cdsRawAgePrefREGASSOCID: TStringField;
    dsRawAgePref: TDataSource;
    qRawDataAgeerr: TSQLQuery;
    qRawDataAgeerrSAMPLENO: TStringField;
    qRawDataAgeerrFRAC: TStringField;
    qRawDataAgeerrISOSYSTEM: TStringField;
    qRawDataAgeerrVARIABLEID: TStringField;
    qRawDataAgeerrDATAVALUE: TFloatField;
    qRawDataAgeerrNORMALISINGSTANDARD: TStringField;
    qRawDataAgeerrSTANDARDVALUE: TFloatField;
    qRawDataAgeerrVARIABLENAME: TStringField;
    dsqRawDataAgeerr: TDataSource;
    cdsRawDataAgeerr: TClientDataSet;
    cdsRawDataAgeerrSAMPLENO: TStringField;
    cdsRawDataAgeerrFRAC: TStringField;
    cdsRawDataAgeerrISOSYSTEM: TStringField;
    cdsRawDataAgeerrVARIABLEID: TStringField;
    cdsRawDataAgeerrDATAVALUE: TFloatField;
    cdsRawDataAgeerrNORMALISINGSTANDARD: TStringField;
    cdsRawDataAgeerrSTANDARDVALUE: TFloatField;
    cdsRawDataAgeerrVARIABLENAME: TStringField;
    dsRawDataAgeerr: TDataSource;
    qRawDataInclude: TSQLQuery;
    qRawDataIncludeSAMPLENO: TStringField;
    qRawDataIncludeFRAC: TStringField;
    qRawDataIncludeISOSYSTEM: TStringField;
    qRawDataIncludeVARIABLEID: TStringField;
    qRawDataIncludeDATAVALUE: TFloatField;
    cdsRawDataInclude: TClientDataSet;
    cdsRawDataIncludeSAMPLENO: TStringField;
    cdsRawDataIncludeFRAC: TStringField;
    cdsRawDataIncludeISOSYSTEM: TStringField;
    cdsRawDataIncludeVARIABLEID: TStringField;
    cdsRawDataIncludeDATAVALUE: TFloatField;
    dsRawDataInclude: TDataSource;
    cdsData: TClientDataSet;
    cdsDatatRec: TIntegerField;
    cdsDataSampleNo: TStringField;
    cdsDataFrac: TStringField;
    cdsDataZoneID: TStringField;
    cdsDataPb207U235: TFloatField;
    cdsDataPb207U235Sigma: TFloatField;
    cdsDataPb206U238: TFloatField;
    cdsDataPb206U238Sigma: TFloatField;
    cdsDataU238Pb206: TFloatField;
    cdsDataU238Pb206Sigma: TFloatField;
    cdsDataPb207Pb206: TFloatField;
    cdsDataPb207Pb206Sigma: TFloatField;
    cdsDataIncludeYN: TStringField;
    cdsDataPercentConcordance: TFloatField;
    cdsDataPreferredAge: TFloatField;
    cdsDataPreferredAgeSigma: TFloatField;
    dsData: TDataSource;
    qVarVar: TSQLQuery;
    dspVarVar: TDataSetProvider;
    cdsVarVar: TClientDataSet;
    dsVarVar: TDataSource;
    cdsVarVarVARIABLEID: TStringField;
    cdsVarVarVARIABLENAME: TStringField;
    cdsVarVarISOSYSTEM: TStringField;
    cdsRawDataXDATAVALUE: TFloatField;
    cdsRawDataYDATAVALUE: TFloatField;
    cdsRawDataZDATAVALUE: TFloatField;
    qRawDataInit: TSQLQuery;
    dsqRawdataInit: TDataSource;
    cdsRawDataInit: TClientDataSet;
    dsRawDataInit: TDataSource;
    qRawDataEps: TSQLQuery;
    dsqRawDataEps: TDataSource;
    cdsRawDataEps: TClientDataSet;
    dsRawDataEps: TDataSource;
    cdsRawSmpqRawDataEps: TDataSetField;
    cdsRawSmpqRawDataInit: TDataSetField;
    qRawDataInitSAMPLENO: TStringField;
    qRawDataInitFRAC: TStringField;
    qRawDataInitISOSYSTEM: TStringField;
    qRawDataInitVARIABLEID: TStringField;
    qRawDataInitDATAVALUE: TFloatField;
    qRawDataInitREGASSOCID: TStringField;
    qRawDataEpsSAMPLENO: TStringField;
    qRawDataEpsFRAC: TStringField;
    qRawDataEpsISOSYSTEM: TStringField;
    qRawDataEpsVARIABLEID: TStringField;
    qRawDataEpsDATAVALUE: TFloatField;
    qRawDataEpsREGASSOCID: TStringField;
    cdsRawDataInitSAMPLENO: TStringField;
    cdsRawDataInitFRAC: TStringField;
    cdsRawDataInitISOSYSTEM: TStringField;
    cdsRawDataInitVARIABLEID: TStringField;
    cdsRawDataInitDATAVALUE: TFloatField;
    cdsRawDataInitREGASSOCID: TStringField;
    cdsRawDataEpsSAMPLENO: TStringField;
    cdsRawDataEpsFRAC: TStringField;
    cdsRawDataEpsISOSYSTEM: TStringField;
    cdsRawDataEpsVARIABLEID: TStringField;
    cdsRawDataEpsDATAVALUE: TFloatField;
    cdsRawDataEpsREGASSOCID: TStringField;
    ImageCollection1: TImageCollection;
    SVGIconImageCollection1: TSVGIconImageCollection;
    procedure VariablesPostError(DataSet: TDataSet; E: EDatabaseError;
      var Action: TDataAction);
  private
    { Private declarations }
  public
    { Public declarations }
    ChosenStyle : string;
    procedure AddNewImportSpec(ImportSpecName : string;
                                     Pos : integer;
                                     VariableID : string;
                                     ColumnLetter : string;
                                     ColumnNo : integer;
                                     IsoSystem : string;
                                     NormalisingStandard : string;
                                     StandardValue : double;
                                     NormalisingFactor : double;
                                    var WasSuccessful : boolean);
    procedure AddNewSmpData(SampleNo : string;
                                Frac : string;
                                ZoneID : string;
                                IsoSystem : string;
                                VariableID : string;
                                DataValue : string;
                                NormalisingStandard : string;
                                StandardValue : double;
                                NormalisingFactor : double;
                                RefNum : string;
                                TechAbr : string;
                                MaterialAbr : string;
                            var WasSuccessful : boolean);
    procedure EmptySmpData(var WasSuccessful : boolean);
    procedure EmptySmpList(var WasSuccessful : boolean);
    procedure EmptySmpFrac(var WasSuccessful : boolean);
    procedure CopySmpList(var WasSuccessful : boolean);
    procedure CopySmpFrac(var WasSuccessful : boolean);
    procedure CalculateConcordiaForAge(Age : double; var t207235 : double;
                                  var t206238 : double; var t207206 : double;
                                  var t238206 : double);
    procedure ConstructRawDataSampleQuery;
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
  dmDVRD.cdsImportSpecVariables.CancelUpdates;
end;

procedure TdmDVRD.AddNewImportSpec(ImportSpecName : string;
                                     Pos : integer;
                                     VariableID : string;
                                     ColumnLetter : string;
                                     ColumnNo : integer;
                                     IsoSystem : string;
                                     NormalisingStandard : string;
                                     StandardValue : double;
                                     NormalisingFactor : double;
                                    var WasSuccessful : boolean);
var
  tmpCount : integer;
  tmpOffset : integer;
  tmpDomainID : integer;
  tmpXMin, tmpXMax : integer;
  tmpDomainHeading : string;
  TD: TDBXTransaction;
begin
  //ShowMessage(ImportSpecName+'  '+VariableID+'  '+ColumnLetter);
  WasSuccessful := true;
  try
    dmDVRD.InsertImportSpecvariables.Close;
    dmDVRD.InsertImportSpecvariables.ParamByName('IMPORTSPECNAME').AsString := ImportSpecName;
    dmDVRD.InsertImportSpecvariables.ParamByName('POS').AsInteger := Pos;
    dmDVRD.InsertImportSpecvariables.ParamByName('VARIABLEID').AsString := VariableID;
    dmDVRD.InsertImportSpecvariables.ParamByName('COLUMNLETTER').AsString := ColumnLetter;
    dmDVRD.InsertImportSpecvariables.ParamByName('COLUMNNO').AsInteger := ColumnNo;
    dmDVRD.InsertImportSpecvariables.ParamByName('ISOSYSTEM').AsString := IsoSystem;
    dmDVRD.InsertImportSpecvariables.ParamByName('NORMALISINGSTANDARD').AsString := NormalisingStandard;
    dmDVRD.InsertImportSpecvariables.ParamByName('STANDARDVALUE').AsFloat := StandardValue;
    dmDVRD.InsertImportSpecvariables.ParamByName('NORMALISINGFACTOR').AsFloat := NormalisingFactor;

    TD := dmDVRD.DateViewRawData.BeginTransaction(TDBXIsolations.ReadCommitted);
    try
      dmDVRD.InsertImportSpecvariables.ExecSQL(false);
      dmDVRD.DateViewRawData.CommitFreeAndNil(TD); //on success, commit the changes
      WasSuccessful := true;
    except
      dmDVRD.DateViewRawData.RollbackFreeAndNil(TD); //on failure, undo the changes
      WasSuccessful := false;
    end;
  except
  end;
end;

procedure TdmDVRD.AddNewSmpData(SampleNo : string;
                                Frac : string;
                                ZoneID : string;
                                IsoSystem : string;
                                VariableID : string;
                                DataValue : string;
                                NormalisingStandard : string;
                                StandardValue : double;
                                NormalisingFactor : double;
                                RefNum : string;
                                TechAbr : string;
                                MaterialAbr : string;
                            var WasSuccessful : boolean);
var
  tmpCount : integer;
  tmpOffset : integer;
  tmpDomainID : integer;
  tmpXMin, tmpXMax : integer;
  tmpDomainHeading : string;
  TD: TDBXTransaction;
begin
  WasSuccessful := true;
  //if (VariableID = 'U_ppm') then
  //  ShowMessage(SampleNo+'   '+Frac+'  '+IsoSystem+'  '+VariableID+'  '+TechAbr+'  '+MaterialAbr+'  '+DataValue+'  '+NormalisingStandard+'  '+FormatFloat('##0.00000',StandardValue)+'  '+FormatFloat('##0.00000',NormalisingFactor)+'  '+RefNum);
  try
    //dmDVRD.InsertSmpData.Close;
    dmDVRD.InsertSmpData.ParamByName('SAMPLENO').AsString := SampleNo;
    dmDVRD.InsertSmpData.ParamByName('FRAC').AsString := Frac;
    dmDVRD.InsertSmpData.ParamByName('ZONEID').AsString := ZoneID;
    dmDVRD.InsertSmpData.ParamByName('ISOSYSTEM').AsString := IsoSystem;
    dmDVRD.InsertSmpData.ParamByName('VARIABLEID').AsString := VariableID;
    dmDVRD.InsertSmpData.ParamByName('DATAVALUE').AsString := DataValue;
    dmDVRD.InsertSmpData.ParamByName('NORMALISINGSTANDARD').AsString := NormalisingStandard;
    dmDVRD.InsertSmpData.ParamByName('STANDARDVALUE').AsFloat := StandardValue;
    dmDVRD.InsertSmpData.ParamByName('NORMALISINGFACTOR').AsFloat := NormalisingFactor;
    dmDVRD.InsertSmpData.ParamByName('REFNUM').AsString := RefNum;
    dmDVRD.InsertSmpData.ParamByName('TECHABR').AsString := TechAbr;
    dmDVRD.InsertSmpData.ParamByName('MATERIALABR').AsString := MaterialAbr;

    TD := dmDVRD.DateViewRawData.BeginTransaction(TDBXIsolations.ReadCommitted);
    try
      dmDVRD.InsertSmpData.ExecSQL(false);
      dmDVRD.DateViewRawData.CommitFreeAndNil(TD); //on success, commit the changes
      WasSuccessful := true;
    except
      dmDVRD.DateViewRawData.RollbackFreeAndNil(TD); //on failure, undo the changes
      WasSuccessful := false;
    end;
  except
  end;
end;

procedure TdmDVRD.EmptySmpData(var WasSuccessful : boolean);
var
  TD: TDBXTransaction;
begin
  WasSuccessful := true;
  try
    TD := dmDVRD.DateViewRawData.BeginTransaction(TDBXIsolations.ReadCommitted);
    try
      dmDVRD.qDeleteSmpData.ExecSQL(false);
      dmDVRD.DateViewRawData.CommitFreeAndNil(TD); //on success, commit the changes
    except
      dmDVRD.DateViewRawData.RollbackFreeAndNil(TD); //on failure, undo the changes
      WasSuccessful := false;
    end;
  except
  end;
end;

procedure TdmDVRD.EmptySmpList(var WasSuccessful : boolean);
var
  TD: TDBXTransaction;
begin
  WasSuccessful := true;
  try
    TD := dmDVRD.DateViewRawData.BeginTransaction(TDBXIsolations.ReadCommitted);
    try
      dmDVRD.DeleteSmpList.ExecSQL(false);
      dmDVRD.DateViewRawData.CommitFreeAndNil(TD); //on success, commit the changes
    except
      dmDVRD.DateViewRawData.RollbackFreeAndNil(TD); //on failure, undo the changes
      WasSuccessful := false;
    end;
  except
  end;
end;

procedure TdmDVRD.EmptySmpFrac(var WasSuccessful : boolean);
var
  TD: TDBXTransaction;
begin
  WasSuccessful := true;
  try
    TD := dmDVRD.DateViewRawData.BeginTransaction(TDBXIsolations.ReadCommitted);
    try
      dmDVRD.DeleteSmpFrac.ExecSQL(false);
      dmDVRD.DateViewRawData.CommitFreeAndNil(TD); //on success, commit the changes
    except
      dmDVRD.DateViewRawData.RollbackFreeAndNil(TD); //on failure, undo the changes
      WasSuccessful := false;
    end;
  except
  end;
end;


procedure TdmDVRD.CopySmpList(var WasSuccessful : boolean);
var
  TD: TDBXTransaction;
begin
  WasSuccessful := true;
  try
    TD := dmDVRD.DateViewRawData.BeginTransaction(TDBXIsolations.ReadCommitted);
    try
      dmDVRD.InsertSmpList.ExecSQL(false);
      dmDVRD.DateViewRawData.CommitFreeAndNil(TD); //on success, commit the changes
    except
      dmDVRD.DateViewRawData.RollbackFreeAndNil(TD); //on failure, undo the changes
      WasSuccessful := false;
    end;
  except
  end;
end;

procedure TdmDVRD.CopySmpFrac(var WasSuccessful : boolean);
var
  TD: TDBXTransaction;
begin
  WasSuccessful := true;
  try
    TD := dmDVRD.DateViewRawData.BeginTransaction(TDBXIsolations.ReadCommitted);
    try
      dmDVRD.InsertSmpFrac.ExecSQL(false);
      dmDVRD.DateViewRawData.CommitFreeAndNil(TD); //on success, commit the changes
    except
      dmDVRD.DateViewRawData.RollbackFreeAndNil(TD); //on failure, undo the changes
      WasSuccessful := false;
    end;
  except
  end;
end;

procedure TdmDVRD.CalculateConcordiaForAge(Age : double; var t207235 : double;
                                  var t206238 : double; var t207206 : double;
                                  var t238206 : double);
const
  DecayConst235U = 9.8485e-10;
  DecayConst238U = 1.55125e-10;
  Constant238235 = 137.88;
begin
  t207235 := Exp(DecayConst235U*Age)-1.0;
  t206238 := Exp(DecayConst238U*Age)-1.0;
  if (Age > 0.00001) then t207206 := (1.0/Constant238235) * (t207235/t206238)
                     else t207206 := 0.046045;
  if (t206238 > 0.0) then  t238206 := 1.0/t206238
                     else t238206 := 0.0;
end;

procedure TdmDVRD.ConstructRawDataSampleQuery;
var
  i : integer;
begin
  dmDVRD.qRawSmp.Close;
  dmDVRD.qRawSmp.SQL.Clear;
  dmDVRD.qRawSmp.SQL.Add('SELECT DISTINCT SMPDATA.SAMPLENO, SMPDATA.FRAC,');
  dmDVRD.qRawSmp.SQL.Add('  SMPFRAC.ZONEID,');
  dmDVRD.qRawSmp.SQL.Add('  VARREGASSOC.REGASSOCID ');
  dmDVRD.qRawSmp.SQL.Add('FROM SMPDATA,SMPLIST,VARREGASSOC,SMPFRAC ');
  dmDVRD.qRawSmp.SQL.Add('WHERE SMPDATA.SAMPLENO=SMPLIST.SAMPLENO');
  dmDVRD.qRawSmp.SQL.Add('AND VARREGASSOC.REGASSOCID=:RegAssocID');
  dmDVRD.qRawSmp.SQL.Add('AND SMPFRAC.SAMPLENO=SMPDATA.SAMPLENO');
  dmDVRD.qRawSmp.SQL.Add('AND SMPFRAC.FRAC=SMPDATA.FRAC');
  dmDVRD.qRawSmp.SQL.Add('AND SMPLIST.SAMPLENO=:tSampleNo');
  dmDVRD.qRawSmp.SQL.Add('ORDER BY SMPDATA.SAMPLENO, SMPDATA.FRAC');
end;


end.
