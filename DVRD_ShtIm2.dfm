�
 TFMSHEETIMPORT2 0�%  TPF0TfmSheetImport2fmSheetImport2LeftRTopyCaptionImport spreadsheet definitionsClientHeight
ClientWidthColor	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclBlackFont.Height�	Font.NameArial
Font.Style 	Icon.Data
�             �     (       @         �                        �  �   �� �   � � ��  ��� ���   �  �   �� �   � � ��  ��� �����������������DO �DDDDDDDDDI�DO �DDDDDDDDDI��ODDDD   DDI��ODD@ ��� DI����D@�����I���  @������I���������� I��� @����� �I�������� wxxI�������������DI��������xx�DI��������w��DI��������wxxxDI������w���DDI��������xxDDI���������DDI������ww���DDDI������w��DDDI����wxx� DDDDI����     ODDDI����wxxx����DDI�������� �DDI���xxxp�� �DDI���������DDI���xxx�����DDI����ww ������DDI����w ��������DI���  ���������DI���������������I���������������I����������������                                                                                                                                OnShowFormShow
TextHeight 	TSplitter	Splitter1Left Top� WidthHeightCursorcrVSplitAlignalBottomBeveled	ExplicitLeft�  TPanelpControlLeft Top WidthHeight!AlignalTop
BevelInnerbvRaised
BevelOuter	bvLoweredTabOrder ExplicitWidth TLabel	lFilePathLeft Top
Width'HeightCaption	lFilePath  TBitBtnbbOpenSheetLeftXTopWidthKHeightHintSelect spread sheetCaption&Open
ImageIndex	ImageNamefolderImagesVirtualImageList1	NumGlyphsTabOrder OnClickbbOpenSheetClick  TBitBtnbbCancelLeftTopWidthKHeight
ImageIndex		ImageNamecancelImagesVirtualImageList1KindbkCancel	NumGlyphsTabOrderOnClickbbCancelClick  TBitBtnbbImportLeft� TopWidthKHeightHint%Import selected data from spreadsheetCaptionOKDefault	
ImageIndex$	ImageNameokImagesVirtualImageList1ModalResult	NumGlyphsTabOrderOnClickbbImportClick   
TStatusBarsbSheetLeft Top�WidthHeightPanels SimplePanel	ExplicitTop�ExplicitWidth  TPanelpSpreadSheetLeft Top!WidthHeight� AlignalClient
BevelOuterbvNoneTabOrderExplicitWidthExplicitHeight�  TTabControl
TabControlLeft Top WidthHeight� AlignalClientTabOrder TabPositiontpBottomOnChangeTabControlChangeExplicitWidthExplicitHeight�   TStringGrid	SheetDataLeft Top WidthHeight� AlignalClientCtl3DDefaultRowHeightOptionsgoFixedVertLinegoFixedHorzLine
goVertLine
goHorzLinegoDrawFocusSelectedgoRowSizinggoColSizinggoThumbTracking ParentCtl3DTabOrderOnSelectCellSheetDataSelectCellExplicitWidthExplicitHeight� 	ColWidths@@@@@ 
RowHeights   TTabSetTabsLeft Top� WidthHeightAlignalBottomEnabledFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameTahoma
Font.Style SoftTop	Style
tsSoftTabsOnChange
TabsChangeExplicitTop� ExplicitWidth   TPanelpDefinitionsLeft Top� WidthHeight� AlignalBottomTabOrderExplicitTop� ExplicitWidth 	TSplitter	Splitter2LeftuTop8WidthHeight� Beveled	ColorclMoneyGreenParentColorExplicitHeight�   	TGroupBoxgbDefineFieldsLeftTop8WidthtHeight� Hint*Specify column letter or 0 (zero) to omit.AlignalLeftCaptionDefine fieldsTabOrder  TLabelLabel5LeftTop WidthZHeightCaptionImport Spec. Name  TLabelLabel6LeftTop:Width%HeightCaptionPosition  TLabelLabel7LeftTopVWidthGHeightCaptionVariable Called  TLabelLabel10LeftlTopWidth#HeightCaptionColumn  TLabelLabel12LeftTopsWidth]HeightCaptionColumn (character)  TEditeImportSpecNameColLeftpTopWidthHeightTabOrder TexteImportSpecNameCol  TEditePositionColLeftpTop8WidthHeightTabOrderTextePositionCol  TEdit
eCalledColLeftpTopTWidthHeightTabOrderText
eCalledCol  TEdit
eColumnColLeftpTopoWidthHeightTabOrderText
eColumnCol  TMemoMemo1Left� TopWidth� Height� 
BevelInnerbvNone	BevelKindbkSoftBorderStylebsNoneColor	clBtnFaceLines.Strings!Variables with "Position" values from -5 to 0 (zero) must be 	provided #for every import specification and "must always link to the following information (names may be changed, however):   -5 Reference number (must match a reference already in 	DateView)   -4 Sample number   -3 Fraction    -2 ZoneID (must match lookup values already in DateView)#   -1 Technique abbreviation (must match lookup values already in 
DateView) %    0 Material analysed abbreviation "(must match lookup values already in 
DateView)      TabOrder   TPanelPanel3LeftTopWidthHeight7AlignalTop
BevelOuterbvNoneTabOrderExplicitWidth 	TGroupBoxgbDefineRowsLeft
TopWidth�Height.CaptionDefine rows to importTabOrder TLabelLabel2LeftTopWidth/HeightCaptionFrom row  TLabelLabel3LeftrTopWidth"HeightCaptionTo row  TSpeedButtonsbFindLastRowLeft� TopWidthHeightHintFind end of file	ImageName
history_b1ImagesVirtualImageList1	NumGlyphsParentShowHintShowHint	OnClicksbFindLastRowClick  TLabelLabel21LeftTopWidth� HeightCaption"Number of rows is based on finding  TLabelLabel22LeftTopWidth� HeightCaption!values in the Import Spec. column  TEditeFromRowLeft<TopWidth1HeightTabOrder Text2  TEditeToRowLeft� TopWidth1HeightTabOrderText3   	TGroupBoxgbDefineTabSheetLeftTopWidth� Height.Caption!Define sheet from which to importTabOrder  	TComboBoxcbSheetNameLeftTopWidth� HeightTabOrder TextSheet1OnChangecbSheetNameChange    	TGroupBox
gbDefaultsLeft}Top8Width�Height� AlignalClientCaptionDefault valuesTabOrderExplicitWidth� TLabelLabel1LeftTopWidthMHeightCaptionDefault minimum  TEditeDefaultMinimumLeftlTopWidthiHeightTabOrder TexteDefaultMinimum    TOpenDialogOpenDialogSprdSheet
DefaultExt.XLSXFilterIExcel 1997-2013|*.XLS;*.XLSX|Excel 1997-2003|*.XLS|Excel 2007-2010|*.XLSXOptions
ofShowHelpofPathMustExistofFileMustExistofShareAware LeftLTop  TVirtualImageListVirtualImageList1AutoFill	ImagesCollectionIndex CollectionNameaboutNameabout CollectionIndexCollectionNamealphabetical_sorting_azNamealphabetical_sorting_az CollectionIndexCollectionNameapprovalNameapproval CollectionIndexCollectionNameapproveNameapprove CollectionIndexCollectionName
area_chartName
area_chart CollectionIndexCollectionName	automaticName	automatic CollectionIndexCollectionName	bar_chartName	bar_chart CollectionIndexCollectionName
calculatorName
calculator CollectionIndexCollectionNamecalendarNamecalendar CollectionIndex	CollectionNamecancelNamecancel CollectionIndex
CollectionName	checkmarkName	checkmark CollectionIndexCollectionNameclear_filtersNameclear_filters CollectionIndexCollectionNamecombo_chartNamecombo_chart CollectionIndexCollectionName
data_sheetName
data_sheet CollectionIndexCollectionNamedocumentNamedocument CollectionIndexCollectionNamedownloadNamedownload CollectionIndexCollectionNameempty_filterNameempty_filter CollectionIndexCollectionNameexportNameexport CollectionIndexCollectionNamefileNamefile CollectionIndexCollectionNamefilled_filterNamefilled_filter CollectionIndexCollectionNamefolderNamefolder CollectionIndexCollectionNamegeneric_sorting_ascNamegeneric_sorting_asc CollectionIndexCollectionNamegeneric_sorting_descNamegeneric_sorting_desc CollectionIndexCollectionNameglobeNameglobe CollectionIndexCollectionNameheat_mapNameheat_map CollectionIndexCollectionNamehomeNamehome CollectionIndexCollectionName
image_fileName
image_file CollectionIndexCollectionNameimportNameimport CollectionIndexCollectionNameinfoNameinfo CollectionIndexCollectionName
inspectionName
inspection CollectionIndexCollectionNameinternalNameinternal CollectionIndexCollectionName
line_chartName
line_chart CollectionIndex CollectionNamemenuNamemenu CollectionIndex!CollectionNamenextNamenext CollectionIndex"CollectionNamenumerical_sorting_12Namenumerical_sorting_12 CollectionIndex#CollectionNamenumerical_sorting_21Namenumerical_sorting_21 CollectionIndex$CollectionNameokNameok CollectionIndex%CollectionName	pie_chartName	pie_chart CollectionIndex&CollectionNameplusNameplus CollectionIndex'CollectionNamepreviousNameprevious CollectionIndex(CollectionNameprintNameprint CollectionIndex)CollectionNamepuzzleNamepuzzle CollectionIndex*CollectionNamerefreshNamerefresh CollectionIndex+CollectionNamerulesNamerules CollectionIndex,CollectionNamescatter_plotNamescatter_plot CollectionIndex-CollectionNamesearchNamesearch CollectionIndex.CollectionNamesettingsNamesettings CollectionIndex/CollectionNamesupportNamesupport CollectionIndex0CollectionNametemplateNametemplate CollectionIndex1CollectionNameundoNameundo CollectionIndex2CollectionNameuploadNameupload CollectionIndex3CollectionNameview_detailsNameview_details  ImageCollectiondmDVRD.SVGIconImageCollection1Left�   