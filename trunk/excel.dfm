�
 TFORM1 0\  TPF0TForm1Form1LeftKTop� AlphaBlendValue� BorderIconsbiSystemMenu
biMinimizebiHelp BorderStylebsSingleCaptionNesy v4.0 BetaClientHeight�ClientWidth8Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style Menu	MainMenu1OldCreateOrder	PositionpoScreenCenterOnCreate
FormCreate	OnDestroyFormDestroyOnShowzeigePixelsPerInch`
TextHeight TLabelLabel1Left0Top� Width>HeightCaptionMandant-Nr.:  TLabelLabel2Left0Top� WidthAHeightCaptionBearbeiter-ID:  TLabelLabel3Left0Top� Width5HeightCaptionVorlauf-Nr.:  TLabelLabel4Left0TopWidth6HeightCaptionStartdatum:  TLabelLabel5Left0TopHWidth9HeightCaption
Endedatum:  TLabelLabel6Left0Top� WidthHeightCaptionJahr:  TLabelyearErrorTextLeft1TopWidthHeightColorclRedParentColorVisible  TLabelendeErrorTextLeft1TopcWidthHeightColorclRedParentColorVisible  TLabelstartErrorTextLeft2Top2WidthHeightColorclRedParentColorVisible  TLabelLabel7Left0TopHWidth6HeightCaptionBerater-Nr.:  TLabel
errBeraterLeft1TopbWidthHeightColorclRedParentColorVisible  TButtonbtnStartLeft� TopWidthyHeightHint.Startet die Generierung der Date-Vorlaufdatei.CaptionKonvertieren StartenDefault	ParentShowHintShowHint	TabOrder OnClickkonvertieren  	TComboBoxinhaltLeftTopWidth� HeightHint3Bitte zu importierende Excel-Tabelle hier aussuchenImeMode	imDisable
ItemHeightParentShowHintShowHint	TabOrder  TEditMNRLeft� Top� WidthyHeightHint2Mandant-Nr. aus Zeile 2, Spalte 1 in Excel-TabelleParentColor	ParentShowHintReadOnly	ShowHint	TabOrder  TEditBearbLeft� Top� WidthyHeightHint4Bearbeiter-ID aus Zeile 3, Spalte 1 in Excel-TabelleParentColor	ParentShowHintReadOnly	ShowHint	TabOrder  TEditVorlaufLeft� Top� WidthyHeightHint1Vorlauf-Nummer der generierten DATEV-vorlaufdatei
AutoSelectParentShowHintShowHint	TabOrderText1  TButtonrefreshFromExcelLeft� Top(WidthyHeightHintA   Lädt die Namen der aktuell in MS-Excel  geöffneten Dateien neu.CaptionMappen neu ladenParentShowHintShowHint	TabOrderOnClickrefreshFromExcelClick  TProgressBarfortschrittLeft TopvWidth8HeightAlignalBottomTabOrder  	TComboBoxcbYearsLeft� Top� WidthyHeightHint'Jahr der generierten Datev-Vorlaufdatei
ItemHeightParentShowHintShowHint	TabOrderOnChangecbYearsChangeOnSelectcbYearsSelect  	TComboBoxstartLeft� TopWidthyHeightHint:Startdatum in Form MMTT der generierten Datev-Vorlaufdatei
ItemHeightParentShowHintShowHint	TabOrderOnChangestartChangeOnSelectstartSelect  	TComboBoxendeLeft� TopHWidthyHeightHint9Endedatum in Form MMTT der generierten Datev-Vorlaufdatei
ItemHeightParentShowHintShowHint	TabOrder	OnChange
endeChangeOnSelect
endeSelect  	TComboBoxcomboBeraterNrLeft� TopHWidthyHeightHint
Berater Nr
ItemHeightParentShowHintShowHint	TabOrder
Text115024OnChangecomboBeraterNrChangeOnSelectcbYearsSelectItems.Strings115024204613   TDdeClientConvclientLeft� Top�   TDdeClientItemitemDdeConvclientLeft� Top�   	TMainMenu	MainMenu1LeftxTop 	TMenuItemDatei1Caption&Datei 	TMenuItemInfo1Caption&InfoShortCutI@OnClick
Info1Click  	TMenuItem	Optionen1Caption	&OptionenShortCutO@OnClickOptionen1Click  	TMenuItemN1Caption-  	TMenuItemBeenden1Caption&BeendenOnClickBeenden1Click     