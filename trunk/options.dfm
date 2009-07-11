object OptionDialog: TOptionDialog
  Left = 0
  Top = 0
  Width = 434
  Height = 183
  AlphaBlend = True
  AlphaBlendValue = 180
  Caption = 'Optionen'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 24
    Top = 16
    Width = 105
    Height = 13
    Caption = 'Template-Verzeichnis:'
  end
  object Label2: TLabel
    Left = 24
    Top = 80
    Width = 103
    Height = 13
    Caption = 'Ausgabe-Verzeichnis:'
  end
  object edTemplateDir: TEdit
    Left = 24
    Top = 40
    Width = 329
    Height = 21
    TabOrder = 0
    Text = 'a:\nesy'
  end
  object edSaveDir: TEdit
    Left = 24
    Top = 96
    Width = 329
    Height = 21
    TabOrder = 1
    Text = 'a:\'
  end
  object btnSelectTemplateDir: TButton
    Left = 352
    Top = 40
    Width = 51
    Height = 25
    Caption = '...'
    TabOrder = 2
    OnClick = btnSelectTemplateDirClick
  end
  object btnSelectSaveDir: TButton
    Left = 352
    Top = 96
    Width = 51
    Height = 25
    Caption = '...'
    TabOrder = 3
    OnClick = btnSelectSaveDirClick
  end
end
