object fEqualRab: TfEqualRab
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'fEqualRab'
  ClientHeight = 117
  ClientWidth = 335
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 24
    Top = 10
    Width = 103
    Height = 13
    Caption = #1054#1090#1082#1088#1099#1090#1100' '#1089#1087#1080#1089#1086#1082' .xls'
  end
  object Edit1: TEdit
    Left = 24
    Top = 29
    Width = 242
    Height = 21
    TabOrder = 0
  end
  object Button1: TButton
    Left = 272
    Top = 26
    Width = 33
    Height = 25
    Caption = '...'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 24
    Top = 64
    Width = 75
    Height = 25
    Caption = #1057#1074#1077#1088#1080#1090#1100
    TabOrder = 2
    OnClick = Button2Click
  end
  object OpenDialog1: TOpenDialog
    Left = 120
    Top = 64
  end
end
