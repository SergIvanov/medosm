object fRepJASO: TfRepJASO
  Left = 0
  Top = 0
  Caption = 'fRepJASO'
  ClientHeight = 176
  ClientWidth = 286
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 28
    Width = 14
    Height = 13
    Caption = #1054#1090
  end
  object Label2: TLabel
    Left = 149
    Top = 28
    Width = 13
    Height = 13
    Caption = #1076#1086
  end
  object Label3: TLabel
    Left = 16
    Top = 59
    Width = 71
    Height = 13
    Caption = #1054#1088#1075#1072#1085#1080#1079#1079#1072#1094#1080#1103
  end
  object DateTimePicker1: TDateTimePicker
    Left = 38
    Top = 24
    Width = 97
    Height = 21
    Date = 41372.430463067130000000
    Time = 41372.430463067130000000
    TabOrder = 0
  end
  object DateTimePicker2: TDateTimePicker
    Left = 173
    Top = 24
    Width = 94
    Height = 21
    Date = 41372.430548622690000000
    Time = 41372.430548622690000000
    TabOrder = 1
  end
  object DBLookupComboboxEh1: TDBLookupComboboxEh
    Left = 14
    Top = 81
    Width = 253
    Height = 21
    EditButtons = <>
    KeyField = 'id_org'
    ListField = 'imya_org'
    ListSource = fDM.DSOrg
    TabOrder = 2
    Visible = True
  end
  object Button1: TButton
    Left = 8
    Top = 128
    Width = 116
    Height = 25
    Caption = #1042#1099#1075#1088#1091#1079#1080#1090#1100
    TabOrder = 3
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 149
    Top = 128
    Width = 119
    Height = 25
    Caption = #1042#1099#1073#1088#1072#1090#1100' '#1087#1072#1094#1080#1077#1085#1090#1086#1074
    TabOrder = 4
    OnClick = Button2Click
  end
end
