object fExJournal: TfExJournal
  Left = 0
  Top = 0
  Caption = #1069#1082#1089#1087#1086#1088#1090' '#1079#1072#1087#1080#1089#1077#1081
  ClientHeight = 163
  ClientWidth = 337
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 337
    Height = 163
    Align = alClient
    TabOrder = 0
    object Label1: TLabel
      Left = 12
      Top = 35
      Width = 14
      Height = 13
      Caption = #1054#1090
    end
    object Label2: TLabel
      Left = 167
      Top = 35
      Width = 13
      Height = 13
      Caption = #1076#1086
    end
    object DateTimePicker1: TDateTimePicker
      Left = 38
      Top = 27
      Width = 113
      Height = 21
      Date = 41120.536932592600000000
      Time = 41120.536932592600000000
      TabOrder = 0
    end
    object DateTimePicker2: TDateTimePicker
      Left = 197
      Top = 27
      Width = 113
      Height = 21
      Date = 41120.537003333330000000
      Time = 41120.537003333330000000
      TabOrder = 1
    end
    object Button1: TButton
      Left = 79
      Top = 70
      Width = 178
      Height = 25
      Caption = #1069#1082#1089#1087#1086#1088#1090' '#1078#1091#1088#1085#1072#1083#1072' '#1091#1095#1077#1090#1072' '#1074' Excel'
      TabOrder = 2
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 79
      Top = 120
      Width = 178
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1085#1080#1077' '#1087#1088#1086#1090#1086#1082#1086#1083#1086#1074' '#1074' Word'
      TabOrder = 3
      OnClick = Button2Click
    end
  end
end
