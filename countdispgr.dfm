object fCountDispGr: TfCountDispGr
  Left = 0
  Top = 0
  Caption = 'fCountDispGr'
  ClientHeight = 292
  ClientWidth = 360
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 360
    Height = 292
    Align = alClient
    TabOrder = 0
    ExplicitLeft = 24
    ExplicitTop = 8
    ExplicitHeight = 397
    object Label1: TLabel
      Left = 48
      Top = 20
      Width = 14
      Height = 13
      Caption = #1054#1090
    end
    object Label2: TLabel
      Left = 183
      Top = 20
      Width = 13
      Height = 13
      Caption = #1076#1086
    end
    object DateTimePicker1: TDateTimePicker
      Left = 68
      Top = 16
      Width = 97
      Height = 21
      Date = 41025.437048900470000000
      Time = 41025.437048900470000000
      TabOrder = 0
    end
    object DateTimePicker2: TDateTimePicker
      Left = 202
      Top = 16
      Width = 97
      Height = 21
      Date = 41025.437102696760000000
      Time = 41025.437102696760000000
      TabOrder = 1
    end
    object Button1: TButton
      Left = 137
      Top = 241
      Width = 75
      Height = 25
      Caption = #1054#1082
      TabOrder = 2
    end
    object DBGridEh1: TDBGridEh
      Left = 0
      Top = 115
      Width = 360
      Height = 120
      DataGrouping.GroupLevels = <>
      Flat = False
      FooterColor = clWindow
      FooterFont.Charset = DEFAULT_CHARSET
      FooterFont.Color = clWindowText
      FooterFont.Height = -11
      FooterFont.Name = 'Tahoma'
      FooterFont.Style = []
      IndicatorOptions = [gioShowRowIndicatorEh]
      TabOrder = 3
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
      object RowDetailData: TRowDetailPanelControlEh
      end
    end
    object RadioButton1: TRadioButton
      Left = 68
      Top = 56
      Width = 197
      Height = 17
      Caption = #1055#1086' '#1075#1088#1091#1087#1087#1072#1084
      TabOrder = 4
    end
    object RadioButton2: TRadioButton
      Left = 68
      Top = 79
      Width = 144
      Height = 17
      Caption = #1042#1087#1077#1088#1074#1099#1077' '#1074#1099#1103#1074#1083#1077#1085#1085#1099#1077
      TabOrder = 5
    end
  end
end
