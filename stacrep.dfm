object fStacRep: TfStacRep
  Left = 0
  Top = 0
  Caption = 'fStacRep'
  ClientHeight = 392
  ClientWidth = 720
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 720
    Height = 105
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 16
      Width = 14
      Height = 13
      Caption = #1054#1090
    end
    object Label2: TLabel
      Left = 184
      Top = 16
      Width = 13
      Height = 13
      Caption = #1076#1086
    end
    object DateTimePicker1: TDateTimePicker
      Left = 48
      Top = 16
      Width = 121
      Height = 21
      Date = 40955.635442789350000000
      Time = 40955.635442789350000000
      TabOrder = 0
    end
    object DateTimePicker2: TDateTimePicker
      Left = 216
      Top = 16
      Width = 129
      Height = 21
      Date = 40955.635522326390000000
      Time = 40955.635522326390000000
      TabOrder = 1
    end
    object Button1: TButton
      Left = 376
      Top = 11
      Width = 89
      Height = 25
      Caption = #1057#1092#1086#1088#1084#1080#1088#1086#1074#1072#1090#1100
      TabOrder = 2
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 496
      Top = 11
      Width = 193
      Height = 25
      Caption = #1042#1099#1075#1088#1091#1079#1080#1090#1100
      TabOrder = 3
      OnClick = Button2Click
    end
    object DBLookupComboboxEh1: TDBLookupComboboxEh
      Left = 48
      Top = 64
      Width = 121
      Height = 21
      EditButtons = <>
      KeyField = 'id'
      ListField = 'naim'
      ListSource = fDM.DSSUsl
      TabOrder = 4
      Visible = True
    end
    object Button3: TButton
      Left = 496
      Top = 42
      Width = 193
      Height = 31
      Caption = #1042#1099#1075#1088#1091#1079#1080#1090#1100' '#1087#1088#1086#1092#1094#1077#1085#1090#1088' 1'#1088' '#1074' 5'#1083' '#1072#1084#1073
      TabOrder = 5
      OnClick = Button3Click
    end
  end
  object DBGridEh1: TDBGridEh
    Left = 0
    Top = 105
    Width = 720
    Height = 287
    Align = alClient
    DataGrouping.GroupLevels = <>
    DataSource = fDM.DSStacRep
    DrawMemoText = True
    Flat = False
    FooterColor = clWindow
    FooterFont.Charset = DEFAULT_CHARSET
    FooterFont.Color = clWindowText
    FooterFont.Height = -11
    FooterFont.Name = 'Tahoma'
    FooterFont.Style = []
    IndicatorOptions = [gioShowRowIndicatorEh]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    Columns = <
      item
        EditButtons = <>
        FieldName = 'fam'
        Footers = <>
        Width = 90
      end
      item
        EditButtons = <>
        FieldName = 'imya'
        Footers = <>
      end
      item
        EditButtons = <>
        FieldName = 'fath'
        Footers = <>
      end
      item
        EditButtons = <>
        FieldName = 'data_r'
        Footers = <>
        Width = 90
      end
      item
        EditButtons = <>
        FieldName = 'profes'
        Footers = <>
        Width = 100
      end
      item
        EditButtons = <>
        FieldName = 'vredn_fact'
        Footers = <>
      end
      item
        EditButtons = <>
        FieldName = 'staz_rab_fact'
        Footers = <>
      end
      item
        EditButtons = <>
        FieldName = 'data_proh'
        Footers = <>
        Width = 90
      end
      item
        EditButtons = <>
        FieldName = 'idusl'
        Footers = <>
      end>
    object RowDetailData: TRowDetailPanelControlEh
    end
  end
end
