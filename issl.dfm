object fIssl: TfIssl
  Left = 0
  Top = 0
  Caption = #1048#1089#1089#1083#1077#1076#1086#1074#1072#1085#1080#1103
  ClientHeight = 300
  ClientWidth = 635
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
    Top = 260
    Width = 635
    Height = 40
    Align = alBottom
    TabOrder = 0
    object DBNavigator1: TDBNavigator
      Left = 16
      Top = 5
      Width = 240
      Height = 25
      DataSource = fDM.DSIssl
      Flat = True
      TabOrder = 0
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 635
    Height = 260
    Align = alClient
    TabOrder = 1
    object DBGridEh1: TDBGridEh
      Left = 1
      Top = 1
      Width = 633
      Height = 258
      Align = alClient
      DataGrouping.GroupLevels = <>
      DataSource = fDM.DSIssl
      Flat = False
      FooterColor = clWindow
      FooterFont.Charset = DEFAULT_CHARSET
      FooterFont.Color = clWindowText
      FooterFont.Height = -11
      FooterFont.Name = 'Tahoma'
      FooterFont.Style = []
      IndicatorOptions = [gioShowRowIndicatorEh]
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
      Columns = <
        item
          EditButtons = <>
          FieldName = 'lookissl'
          Footers = <>
          Title.Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
          Title.Font.Charset = DEFAULT_CHARSET
          Title.Font.Color = clWindowText
          Title.Font.Height = -11
          Title.Font.Name = 'Tahoma'
          Title.Font.Style = []
        end
        item
          EditButtons = <>
          FieldName = 'date1'
          Footers = <>
          Title.Caption = #1044#1072#1090#1072'1'
          Width = 70
        end
        item
          EditButtons = <>
          FieldName = 'znach1'
          Footers = <>
          Title.Caption = #1047#1085#1072#1095#1077#1085#1080#1077'1'
          Width = 160
        end
        item
          EditButtons = <>
          FieldName = 'date2'
          Footers = <>
          Title.Caption = #1044#1072#1090#1072'2'
          Width = 70
        end
        item
          EditButtons = <>
          FieldName = 'znach2'
          Footers = <>
          Title.Caption = #1047#1085#1072#1095#1077#1085#1080#1077'2'
          Width = 160
        end
        item
          EditButtons = <>
          FieldName = 'date3'
          Footers = <>
          Title.Caption = #1044#1072#1090#1072'3'
          Width = 70
        end
        item
          EditButtons = <>
          FieldName = 'znach3'
          Footers = <>
          Title.Caption = #1047#1085#1072#1095#1077#1085#1080#1077'3'
          Width = 160
        end>
      object RowDetailData: TRowDetailPanelControlEh
      end
    end
  end
end
