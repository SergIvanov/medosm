object fEdSpravOrg: TfEdSpravOrg
  Left = 778
  Top = 227
  Caption = 'fEdSpravOrg'
  ClientHeight = 455
  ClientWidth = 704
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 704
    Height = 414
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 662
    object DBGridEh1: TDBGridEh
      Left = 1
      Top = 1
      Width = 702
      Height = 412
      Align = alClient
      DataSource = fDM.DSOrg
      DrawMemoText = True
      DynProps = <>
      FooterParams.Color = clWindow
      STFilter.InstantApply = True
      STFilter.Local = True
      STFilter.Visible = True
      TabOrder = 0
      Columns = <
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'id_org'
          Footers = <>
          Title.Caption = #1050#1086#1076
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'imya_org'
          Footers = <>
          Title.Caption = #1048#1084#1103' '#1089#1086#1082#1088#1072#1097#1077#1085#1085#1086#1077
          Width = 80
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'imya_poln'
          Footers = <>
          Title.Caption = #1055#1086#1083#1085#1086#1077' '#1085#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
          Width = 300
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'form_sobstv'
          Footers = <>
          Title.Caption = #1060#1086#1088#1084#1072' '#1089#1086#1073#1089#1090#1074#1077#1085#1085#1086#1089#1090#1080
          Width = 50
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'dif_price'
          Footers = <>
        end
        item
          CellButtons = <>
          DynProps = <>
          EditButtons = <>
          FieldName = 'zd'
          Footers = <>
          Title.Caption = #1046#1044' '#1087#1088#1077#1076#1087#1088#1080#1103#1090#1080#1077
          Width = 89
        end>
      object RowDetailData: TRowDetailPanelControlEh
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 414
    Width = 704
    Height = 41
    Align = alBottom
    TabOrder = 1
    ExplicitWidth = 662
    object Button1: TButton
      Left = 16
      Top = 6
      Width = 129
      Height = 25
      Caption = #1054#1073#1085#1086#1074#1080#1090#1100' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082
      TabOrder = 0
      OnClick = Button1Click
    end
    object DBNavigator1: TDBNavigator
      Left = 224
      Top = 5
      Width = 240
      Height = 25
      DataSource = fDM.DSOrg
      Flat = True
      TabOrder = 1
    end
  end
end
