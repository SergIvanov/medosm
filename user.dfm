object fUser: TfUser
  Left = 0
  Top = 0
  BorderIcons = []
  Caption = #1040#1091#1090#1077#1085#1090#1080#1092#1080#1082#1072#1094#1080#1103
  ClientHeight = 195
  ClientWidth = 238
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 238
    Height = 195
    Align = alClient
    TabOrder = 0
    object Label1: TLabel
      Left = 55
      Top = 24
      Width = 128
      Height = 13
      Caption = #1042#1099#1073#1077#1088#1080#1090#1077' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103':'
    end
    object lbl1: TLabel
      Left = 55
      Top = 83
      Width = 37
      Height = 13
      Caption = #1055#1072#1088#1086#1083#1100
    end
    object Button1: TButton
      Left = 55
      Top = 129
      Width = 145
      Height = 25
      Caption = #1042#1099#1073#1088#1072#1090#1100
      TabOrder = 0
      OnClick = Button1Click
    end
    object ComboBox1: TComboBox
      Left = 55
      Top = 56
      Width = 145
      Height = 21
      TabOrder = 1
      Items.Strings = (
        #1040#1076#1084#1080#1085#1080#1089#1090#1088#1072#1090#1086#1088)
    end
    object edt1: TEdit
      Left = 56
      Top = 102
      Width = 144
      Height = 21
      TabOrder = 2
      OnKeyPress = edt1KeyPress
    end
    object btnClose: TButton
      Left = 56
      Top = 160
      Width = 144
      Height = 25
      Caption = #1047#1072#1082#1088#1099#1090#1100
      TabOrder = 3
      OnClick = btnCloseClick
    end
  end
end
