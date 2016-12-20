object fChengeCena: TfChengeCena
  Left = 0
  Top = 0
  Caption = #1054#1090#1086#1073#1088#1072#1079#1080#1090#1100' '#1080#1079#1084#1077#1085#1077#1085#1080#1077' '#1094#1077#1085#1099' '#1074' '#1088#1077#1077#1089#1090#1088#1077
  ClientHeight = 142
  ClientWidth = 308
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
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 308
    Height = 142
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 303
    ExplicitHeight = 129
    object Label1: TLabel
      Left = 28
      Top = 8
      Width = 85
      Height = 13
      Caption = #1057#1090#1072#1088#1072#1103' '#1094#1077#1085#1072' (-1)'
    end
    object Label2: TLabel
      Left = 185
      Top = 8
      Width = 83
      Height = 13
      Caption = #1053#1086#1074#1072#1103' '#1094#1077#1085#1072' (+1)'
    end
    object Label3: TLabel
      Left = 28
      Top = 58
      Width = 119
      Height = 13
      Caption = #1057#1090#1072#1088#1072#1103' '#1094#1077#1085#1072' (-1 '#1080#1083#1080' 1):'
    end
    object Label4: TLabel
      Left = 185
      Top = 58
      Width = 113
      Height = 13
      Caption = #1053#1086#1074#1072#1103' '#1094#1077#1085#1072' (-1 '#1080#1083#1080' 1):'
    end
    object DBLookupComboboxEh1: TDBLookupComboboxEh
      Left = 16
      Top = 27
      Width = 121
      Height = 21
      DataField = 'iduslold'
      DataSource = fDM.DSRab
      EditButtons = <>
      KeyField = 'id'
      ListField = 'naim'
      ListSource = fDM.DSSUslM
      TabOrder = 0
      Visible = True
    end
    object DBLookupComboboxEh2: TDBLookupComboboxEh
      Left = 167
      Top = 27
      Width = 121
      Height = 21
      DataField = 'idusl'
      DataSource = fDM.DSRab
      EditButtons = <>
      KeyField = 'id'
      ListField = 'naim'
      ListSource = fDM.DSSUslM
      TabOrder = 1
      Visible = True
    end
    object Button1: TButton
      Left = 117
      Top = 104
      Width = 75
      Height = 25
      Caption = #1054#1082
      TabOrder = 2
      OnClick = Button1Click
    end
    object DBEditEh1: TDBEditEh
      Left = 28
      Top = 77
      Width = 51
      Height = 21
      DataField = 'extestold'
      DataSource = fDM.DSRab
      EditButtons = <>
      MaxLength = 2
      TabOrder = 3
      Visible = True
      EditMask = '##'
    end
    object DBEditEh2: TDBEditEh
      Left = 185
      Top = 77
      Width = 51
      Height = 21
      DataField = 'extest'
      DataSource = fDM.DSRab
      EditButtons = <>
      MaxLength = 2
      TabOrder = 4
      Visible = True
      EditMask = '##'
    end
  end
end
