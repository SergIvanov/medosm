﻿object fRepJASOPat: TfRepJASOPat
  Left = 0
  Top = 0
  Caption = 'fRepJASOPat'
  ClientHeight = 352
  ClientWidth = 630
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 630
    Height = 75
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 24
      Width = 12
      Height = 13
      Caption = #1086#1090
    end
    object до: TLabel
      Left = 135
      Top = 24
      Width = 13
      Height = 13
      Caption = #1076#1086
    end
    object DateTimePicker1: TDateTimePicker
      Left = 32
      Top = 16
      Width = 97
      Height = 21
      Date = 41603.665542847220000000
      Time = 41603.665542847220000000
      TabOrder = 0
    end
    object DateTimePicker2: TDateTimePicker
      Left = 154
      Top = 16
      Width = 97
      Height = 21
      Date = 41603.665572638890000000
      Time = 41603.665572638890000000
      TabOrder = 1
    end
    object CheckBox1: TCheckBox
      Left = 8
      Top = 43
      Width = 64
      Height = 17
      Caption = #1060#1080#1083#1100#1090#1088
      TabOrder = 2
      OnClick = CheckBox1Click
    end
    object Button1: TButton
      Left = 544
      Top = 44
      Width = 75
      Height = 25
      Caption = #1054#1090#1095#1077#1090
      TabOrder = 3
      OnClick = Button1Click
    end
    object DBLookupComboboxEh1: TDBLookupComboboxEh
      Left = 272
      Top = 16
      Width = 253
      Height = 21
      EditButtons = <>
      KeyField = 'id_org'
      ListField = 'imya_org'
      ListSource = fDM.DSOrg
      TabOrder = 4
      Visible = True
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 75
    Width = 630
    Height = 277
    Align = alClient
    TabOrder = 1
    object CheckListBox1: TCheckListBox
      Left = 1
      Top = 1
      Width = 628
      Height = 275
      Align = alClient
      ItemHeight = 13
      TabOrder = 0
    end
  end
end
