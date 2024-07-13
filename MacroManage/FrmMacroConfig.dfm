object MacroConfigF: TMacroConfigF
  Left = 0
  Top = 0
  Caption = 'Macro Config'
  ClientHeight = 299
  ClientWidth = 343
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object JvLabel6: TJvLabel
    AlignWithMargins = True
    Left = 8
    Top = 40
    Width = 100
    Height = 25
    Alignment = taCenter
    AutoSize = False
    Caption = 'MacroName'
    Color = 14671839
    FrameColor = clGrayText
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #47569#51008' '#44256#46357
    Font.Style = [fsBold]
    Layout = tlCenter
    ParentColor = False
    ParentFont = False
    RoundedFrame = 3
    Transparent = True
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -13
    HotTrackFont.Name = #47569#51008' '#44256#46357
    HotTrackFont.Style = []
  end
  object JvLabel1: TJvLabel
    AlignWithMargins = True
    Left = 8
    Top = 80
    Width = 100
    Height = 25
    Alignment = taCenter
    AutoSize = False
    Caption = 'MacroDesc'
    Color = 14671839
    FrameColor = clGrayText
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #47569#51008' '#44256#46357
    Font.Style = [fsBold]
    Layout = tlCenter
    ParentColor = False
    ParentFont = False
    RoundedFrame = 3
    Transparent = True
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -13
    HotTrackFont.Name = #47569#51008' '#44256#46357
    HotTrackFont.Style = []
  end
  object JvLabel2: TJvLabel
    AlignWithMargins = True
    Left = 8
    Top = 120
    Width = 100
    Height = 25
    Alignment = taCenter
    AutoSize = False
    Caption = 'RepeatCount'
    Color = 14671839
    FrameColor = clGrayText
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #47569#51008' '#44256#46357
    Font.Style = [fsBold]
    Layout = tlCenter
    ParentColor = False
    ParentFont = False
    RoundedFrame = 3
    Transparent = True
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -13
    HotTrackFont.Name = #47569#51008' '#44256#46357
    HotTrackFont.Style = []
  end
  object JvLabel3: TJvLabel
    AlignWithMargins = True
    Left = 8
    Top = 160
    Width = 100
    Height = 25
    Alignment = taCenter
    AutoSize = False
    Caption = 'ActionDesc'
    Color = 14671839
    FrameColor = clGrayText
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #47569#51008' '#44256#46357
    Font.Style = [fsBold]
    Layout = tlCenter
    ParentColor = False
    ParentFont = False
    RoundedFrame = 3
    Transparent = True
    HotTrackFont.Charset = ANSI_CHARSET
    HotTrackFont.Color = clWindowText
    HotTrackFont.Height = -13
    HotTrackFont.Name = #47569#51008' '#44256#46357
    HotTrackFont.Style = []
  end
  object MacroName: TEdit
    Left = 115
    Top = 40
    Width = 185
    Height = 21
    Hint = 'Text'
    Alignment = taCenter
    ImeName = 'Microsoft IME 2010'
    TabOrder = 0
  end
  object MacroDesc: TEdit
    Left = 115
    Top = 80
    Width = 185
    Height = 21
    Hint = 'Text'
    Alignment = taCenter
    ImeName = 'Microsoft IME 2010'
    TabOrder = 1
  end
  object RepeatCount: TEdit
    Left = 115
    Top = 120
    Width = 185
    Height = 21
    Hint = 'Text'
    Alignment = taCenter
    ImeName = 'Microsoft IME 2010'
    TabOrder = 2
  end
  object ActionDesc: TEdit
    Left = 115
    Top = 160
    Width = 185
    Height = 21
    Hint = 'Text'
    Alignment = taCenter
    ImeName = 'Microsoft IME 2010'
    TabOrder = 3
  end
  object IsExecute: TCheckBox
    Left = 8
    Top = 200
    Width = 121
    Height = 17
    Hint = 'Checked'
    Alignment = taLeftJustify
    Caption = 'IsExecute'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #47569#51008' '#44256#46357
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
  end
  object IsDisplayCustomDesc: TCheckBox
    Left = 8
    Top = 231
    Width = 121
    Height = 17
    Hint = 'Checked'
    Alignment = taLeftJustify
    Caption = 'IsDisplayCustomDesc'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #47569#51008' '#44256#46357
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 5
  end
  object Panel1: TPanel
    Left = 0
    Top = 258
    Width = 343
    Height = 41
    Align = alBottom
    TabOrder = 6
    ExplicitLeft = 80
    ExplicitTop = 280
    ExplicitWidth = 185
    object BitBtn1: TBitBtn
      Left = 54
      Top = 8
      Width = 75
      Height = 25
      Kind = bkCancel
      NumGlyphs = 2
      TabOrder = 0
    end
    object BitBtn2: TBitBtn
      Left = 192
      Top = 8
      Width = 75
      Height = 25
      Kind = bkOK
      NumGlyphs = 2
      TabOrder = 1
    end
  end
end
