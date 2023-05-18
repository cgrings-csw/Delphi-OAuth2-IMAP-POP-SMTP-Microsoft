object FormIMAPTest: TFormIMAPTest
  Left = 0
  Top = 0
  Caption = 'OAuth2 Microsoft Email (OAuth2, IMPA4, POP3 e SMTP)'
  ClientHeight = 564
  ClientWidth = 935
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    935
    564)
  PixelsPerInch = 96
  TextHeight = 13
  object lblTenantId: TLabel
    Left = 8
    Top = 8
    Width = 45
    Height = 13
    Caption = 'Id tenant'
  end
  object lblEmailAccount: TLabel
    Left = 543
    Top = 49
    Width = 65
    Height = 13
    Caption = 'Email account'
  end
  object lblClientId: TLabel
    Left = 8
    Top = 49
    Width = 38
    Height = 13
    Caption = 'Client id'
  end
  object lblClientSecret: TLabel
    Left = 243
    Top = 49
    Width = 60
    Height = 13
    Caption = 'Client secret'
  end
  object lblScope: TLabel
    Left = 570
    Top = 8
    Width = 29
    Height = 13
    Caption = 'Scope'
  end
  object lblEmailPassword: TLabel
    Left = 781
    Top = 49
    Width = 73
    Height = 13
    Caption = 'Email password'
  end
  object lblUrlToken: TLabel
    Left = 243
    Top = 8
    Width = 43
    Height = 13
    Caption = 'Url token'
  end
  object btnOAuth2: TButton
    Left = 8
    Top = 104
    Width = 75
    Height = 25
    Caption = 'OAuth2'
    TabOrder = 7
    OnClick = btnOAuth2Click
  end
  object btn_Test_outlook_IMAP: TButton
    Left = 87
    Top = 104
    Width = 75
    Height = 25
    Caption = 'IMAP4'
    TabOrder = 8
    OnClick = btn_Test_outlook_IMAPClick
  end
  object btnSmtp: TButton
    Left = 245
    Top = 104
    Width = 75
    Height = 25
    Caption = 'SMTP'
    TabOrder = 10
    OnClick = btnSmtpClick
  end
  object btnPop: TButton
    Left = 166
    Top = 104
    Width = 75
    Height = 25
    Caption = 'POP3'
    TabOrder = 9
    OnClick = btnPopClick
  end
  object Memo1: TMemo
    Left = 8
    Top = 144
    Width = 919
    Height = 412
    Anchors = [akLeft, akTop, akRight, akBottom]
    ScrollBars = ssBoth
    TabOrder = 11
    ExplicitWidth = 799
  end
  object edtTenantId: TEdit
    Left = 8
    Top = 24
    Width = 217
    Height = 21
    TabOrder = 0
  end
  object edtEmailAccount: TEdit
    Left = 541
    Top = 64
    Width = 234
    Height = 21
    TabOrder = 5
  end
  object edtClientId: TEdit
    Left = 8
    Top = 64
    Width = 229
    Height = 21
    TabOrder = 3
  end
  object edtClientSecret: TEdit
    Left = 241
    Top = 64
    Width = 296
    Height = 21
    PasswordChar = '*'
    TabOrder = 4
  end
  object edtScope: TEdit
    Left = 554
    Top = 24
    Width = 373
    Height = 21
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 2
  end
  object edtEmailPassword: TEdit
    Left = 779
    Top = 64
    Width = 148
    Height = 21
    PasswordChar = '*'
    TabOrder = 6
  end
  object edtUrlToken: TEdit
    Left = 229
    Top = 24
    Width = 321
    Height = 21
    TabOrder = 1
  end
  object chkLogDetailAuth: TCheckBox
    Left = 352
    Top = 108
    Width = 97
    Height = 17
    Caption = 'Detail OAuth2'
    TabOrder = 12
    OnClick = chkLogDetailAuthClick
  end
  object chkLogDetailImap: TCheckBox
    Left = 471
    Top = 108
    Width = 97
    Height = 17
    Caption = 'Detail IMAP'
    TabOrder = 13
    OnClick = chkLogDetailImapClick
  end
  object chkLogDetailPop: TCheckBox
    Left = 591
    Top = 108
    Width = 97
    Height = 17
    Caption = 'Detail POP'
    TabOrder = 14
    OnClick = chkLogDetailPopClick
  end
  object chkLogDetailSmtp: TCheckBox
    Left = 710
    Top = 108
    Width = 97
    Height = 17
    Caption = 'Detail SMTP'
    TabOrder = 15
    OnClick = chkLogDetailSmtpClick
  end
end
