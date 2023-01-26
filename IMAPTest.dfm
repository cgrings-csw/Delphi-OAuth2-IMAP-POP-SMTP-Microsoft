object FormIMAPTest: TFormIMAPTest
  Left = 0
  Top = 0
  Caption = 'OAuth2 Microsoft Email (OAuth2, IMPA4, POP3 e SMTP)'
  ClientHeight = 564
  ClientWidth = 815
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    815
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
    Left = 511
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
    Left = 702
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
    Left = 89
    Top = 104
    Width = 75
    Height = 25
    Caption = 'IMAP4'
    TabOrder = 8
    OnClick = btn_Test_outlook_IMAPClick
  end
  object btnSmtp: TButton
    Left = 251
    Top = 104
    Width = 75
    Height = 25
    Caption = 'SMTP'
    TabOrder = 10
    OnClick = btnSmtpClick
  end
  object btnPop: TButton
    Left = 170
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
    Width = 799
    Height = 412
    Anchors = [akLeft, akTop, akRight, akBottom]
    ScrollBars = ssBoth
    TabOrder = 15
  end
  object edtTenantId: TEdit
    Left = 8
    Top = 24
    Width = 217
    Height = 21
    TabOrder = 0
  end
  object edtEmailAccount: TEdit
    Left = 511
    Top = 64
    Width = 185
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
    Left = 243
    Top = 64
    Width = 262
    Height = 21
    TabOrder = 4
  end
  object edtScope: TEdit
    Left = 570
    Top = 24
    Width = 237
    Height = 21
    TabOrder = 2
  end
  object edtEmailPassword: TEdit
    Left = 702
    Top = 64
    Width = 105
    Height = 21
    PasswordChar = '*'
    TabOrder = 6
  end
  object edtUrlToken: TEdit
    Left = 243
    Top = 24
    Width = 321
    Height = 21
    TabOrder = 1
  end
  object chkLogDetailImap: TCheckBox
    Left = 461
    Top = 108
    Width = 97
    Height = 17
    Caption = 'Log detail IMAP'
    TabOrder = 12
    OnClick = chkLogDetailImapClick
  end
  object chkLogDetailPop: TCheckBox
    Left = 582
    Top = 108
    Width = 97
    Height = 17
    Caption = 'Log detail POP'
    TabOrder = 13
    OnClick = chkLogDetailPopClick
  end
  object chkLogDetailSmtp: TCheckBox
    Left = 702
    Top = 108
    Width = 97
    Height = 17
    Caption = 'Log detail SMTP'
    TabOrder = 14
    OnClick = chkLogDetailSmtpClick
  end
  object chkLogDetailAuth: TCheckBox
    Left = 341
    Top = 108
    Width = 97
    Height = 17
    Caption = 'Log detail OAuth2'
    TabOrder = 11
    OnClick = chkLogDetailAuthClick
  end
end
