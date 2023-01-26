unit IMAPTest;

interface

uses
  Vcl.Forms,
  Vcl.Controls,
  Vcl.StdCtrls,
  Vcl.ComCtrls,
  System.SysUtils,
  System.Types,
  System.UITypes,
  System.Classes,
  System.Variants,
  Data.Bind.Components,
  Data.Bind.ObjectScope,
  IdTCPClient,
  IdExplicitTLSClientServerBase,
  IdMessageClient,
  IdIMAP4,
  IdBaseComponent,
  IdComponent,
  IdIOHandler,
  IdIOHandlerSocket,
  IdSASLCollection,
  IdIOHandlerStack,
  IdSSL,
  IdSSLOpenSSL,
  IdTCPConnection,
  IdSASLXOAUTH,
  ROPCFlow,
  Global,
  IdPOP3,
  IdSMTPBase,
  IdSMTP;

type
  TFormIMAPTest = class(TForm)
    btnOAuth2: TButton;
    btnPop: TButton;
    btnSmtp: TButton;
    btn_Test_outlook_IMAP: TButton;
    edtClientId: TEdit;
    edtClientSecret: TEdit;
    edtEmailAccount: TEdit;
    edtEmailPassword: TEdit;
    edtScope: TEdit;
    edtTenantId: TEdit;
    edtUrlToken: TEdit;
    lblClientId: TLabel;
    lblClientSecret: TLabel;
    lblEmailAccount: TLabel;
    lblEmailPassword: TLabel;
    lblScope: TLabel;
    lblTenantId: TLabel;
    lblUrlToken: TLabel;
    Memo1: TMemo;
    procedure btnOAuth2Click(Sender: TObject);
    procedure btnPopClick(Sender: TObject);
    procedure btnSmtpClick(Sender: TObject);
    procedure btn_Test_outlook_IMAPClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  strict private
  const
    FHost: string = 'outlook.office365.com';
    FScope: string = 'https://outlook.office365.com/.default';
    FUrlToken: string = 'https://login.microsoftonline.com/%s/oauth2/v2.0/token';
  var
    FIdIMAP4: TIdIMAP4;
    FIdPOP3: TIdPOP3;
    FIdSMTP: TIdSMTP;
    FIdSSLIOHandlerSocketOpenSSLImap: TIdSSLIOHandlerSocketOpenSSL;
    FIdSSLIOHandlerSocketOpenSSLPop: TIdSSLIOHandlerSocketOpenSSL;
    FIdSSLIOHandlerSocketOpenSSLSmtp: TIdSSLIOHandlerSocketOpenSSL;
    FRopcFlow: TropcFlow;
    FxOAuthSASLImap: TIdSASLListEntry;
    FxOAuthSASLPop3: TIdSASLListEntry;
    FxOAuthSASLSmtp: TIdSASLListEntry;
    procedure AfterAccessToken(const AccessToken, TokenType: string; const ExpiresIn: Int32; const Scope: string);
    procedure AutentichOAuth2;
    procedure ConnectImap;
    procedure ConnectPop;
    procedure ConnectSmtp;
    procedure CrateIdSSLIOHandlerSocketOpenSSLImap;
    procedure CrateIdSSLIOHandlerSocketOpenSSLSmtp;
    procedure CrearteIdIMAP4;
    procedure CreateIdPOP3;
    procedure CreateIdSMTP;
    procedure CreateIdSSLIOHandlerSocketOpenSSLPop;
    procedure DoLog(const Text: string);
    procedure ErrorAccessToken(const Error, ErrorDescription: string);
    procedure OAuth2AuthorizationCodeAfterAccessToken(Sender: TObject; const Access_Token, Token_Type, Expires_In, Refresh_Token, Scope, RawParams: string; var
      Handled: Boolean);
    procedure OAuth2ClientCredentialsAfterAccessToken(const Sender: TObject; const AccessToken, TokenType, ExpiresIn, RefreshToken, Scope, RawParams: string; var
      Handled: Boolean);
    procedure SettingAuthentication;
    procedure StatusConnection(ASender: TObject; const AStatus: TIdStatus; const AStatusText: string);
  end;

var
  FormIMAPTest: TFormIMAPTest;

implementation

uses
  IdSASL, IdMessage, XSuperObject;

{$R *.dfm}

procedure TFormIMAPTest.AfterAccessToken(const AccessToken, TokenType: string; const ExpiresIn: Int32; const Scope: string);
begin
  DoLog('AccessToken: ' + AccessToken);
  DoLog('Token_Type: ' + TokenType);
  DoLog(Format('Expires_In: %d', [ExpiresIn]));
  DoLog('Scope: ' + Scope);

  TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token := AccessToken;
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).ExpireTime := IntToStr(ExpiresIn);
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).User := edtEmailAccount.Text; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).Token := AccessToken;
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).ExpireTime := IntToStr(ExpiresIn);
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).User := edtEmailAccount.Text; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token := AccessToken;
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).ExpireTime := IntToStr(ExpiresIn);
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).User := edtEmailAccount.Text; // outlook email account
end;

procedure TFormIMAPTest.AutentichOAuth2;
begin
  SettingAuthentication;
  DoLog('Start ROPC Flow');
  FRopcFlow.Start;
end;

procedure TFormIMAPTest.btnOAuth2Click(Sender: TObject);
begin
  AutentichOAuth2;
end;

procedure TFormIMAPTest.btnPopClick(Sender: TObject);
begin
  ConnectPop;
end;

procedure TFormIMAPTest.btnSmtpClick(Sender: TObject);
begin
  ConnectSmtp;
end;

procedure TFormIMAPTest.btn_Test_outlook_IMAPClick(Sender: TObject);
begin
  ConnectImap;
end;

procedure TFormIMAPTest.ConnectImap;
begin
  try
    try
      if not TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token.IsEmpty then
      begin
        if not TIdSASLXOAuth(FxOAuthSASLImap.SASL).IsTokenExpired then
        begin
          FIdIMAP4.UseTLS := utUseImplicitTLS;
          DoLog(sLineBreak + 'IMAP');

          DoLog('Start Connect Outlook');
          FIdIMAP4.Connect;
          try
            DoLog('Connected Outlook');
            FIdIMAP4.SelectMailBox('INBOX');

            DoLog(Format('Your Outlook TotalMsgs: %d', [FIdIMAP4.MailBox.TotalMsgs]));
          finally
            FIdIMAP4.Disconnect;
            DoLog('Disconnected Outlook');
          end;
        end
        else
          DoLog('Access Token is expired!!');
      end
      else
        DoLog('Access Token is empty!!');
    except
      on E: Exception do
      begin
        DoLog('IMAP Exception: ' + E.ToString);
      end;
    end;
  finally
    if TIdSASLXOAuth(FxOAuthSASLImap.SASL).IsTokenExpired then
      TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token := EmptyStr;
  end;
end;

procedure TFormIMAPTest.ConnectPop;
const
  ST_OK = '+OK';
  ST_SASLCONTINUE = '+';  {Do not translate}
begin
  try
    try
      if not TIdSASLXOAuth(FxOAuthSASLPop3.SASL).Token.IsEmpty then
      begin
        if not TIdSASLXOAuth(FxOAuthSASLPop3.SASL).IsTokenExpired then
        begin
          FIdPOP3.AuthType := patSASL;
          DoLog(sLineBreak + 'POP3');
          DoLog('Start Connect Outlook');
          FIdPOP3.Host := 'outlook.office365.com';
          FIdPOP3.UseTLS := utUseImplicitTLS;

          DoLog('Connect Outlook');
          FIdPOP3.Connect;
          try
            FIdPOP3.CAPA;
            DoLog('Login Outlook');
            FIdPOP3.SASLMechanisms.LoginSASL('AUTH', FIdPOP3.Host, 'pop', [ST_OK], [ST_SASLCONTINUE], FIdPOP3, FIdPOP3.Capabilities, 'SASL'); {do not localize}
            DoLog(Format('Your Outlook TotalMsgs: %d', [FIdPOP3.CheckMessages]));
          finally
            FIdPOP3.Disconnect;
            DoLog('Disconnected Outlook');
          end;
        end
        else
          DoLog('Access Token is expired!!');
      end
      else
        DoLog('Access Token is empty!!');
    except
      on E: Exception do
      begin
        DoLog('POP3 Exception: ' + E.ToString);
      end;
    end;
  finally
    if TIdSASLXOAuth(FxOAuthSASLPop3.SASL).IsTokenExpired then
      TIdSASLXOAuth(FxOAuthSASLPop3.SASL).Token := EmptyStr;
  end;
end;

procedure TFormIMAPTest.ConnectSmtp;
var
  IdMessage: TIdMessage;
begin
  try
    try
      if not TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token.IsEmpty then
      begin
        if not TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).IsTokenExpired then
        begin
          FIdSMTP.AuthType := satSASL;
          DoLog(sLineBreak + 'SMTP');
          DoLog('Start Connect Outlook');
          FIdSMTP.Host := 'outlook.office365.com';

          DoLog('Connect Outlook');
          FIdSMTP.Connect;
          try
            DoLog('Autenticate Outlook');
            FIdSMTP.Authenticate;

            IdMessage := TIdMessage.Create(Self);
            IdMessage.From.Address := edtEmailAccount.Text;
            IdMessage.From.Name := 'Test';
            IdMessage.ReplyTo.EMailAddresses := IdMessage.From.Address;
            IdMessage.Recipients.Add.Text := edtEmailAccount.Text;
            IdMessage.Subject := 'Hello World';
            IdMessage.Body.Text := 'Hello Body';

            FIdSMTP.Send(IdMessage);
            DoLog('Message Sent');
          finally
            FIdSMTP.Disconnect;
            DoLog('Disconnected Outlook');
          end;
        end
        else
          DoLog('Access Token is expired!!');
      end
      else
        DoLog('Access Token is empty!!');
    except
      on E: Exception do
      begin
        DoLog('POP3 Exception: ' + E.ToString);
      end;
    end;
  finally
    if TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).IsTokenExpired then
      TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token := EmptyStr;
  end;
end;

procedure TFormIMAPTest.CrateIdSSLIOHandlerSocketOpenSSLImap;
begin
  FIdSSLIOHandlerSocketOpenSSLImap := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
  FIdSSLIOHandlerSocketOpenSSLImap.Host := FHost;
  FIdSSLIOHandlerSocketOpenSSLImap.Port := 993;
  FIdSSLIOHandlerSocketOpenSSLImap.SSLOptions.Method := sslvTLSv1_2;
  FIdSSLIOHandlerSocketOpenSSLImap.SSLOptions.SSLVersions := [sslvTLSv1_2];
//  FIdSSLIOHandlerSocketOpenSSLImap.SSLOptions.Mode := sslmClient;
end;

procedure TFormIMAPTest.CrateIdSSLIOHandlerSocketOpenSSLSmtp;
begin
  FIdSSLIOHandlerSocketOpenSSLSmtp := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
  FIdSSLIOHandlerSocketOpenSSLSmtp.Host := FHost;
  FIdSSLIOHandlerSocketOpenSSLSmtp.Port := 587;
  FIdSSLIOHandlerSocketOpenSSLSmtp.SSLOptions.Method := sslvTLSv1_2;
  FIdSSLIOHandlerSocketOpenSSLSmtp.SSLOptions.SSLVersions := [sslvTLSv1_2];
  FIdSSLIOHandlerSocketOpenSSLSmtp.SSLOptions.Mode := sslmClient;
end;

procedure TFormIMAPTest.CrearteIdIMAP4;
begin
  FIdIMAP4 := TIdIMAP4.Create(Self);
  FIdIMAP4.IOHandler := FIdSSLIOHandlerSocketOpenSSLImap;
  FIdIMAP4.Host := FHost;
  FIdIMAP4.Port := 993;
  FIdIMAP4.UseTLS := utUseImplicitTLS;
  FIdIMAP4.AuthType := iatSASL;
  FIdIMAP4.MilliSecsToWaitToClearBuffer := 10;
  FIdIMAP4.OnStatus := StatusConnection;
end;

procedure TFormIMAPTest.CreateIdPOP3;
begin
  FIdPOP3 := TIdPOP3.Create(Self);
  FIdPOP3.IOHandler := FIdSSLIOHandlerSocketOpenSSLPop;
  FIdPOP3.UseTLS := utUseImplicitTLS;
  FIdPOP3.AuthType := patSASL;
  FIdPOP3.AutoLogin := False;
  FIdPOP3.OnStatus := StatusConnection;
end;

procedure TFormIMAPTest.CreateIdSMTP;
begin
  FIdSMTP := TIdSMTP.Create(Self);
  FIdSMTP.IOHandler := FIdSSLIOHandlerSocketOpenSSLSmtp;
  FIdSMTP.AuthType := satSASL;
  FIdSMTP.UseTLS := utUseRequireTLS;
  FIdSMTP.OnStatus := StatusConnection;
end;

procedure TFormIMAPTest.CreateIdSSLIOHandlerSocketOpenSSLPop;
begin
  FIdSSLIOHandlerSocketOpenSSLPop := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
  FIdSSLIOHandlerSocketOpenSSLPop.Host := FHost;
  FIdSSLIOHandlerSocketOpenSSLPop.Port := 995;
  FIdSSLIOHandlerSocketOpenSSLPop.SSLOptions.Method := sslvTLSv1_2;
  FIdSSLIOHandlerSocketOpenSSLPop.SSLOptions.SSLVersions := [sslvTLSv1_2];
  FIdSSLIOHandlerSocketOpenSSLPop.SSLOptions.Mode := sslmClient;
end;

procedure TFormIMAPTest.DoLog(const Text: string);
begin
  Memo1.Lines.Add(Text);
end;

procedure TFormIMAPTest.ErrorAccessToken(const Error, ErrorDescription: string);
begin
  DoLog('Error: ' + Error);
  DoLog('Error_Description: ' + ErrorDescription);
end;

procedure TFormIMAPTest.FormCreate(Sender: TObject);
begin
  edtTenantId.Text := TENANTID;
  edtClientId.Text := CLIENTID;
  edtClientSecret.Text := CLIENTSECRET;
  edtEmailAccount.Text := EMAILACCOUNT;
  edtEmailPassword.Text := EMAILPASSWORD;
  edtScope.Text := FScope;
  edtUrlToken.Text := FUrlToken;

  CrateIdSSLIOHandlerSocketOpenSSLImap;
  CrearteIdIMAP4;
  FxOAuthSASLImap := FIdIMAP4.SASLMechanisms.Add;
  FxOAuthSASLImap.SASL := TIdSASLXOAuth.Create(Self);

  CreateIdSSLIOHandlerSocketOpenSSLPop;
  CreateIdPOP3;
  FxOAuthSASLPop3 := FIdPOP3.SASLMechanisms.Add;
  FxOAuthSASLPop3.SASL := TIdSASLXOAuth.Create(Self);

  CrateIdSSLIOHandlerSocketOpenSSLSmtp;
  CreateIdSMTP;
  FxOAuthSASLSmtp := FIdSMTP.SASLMechanisms.Add;
  FxOAuthSASLSmtp.SASL := TIdSASLXOAuth.Create(Self);

  FRopcFlow := TropcFlow.Create;
end;

procedure TFormIMAPTest.FormDestroy(Sender: TObject);
begin
  FRopcFlow.Free;
end;

procedure TFormIMAPTest.StatusConnection(ASender: TObject; const AStatus: TIdStatus; const AStatusText: string);
begin
  DoLog(AStatusText);
end;

procedure TFormIMAPTest.OAuth2AuthorizationCodeAfterAccessToken(Sender: TObject; const Access_Token, Token_Type, Expires_In, Refresh_Token, Scope, RawParams:
  string; var Handled: Boolean);
begin
  DoLog('AccessToken: ' + Access_Token);
  DoLog('Token_Type: ' + Token_Type);
  DoLog('Expires_In: ' + Expires_In);
  DoLog('Refresh_Token: ' + Refresh_Token);
  DoLog('Scope: ' + Scope);

  TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token := Access_Token;
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).ExpireTime := Expires_In;
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).User := edtEmailAccount.Text; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).Token := Access_Token;
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).ExpireTime := Expires_In;
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).User := edtEmailAccount.Text; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token := Access_Token;
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).ExpireTime := Expires_In;
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).User := edtEmailAccount.Text; // outlook email account
end;

procedure TFormIMAPTest.OAuth2ClientCredentialsAfterAccessToken(const Sender: TObject; const AccessToken, TokenType, ExpiresIn, RefreshToken, Scope, RawParams:
  string; var Handled: Boolean);
begin
  DoLog('AccessToken: ' + AccessToken);
  DoLog('Token_Type: ' + TokenType);
  DoLog('Expires_In: ' + ExpiresIn);
  DoLog('Refresh_Token: ' + RefreshToken);
  DoLog('Scope: ' + Scope);

  TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token := AccessToken;
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).ExpireTime := ExpiresIn;
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).User := edtEmailAccount.Text; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).Token := AccessToken;
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).ExpireTime := ExpiresIn;
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).User := edtEmailAccount.Text; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token := AccessToken;
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).ExpireTime := ExpiresIn;
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).User := edtEmailAccount.Text; // outlook email account
end;

procedure TFormIMAPTest.SettingAuthentication;
begin
  FRopcFlow.TenantID := edtTenantId.Text;
  FRopcFlow.ClientID := edtClientId.Text;
  FRopcFlow.ClientSecret := edtClientSecret.Text;
  FRopcFlow.Scope := edtScope.Text;
  FRopcFlow.Username := edtEmailAccount.Text;
  FRopcFlow.Password := edtEmailPassword.Text;
  FRopcFlow.OnAfterAccessToken := AfterAccessToken;
  FRopcFlow.OnErrorAccessToken := ErrorAccessToken;
end;

end.

