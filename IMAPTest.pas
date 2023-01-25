unit IMAPTest;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls,
  FMX.ScrollBox, FMX.Memo,
  IdTCPClient, IdExplicitTLSClientServerBase, IdMessageClient, IdIMAP4,
  IdBaseComponent, IdComponent, IdIOHandler, IdIOHandlerSocket, IdSASLCollection,
  IdIOHandlerStack, IdSSL, IdSSLOpenSSL,
  IdTCPConnection,
  IdSASLXOAUTH,
  DeviceAuthFlow,
  ROPCFlow,
  Global, FMX.Memo.Types, Data.Bind.Components, Data.Bind.ObjectScope, REST.Client, REST.Authenticator.OAuth, IdHTTP, REST.Authenticator.Basic, REST.Types,
  IdPOP3, IdSMTPBase, IdSMTP;

type
  TFormIMAPTest = class(TForm)
    Memo1: TMemo;
    btn_Test_outlook_IMAP: TButton;
    IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;
    IdIMAP4: TIdIMAP4;
    Button1: TButton;
    btnPop: TButton;
    IdSSLIOHandlerSocketPOP: TIdSSLIOHandlerSocketOpenSSL;
    IdPOP3: TIdPOP3;
    btnSmtp: TButton;
    IdSSLIOHandlerSocketSMTP: TIdSSLIOHandlerSocketOpenSSL;
    IdSMTP1: TIdSMTP;
    procedure OAuth2_Authorization_CodeAfterAuthorizeCode(Sender: TObject;
      const Code, State, Scope, RawParams: string; var Handled: Boolean);
    procedure OAuth2_Authorization_CodeAfterAccessToken(Sender: TObject;
      const Access_Token, Token_Type, Expires_In, Refresh_Token, Scope,
      RawParams: string; var Handled: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure btn_Test_outlook_IMAPClick(Sender: TObject);
    procedure btn_Device_Auth_FlowClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnPopClick(Sender: TObject);
    procedure btnSmtpClick(Sender: TObject);
    procedure OAuth2_Client_CredentialsAfterAccessToken(Sender: TObject;
      const Access_Token, Token_Type, Expires_In, Refresh_Token, Scope,
      RawParams: string; var Handled: Boolean);
  private
    FxOAuthSASLImap: TIdSASLListEntry;
    FxOAuthSASLPop3: TIdSASLListEntry;
    FxOAuthSASLSmtp: TIdSASLListEntry;
    FDevice_Authorization_Flow: TDevice_Authorization_Flow;
    FROPC_Flow: TROPC_Flow;
    procedure DoLog(logText: string);
  end;

var
  FormIMAPTest: TFormIMAPTest;

implementation

uses
  IdSASL, IdMessage;


{$R *.fmx}

procedure TFormIMAPTest.btn_Device_Auth_FlowClick(Sender: TObject);
begin
  DoLog('Start Device Authorization Flow');
  FDevice_Authorization_Flow.Start;
end;

procedure TFormIMAPTest.btn_Test_outlook_IMAPClick(Sender: TObject);
begin
  try
    try
      if not TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token.IsEmpty then
      begin
        if not TIdSASLXOAuth(FxOAuthSASLImap.SASL).IsTokenExpired then
        begin
          IdIMAP4.UseTLS := utUseImplicitTLS;
          DoLog(sLineBreak + 'IMAP');

          DoLog('Start Connect Outlook');
          IdIMAP4.Connect;
          try
            DoLog('Connected Outlook');
            IdIMAP4.SelectMailBox('INBOX');

            DoLog('Your Outlook TotalMsgs: ' + IdIMAP4.MailBox.TotalMsgs.ToString);
          finally
            IdIMAP4.Disconnect;
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

procedure TFormIMAPTest.Button1Click(Sender: TObject);
begin
  DoLog('Start ROPC Flow');
  FROPC_Flow.Start;
end;

procedure TFormIMAPTest.btnPopClick(Sender: TObject);
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
          IdPOP3.AuthType := patSASL;
          DoLog(sLineBreak + 'POP3');
          DoLog('Start Connect Outlook');
          IdPOP3.Host := 'outlook.office365.com';
          IdPOP3.UseTLS := utUseImplicitTLS;

          DoLog('Connect Outlook');
          IdPOP3.Connect;
          try
            IdPOP3.CAPA;
            DoLog('Login Outlook');
            IdPOP3.SASLMechanisms.LoginSASL('AUTH', IdPOP3.Host, 'pop', [ST_OK], [ST_SASLCONTINUE], IdPOP3, IdPOP3.Capabilities, 'SASL', False); {do not localize}
            DoLog('Your Outlook TotalMsgs: ' + IdPOP3.CheckMessages.ToString);
          finally
            IdPOP3.Disconnect;
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

procedure TFormIMAPTest.btnSmtpClick(Sender: TObject);
var
  IdMessage: TIdMessage;
begin
  try
    try
      if not TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token.IsEmpty then
      begin
        if not TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).IsTokenExpired then
        begin
          IdSMTP1.AuthType := satSASL;
          DoLog(sLineBreak + 'SMTP');
          DoLog('Start Connect Outlook');
          IdSMTP1.Host := 'outlook.office365.com';

          DoLog('Connect Outlook');
          IdSMTP1.Connect;
          try
            DoLog('Autenticate Outlook');
            IdSMTP1.Authenticate;

            IdMessage := TIdMessage.Create(Self);
            IdMessage.From.Address := EMAILACCOUNT;
            IdMessage.From.Name := 'Test';
            IdMessage.ReplyTo.EMailAddresses := IdMessage.From.Address;
            IdMessage.Recipients.Add.Text := EMAILACCOUNT;
            IdMessage.Subject := 'Hello World';
            IdMessage.Body.Text := 'Hello Body';

            IdSMTP1.Send(IdMessage);
            DoLog('Message Sent');
          finally
            IdSMTP1.Disconnect;
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

procedure TFormIMAPTest.DoLog(logText: string);
begin
  Memo1.Lines.Add(logText);
end;

procedure TFormIMAPTest.FormCreate(Sender: TObject);
begin
  FxOAuthSASLImap := IdIMAP4.SASLMechanisms.Add;
  FxOAuthSASLImap.SASL := TIdSASLXOAuth.Create(Self);

  FxOAuthSASLPop3 := IdPOP3.SASLMechanisms.Add;
  FxOAuthSASLPop3.SASL := TIdSASLXOAuth.Create(Self);

  FxOAuthSASLSmtp := IdSMTP1.SASLMechanisms.Add;
  FxOAuthSASLSmtp.SASL := TIdSASLXOAuth.Create(Self);

  //Device Authorization Flow
  FDevice_Authorization_Flow := TDevice_Authorization_Flow.Create;
  FDevice_Authorization_Flow.TenantID := TENANTID;
  FDevice_Authorization_Flow.ClientID := CLIENTID;
  FDevice_Authorization_Flow.Scope := SCOPE;

  FDevice_Authorization_Flow.OnAfterAuthorizeCode :=
      procedure(AuthCode: string)
      begin
//        btn_Device_Auth_Flow.Enabled := False;
        DoLog('Your Auth Code: ' + AuthCode);
      end;

  FDevice_Authorization_Flow.OnAfterAuthorizeGetExpireTime :=
      procedure(ExpireTime: Integer)
      begin
//        btn_Device_Auth_Flow.Text := 'Use Device Authorization Flow - ' + IntToStr(ExpireTime) + 's';
      end;
  FDevice_Authorization_Flow.OnAfterAccessToken :=
      procedure(Device_ID, Access_Token, Token_Type: string; Expires_In: Integer; Scope: string)
      begin
//        btn_Device_Auth_Flow.Enabled := True;
//        btn_Device_Auth_Flow.Text := 'Use Device Authorization Flow';
        DoLog('Device_ID: ' + Device_ID + CRLF +
              'AccessToken: ' + Access_Token + CRLF +
              'Token_Type: ' + Token_Type + CRLF +
              'Expires_In: ' + IntToStr(Expires_In) + CRLF +
              'Scope: ' + Scope);
        TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token := Access_Token;
        TIdSASLXOAuth(FxOAuthSASLImap.SASL).ExpireTime := IntToStr(Expires_In);
        TIdSASLXOAuth(FxOAuthSASLImap.SASL).User := EMAILACCOUNT; // outlook email account

        TIdSASLXOAuth(FxOAuthSASLPop3.SASL).Token := Access_Token;
        TIdSASLXOAuth(FxOAuthSASLPop3.SASL).ExpireTime := IntToStr(Expires_In);
        TIdSASLXOAuth(FxOAuthSASLPop3.SASL).User := EMAILACCOUNT; // outlook email account

        TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token := Access_Token;
        TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).ExpireTime := IntToStr(Expires_In);
        TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).User := EMAILACCOUNT; // outlook email account
      end;

  FDevice_Authorization_Flow.OnErrorAccessToken :=
      procedure(Error, ErrorDescription: string)
      begin
//        btn_Device_Auth_Flow.Text := 'Use Device Authorization Flow';
        DoLog('Error: ' + Error + CRLF +
              'Error_Description: ' + ErrorDescription);
//        btn_Device_Auth_Flow.Enabled := True;
      end;

  //ROPC Flow
  FROPC_Flow := TROPC_Flow.Create;
  FROPC_Flow.TenantID := TENANTID;
  FROPC_Flow.ClientID := CLIENTID;
  FROPC_Flow.ClientSecret := CLIENTSECRET;
  FROPC_Flow.Scope := SCOPE;
  FROPC_Flow.Username := EMAILACCOUNT;
  FROPC_Flow.Password := EMAILPASSWORD;
  FROPC_Flow.OnAfterAccessToken :=
      procedure(Access_Token, Token_Type: string; Expires_In: Integer; Scope: string)
      begin
        DoLog('AccessToken: ' + Access_Token + CRLF +
              'Token_Type: ' + Token_Type + CRLF +
              'Expires_In: ' + IntToStr(Expires_In) + CRLF +
              'Scope: ' + Scope);
        TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token := Access_Token;
        TIdSASLXOAuth(FxOAuthSASLImap.SASL).ExpireTime := IntToStr(Expires_In);
        TIdSASLXOAuth(FxOAuthSASLImap.SASL).User := EMAILACCOUNT; // outlook email account

        TIdSASLXOAuth(FxOAuthSASLPop3.SASL).Token := Access_Token;
        TIdSASLXOAuth(FxOAuthSASLPop3.SASL).ExpireTime := IntToStr(Expires_In);
        TIdSASLXOAuth(FxOAuthSASLPop3.SASL).User := EMAILACCOUNT; // outlook email account

        TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token := Access_Token;
        TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).ExpireTime := IntToStr(Expires_In);
        TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).User := EMAILACCOUNT; // outlook email account
      end;

  FROPC_Flow.OnErrorAccessToken :=
      procedure(Error, ErrorDescription: string)
      begin
        DoLog('Error: ' + Error + CRLF + 'Error_Description: ' + ErrorDescription);
      end;
end;

procedure TFormIMAPTest.FormDestroy(Sender: TObject);
begin
  if Assigned(FDevice_Authorization_Flow) then
    FreeAndNil(FDevice_Authorization_Flow);

  if Assigned(FROPC_Flow) then
    FreeAndNil(FROPC_Flow);
end;

procedure TFormIMAPTest.OAuth2_Authorization_CodeAfterAccessToken(Sender: TObject;
  const Access_Token, Token_Type, Expires_In, Refresh_Token, Scope,
  RawParams: string; var Handled: Boolean);
begin
  DoLog('AccessToken: ' + Access_Token + CRLF +
        'Token_Type: ' + Token_Type + CRLF +
        'Expires_In: ' + Expires_In + CRLF +
        'Refresh_Token: ' + Refresh_Token + CRLF +
        'Scope: ' + Scope);
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token := Access_Token;
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).ExpireTime := Expires_In;
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).User := EMAILACCOUNT; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).Token := Access_Token;
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).ExpireTime := Expires_In;
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).User := EMAILACCOUNT; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token := Access_Token;
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).ExpireTime := Expires_In;
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).User := EMAILACCOUNT; // outlook email account
end;

procedure TFormIMAPTest.OAuth2_Authorization_CodeAfterAuthorizeCode(Sender: TObject;
  const Code, State, Scope, RawParams: string; var Handled: Boolean);
begin
  DoLog('Code: ' + Code + CRLF +
        'State: ' + State + CRLF +
        'Scope: ' + Scope);
end;

procedure TFormIMAPTest.OAuth2_Client_CredentialsAfterAccessToken(
  Sender: TObject; const Access_Token, Token_Type, Expires_In, Refresh_Token,
  Scope, RawParams: string; var Handled: Boolean);
begin
  DoLog('AccessToken: ' + Access_Token + CRLF +
        'Token_Type: ' + Token_Type + CRLF +
        'Expires_In: ' + Expires_In + CRLF +
        'Refresh_Token: ' + Refresh_Token + CRLF +
        'Scope: ' + Scope);

  TIdSASLXOAuth(FxOAuthSASLImap.SASL).Token := Access_Token;
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).ExpireTime := Expires_In;
  TIdSASLXOAuth(FxOAuthSASLImap.SASL).User := EMAILACCOUNT; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).Token := Access_Token;
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).ExpireTime := Expires_In;
  TIdSASLXOAuth(FxOAuthSASLPop3.SASL).User := EMAILACCOUNT; // outlook email account

  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).Token := Access_Token;
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).ExpireTime := Expires_In;
  TIdSASLXOAuth(FxOAuthSASLSmtp.SASL).User := EMAILACCOUNT; // outlook email account
end;

end.
