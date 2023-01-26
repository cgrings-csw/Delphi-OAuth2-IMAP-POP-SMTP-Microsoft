unit ROPCFlow;

interface

uses
  Winapi.ShellAPI,
  System.Classes,
  System.SysUtils,
  IdHTTP,
  IdSSLOpenSSL;

type
  TOnErrorAccessToken = reference to procedure(const Error, ErrorDescription: string);
  TOnAfterAccessToken = reference to procedure(const AccessToken, TokenType: string; const ExpiresIn: Int32; const Scope: string);

  TropcFlow = class sealed(TObject)
  strict private
  const
    FFormatClientid: string = 'client_id=%s';
    FFormatClientSecret: string = 'client_secret=%s';
    FFormatGrantType: string = 'grant_type=password';
    FFormatPassword: string = 'password=%s';
    FFormatScope: string = 'scope=%s';
    FFormatTokeUrl: string = 'https://login.microsoftonline.com/%s/oauth2/v2.0/token';
    FFormatUserName: string = 'username=%s';
  var
    FClientID: string;
    FClientSecret: string;
    FExpireIn: Int32;
    FIdHTTP: TIdHTTP;
    FInterval: Int32;
    FOnAfterAccessToken: TOnAfterAccessToken;
    FOnErrorAccessToken: TOnErrorAccessToken;
    FPassword: string;
    FScope: string;
    FTenantID: string;
    FUsername: string;
    FVerification_URI: string;
    LHandler: TIdSSLIOHandlerSocketOpenSSL;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Start;
    property ClientID: string read FClientID write FClientID;
    property ClientSecret: string read FClientSecret write FClientSecret;
    property Password: string read FPassword write FPassword;
    property Scope: string read FScope write FScope;
    property TenantID: string read FTenantID write FTenantID;
    property Username: string read FUsername write FUsername;
    property OnAfterAccessToken: TOnAfterAccessToken read FOnAfterAccessToken write FOnAfterAccessToken;
    property OnErrorAccessToken: TOnErrorAccessToken read FOnErrorAccessToken write FOnErrorAccessToken;
  end;

implementation

uses
  XSuperObject;

constructor TropcFlow.Create;
begin
  LHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
  LHandler.SSLOptions.SSLVersions := [sslvTLSv1_2];
  LHandler.SSLOptions.Mode := sslmClient;
  LHandler.SSLOptions.VerifyMode := [];
  LHandler.SSLOptions.VerifyDepth := 0;

  FIdHTTP := TIdHTTP.Create(nil);
  FIdHTTP.IOHandler := LHandler;
  FIdHTTP.Request.ContentEncoding := 'UTF-8';
  FIdHTTP.Request.ContentType := 'application/x-www-form-urlencoded';
end;

destructor TropcFlow.Destroy;
begin
  LHandler.Free;
  FIdHTTP.Free;

  inherited;
end;

procedure TropcFlow.Start;
var
  FErrResponseJSON: ISuperObject;
  FResponseJSON: ISuperObject;
  FResponseString: string;
  postData: TStrings;
begin
  if (not FClientID.IsEmpty) and
     (not FClientSecret.IsEmpty) and
     (not FTenantID.IsEmpty) and
     (not FScope.IsEmpty) and
     (not FUsername.IsEmpty) and
     (not FPassword.IsEmpty)  then
  begin
    try
      try
        // Post Data
        postData := TStringList.Create;
        postData.Add(Format(FFormatClientid, [FClientID]));
        postData.Add(Format(FFormatClientSecret, [FClientSecret]));
        postData.Add(Format(FFormatScope, [FScope]));
        postData.Add(Format(FFormatUserName, [FUsername]));
        postData.Add(Format(FFormatPassword, [FPassword]));
        postData.Add(FFormatGrantType);
        // Call Device Auth API
        FResponseString := FIdHTTP.Post(Format(FFormatTokeUrl, [FTenantID]), postData);
        // Response JSON
        FResponseJSON := SO(FResponseString);
        // Callback Auth Code
        if Assigned(FOnAfterAccessToken) then
          FOnAfterAccessToken(FResponseJSON.S['access_token'],
                              FResponseJSON.S['token_type'],
                              FResponseJSON.I['expires_in'],
                              FResponseJSON.S['scope']);
      except
        on E: EIdHTTPProtocolException do
        begin
          // Http Error
          FErrResponseJSON := SO(E.ErrorMessage);

          if Assigned(OnErrorAccessToken) then
            OnErrorAccessToken(FResponseJSON.S['error'], FResponseJSON.S['error_description']);
        end;
      end;
    finally
      if Assigned(postData) then
        postData.Free;
    end;
  end
  else
  begin
    raise Exception.Create('Not set Client ID or Client Secret or Tenant ID or Scope or Username or Password');
  end;
end;

end.
