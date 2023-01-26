program OAuth2IMAP;

uses
  VCL.Forms,
  IMAPTest in 'IMAPTest.pas' {FormIMAPTest},
  IdSASLXOAUTH in 'IdSASLXOAUTH.pas',
  ROPCFlow in 'ROPCFlow.pas',
  Global in 'Global.pas',
  XSuperObject in 'x-superobject\XSuperObject.pas',
  XSuperJSON in 'x-superobject\XSuperJSON.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFormIMAPTest, FormIMAPTest);
  Application.Run;
end.
