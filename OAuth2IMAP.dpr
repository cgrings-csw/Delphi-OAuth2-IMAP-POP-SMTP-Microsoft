program OAuth2IMAP;

uses
  VCL.Forms,
  IMAPTest in 'IMAPTest.pas' {FormIMAPTest},
  IdSASLXOAUTH in 'IdSASLXOAUTH.pas',
  ROPCFlow in 'ROPCFlow.pas',
  Global in 'Global.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFormIMAPTest, FormIMAPTest);
  Application.Run;
end.
