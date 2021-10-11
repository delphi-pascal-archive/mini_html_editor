program Example;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  About in 'About.pas' {AboutForm};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TAboutForm, AboutForm);
  Application.Run;
end.
