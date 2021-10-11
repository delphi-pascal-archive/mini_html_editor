unit About;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, MMSystem, IniFiles;

type
  TAboutForm = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Image2: TImage;
    NLabel: TLabel;
    procedure HoverOff(Sender: TObject);
    procedure HoverOn(Sender: TObject);
    procedure Image1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Label4Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure NLabelClick(Sender: TObject);
    procedure Image1DblClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private

    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutForm: TAboutForm;

implementation

{$R *.dfm}

const MidiFile = 'hakama.mid';

function FileVersion(AFileName: string): string; 
var
  szName: array[0..255] of Char;
  P: Pointer;
  Value: Pointer;
  Len: UINT;
  GetTranslationString: string;
  FFileName: PChar;
  FValid: boolean;
  FSize: DWORD;
  FHandle: DWORD;
  FBuffer: PChar;
begin
  try
    FFileName := StrPCopy(StrAlloc(Length(AFileName) + 1), AFileName);
    FValid := False;
    FSize := GetFileVersionInfoSize(FFileName, FHandle);
    if FSize > 0 then
  try
    GetMem(FBuffer, FSize);
    FValid := GetFileVersionInfo(FFileName, FHandle, FSize, FBuffer);
  except
    FValid := False;
  raise;
  end;
  Result := '';
  if FValid then
    VerQueryValue(FBuffer, '\VarFileInfo\Translation', p, Len)
  else
    p := nil;
  if P <> nil then
    GetTranslationString := IntToHex(MakeLong(HiWord(Longint(P^)),
    LoWord(Longint(P^))), 8);
  if FValid then
  begin
    StrPCopy(szName, '\StringFileInfo\' + GetTranslationString + '\FileVersion');
    if VerQueryValue(FBuffer, szName, Value, Len) then
    Result := StrPas(PChar(Value));
  end;
  finally
  try
    if FBuffer <> nil then
      FreeMem(FBuffer, FSize);
  except
  end;
  try
    StrDispose(FFileName);
  except
  end;
  end;
end;

// ==>> Заставляем About двигаться в положении BorderStyle - bsNone ==========>>
procedure TAboutForm.Image1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button=mbLeft then
  begin
    ReleaseCapture;
    Perform(WM_SYSCOMMAND, $F012, 0);
  end;
end;

procedure TAboutForm.HoverOn(Sender: TObject);
var
  I: integer;
begin
  for I:=0 to AboutForm.ComponentCount-1 do
    if ((Components[I] is TImage) and (Components[I].Tag=(Sender as TLabel).Tag)) then
      (Components[I] as TImage).Visible:=true;
end;

procedure TAboutForm.HoverOff(Sender: TObject);
var
  I: integer;
begin
  for I:=0 to AboutForm.ComponentCount-1 do
    if ((Components[I] is TImage) and (Components[I].Tag in [1..5])) then
      (Components[I] as TImage).Visible := false;
end;

procedure TAboutForm.Label4Click(Sender: TObject);
begin
  ShowMessage('Ну чего ты тыкаешь сюда? Не видишь кода нету :)');
end;

procedure TAboutForm.FormCreate(Sender: TObject);
begin
  Label1.Caption := 'ver ' + FileVersion(Paramstr(0));
  Image2.Picture.Icon := Application.Icon;
end;

procedure TAboutForm.NLabelClick(Sender: TObject);
begin
  Close;
end;

procedure TAboutForm.Image1DblClick(Sender: TObject);
begin
  Close;
end;

procedure TAboutForm.FormActivate(Sender: TObject);
begin
  MCISendString(PChar('play ' + MidiFile), nil, 0, 0);
end;

procedure TAboutForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  MCISendString(PChar('stop ' + MidiFile), nil, 0, 0);
end;

end.
