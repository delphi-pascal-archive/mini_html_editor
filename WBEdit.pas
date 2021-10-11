{ *****************************************************************************
  ->> Библиотека функций для работы с TWebBrowser в режиме редактирования
  
  ->> Описание: 
      Модуль содержит в себе "инструменты" для быстрого доступа к функциям
      TWebBrowser в режиме редактора, а именно:
      
      * Позволяет редактировать текст HTML документа, не прибегая к 
        редактированию самого HTML кода (Визуальный редактор);
      * Имеется возможность менять цвет объектов, размер и начертание шрифта;
      * Вставка объектов, спец. символов.
  
  unit ver 1.0

  Автор: Kordal, kordall@mail.ru, icq 8281400
  Поддержка: Maniak, infinitykornets@gmail.com
  Copyright © 2007 by Localserver Software

  Отдельная благодарность Samum, автору статьи 
  "Визуальный HTML редактор своими руками II", которая и послужила толчком в
  написании данного модуля. 
 ***************************************************************************** }
unit WBEdit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls,
  OleCtrls, SHDocVw, ActiveX, MSHTML;

type
  TWBEdit = class
  private
    NumKey        : Char;

    EDInputQuery  : TForm;
    EDFontDialog  : TFontDialog;
    EDColorDialog : TColorDialog;
    EDPanel       : TPanel;
    EDComboBox    : TComboBox;
    EDLabels      : array[1..4] of TLabel;
    EDEdits       : array[1..4] of TEdit;
    EDButtons     : array[1..2] of TButton;

    CurrentWB     : IWebBrowser;
    Editor        : IHTMLDocument2;

    procedure ShowEdDialog(DlgType: Word);
    { Events }
    procedure RTButtOnClick(Sender: TObject);
    procedure RTEditKeyPress(Sender: TObject; var Key: Char);
    procedure RTEditOnChange(Sender: TObject);
    procedure RTEditKeyPressNIL(Sender: TObject; var Key: Char);
    procedure RTEditOnChangeNIL(Sender: TObject);
    { Tools }
     function Replace(Str, S1, S2: String): String;
    procedure FileReplaceString(const FileName, searchstring, replacestring: string);
    procedure PostKeyEx32(key: Word; const shift: TShiftState; specialkey: Boolean);
     function RGBtoHTMLColor(cl: TColor; ResultType: Word): String;
    { Ets }
    //procedure EDSetFocus(WB: TWebBrowser);
    procedure EDLoadHTMLsource(WB: TWebBrowser; HTMLCode: String);
     function EDGetHTMLsource(const FileName: string; WB: TWebBrowser; Filter: Boolean): String;
    procedure EDExecCommand(Command: WideString; ShowUI: WordBool; Value: OleVariant);
    procedure EDInsertHTMLcode(HTMLcode: WideString);
    procedure EDDesignMode(Mode: Boolean);
    procedure EDStyle(WB: TWebBrowser{; Zoom: byte});
  public
    Disp: IDispatch;

    constructor Create;
     destructor Destroy; override;

    procedure Edit_Delete;
    procedure Edit_Cut;
    procedure Edit_Copy;
    procedure Edit_Paste;
    procedure CreateLink;
    procedure SelectAll;
    procedure FindText;
    procedure FormatRemove;
    procedure FormatColor(Color: TColor);
    procedure FormatColorDialog;
    procedure FormatFont(FName, FStyle: String; FColor: TColor; FSize: byte; ClearFormat: Boolean);
    procedure FormatFontDialog;
    procedure FormatFontName(FontName: String);
    procedure FormatFontSize;
    procedure FormatAlignLeft;
    procedure FormatAlignCenter;
    procedure FormatAlignRight;
    procedure FormatBold;
    procedure FormatItalic;
    procedure FormatUnderline;
    procedure FormatSortList;
    procedure InsertHTMLcode(Code: WideString);
    procedure InsertImage;
    procedure Insert_2BR;
    procedure Insert_HR;
    procedure Insert_SUB;
    procedure Insert_SUP;
    procedure Insert_copy;
    procedure Insert_reg;
    procedure Insert_nbsp;
    procedure Insert_tm;

    procedure LoadHTMLsrc(WB: TWebBrowser; Code: String);
     function GetHTMLsrc(const FileName: string; WB: TWebBrowser; Filter: Boolean): String;
    procedure DesignMode(Mode: Boolean);
    procedure Style(WB: TWebBrowser);
  end;

implementation

const
  IS_COLOR = 1;
  IS_BACKCOLOR = 2;
  IS_CODECOLOR = 3;

  IS_FONTSIZE = 1;
  IS_SUB = 2;
  IS_SUP =3;
  IS_HYPERLINK = 4;
  IS_IMAGE = 5;

var
  IS_EDDIALOG: String[15];


{ Конструктор, деструктор }


{ *****************************************************************************
  ->> Диалог выбора параметров
 ***************************************************************************** }
constructor TWBEdit.Create;
var
  i: byte;
begin
  inherited Create;
  { Форма ShowEdDialog }
  EDInputQuery := TForm.Create(Application);
  with EDInputQuery do
  try
    Canvas.Font  := Font;
    ShowHint     := True;
    BorderStyle  := bsSingle;
    BorderIcons  := [biSystemMenu];
    FormStyle    := fsStayOnTop;
    Height       := {196; } 115;
    Width        := 350;
    Left         := (Screen.Width - Width) div 2;
    Top          := (Screen.Height - Height) div 2;

    { Dialogs }
    EDFontDialog := TFontDialog.Create(EDInputQuery);
    EDColorDialog := TColorDialog.Create(EDInputQuery);

    { Panel }
    EDPanel  := TPanel.Create(EDInputQuery);
    with EDPanel do
    begin
      Parent  := EDInputQuery;
      Left    := 0;
      Top     := 0;
      Width   := 350;
      Height  := {120;} 40;
      Visible := True;
    end;

    { Labels }
    for i := 1 to 4 do
    begin
      EDLabels[i] := TLabel.Create(EDPanel);
      with EDLabels[i] do
      begin
        Parent    := EDPanel;
        Font.Size := 10;
        Left      := 10;
        Top       := (EDPanel.Height div 5) * i-8;
        Tag       := i;
        Visible   := True;
      end;
    end;

    { ComboBox }
    EDComboBox := TComboBox.Create(EDPanel);
    with EDComboBox do
    begin
      Parent     := EDPanel;
      Left       := 100;
      Top        := (EDPanel.Height div 5) * 3-10;
      Tag        := 0;
      Width      := 233;
      Visible    := False;
    end;

    { Edits }
    for i := 1 to 4 do
    begin
      EDEdits[i] := TEdit.Create(EDPanel);
      with EDEdits[i] do
      begin
        Parent     := EDPanel;
        Left       := 100;
        Top        := (EDPanel.Height div 5) * i-10;
        Tag        := i;
        Width      := 233;
        Visible    := True;
      end;
    end;

    { Buttons }
    for i := 1 to 2 do
    begin
      EDButtons[i] := TButton.Create(EDInputQuery);
      with EDButtons[i] do
      begin
        Parent  := EDInputQuery;
        Top     := {130;} 50;
        Tag     := i;
        Width   := 120;
        OnClick := RTButtonClick;
        Visible := True;
      end;
    end;
    EDButtons[1].Caption := 'Применить';
    EDButtons[1].Left := (EDInputQuery.Width - EDButtons[1].Width * 2) div 3;
    EDButtons[2].Caption := 'Отмена';
    EDButtons[2].Left := EDButtons[1].Left * 2 + EDButtons[1].Width;
  finally

  end;
end;

destructor TWBEdit.Destroy;
begin
  inherited Destroy;
  //
end;


{ Внутренние функции }


procedure TWBEdit.ShowEdDialog(DlgType: Word);
  { IS_FONTSIZE }
  procedure ShowFontSizeDialog;
  var
    i: byte;
  begin
    IS_EDDIALOG := 'IS_FONTSIZE';

    EDInputQuery.Height := 115;
    EDInputQuery.Caption := 'Укажите размер шрифта';
    EDPanel.Height := 40;
    // Label
    EDLabels[1].Show;
    EDLabels[1].Caption := 'Значение:';
    EDLabels[1].Top := (EDPanel.Height div 2) - 9;
    for i:=2 to 4 do EDLabels[i].Hide;
    // Edit
    EDEdits[1].Clear;
    EDEdits[1].Show;
    EDEdits[1].OnKeyPress := RTEditKeyPress;
    EDEdits[1].OnChange   := RTEditOnChange;
    EDEdits[1].Top := (EDPanel.Height div 2) - 11;
    EDEdits[1].Hint := 'Может быть целочисленным от 1-7.';
    for i:=2 to 4 do EDEdits[i].Hide;
    // ComboBox
    EDComboBox.Hide;
    // Button
    EDButtons[1].Top := 50;
    EDButtons[2].Top := 50;
    // Show Input Dialog
    EDInputQuery.Show;
  end;

  { IS_SUB }
  procedure ShowSubDialog;
  var
    i: byte;
  begin
    IS_EDDIALOG := 'IS_SUB';

    EDInputQuery.Height := 115;
    EDInputQuery.Caption := 'Укажите нижний индекс';
    EDPanel.Height := 40;
    // Label
    EDLabels[1].Show;
    EDLabels[1].Caption := 'Значение:';
    EDLabels[1].Top := (EDPanel.Height div 2) - 9;
    for i:=2 to 4 do EDLabels[i].Hide;
    // Edit
    EDEdits[1].Show;
    EDEdits[1].OnKeyPress := RTEditKeyPressNIL;
    EDEdits[1].OnChange   := RTEditOnChangeNIL;
    EDEdits[1].Top := (EDPanel.Height div 2) - 11;
    EDEdits[1].Hint := 'Может содержать любое слово.';
    EDEdits[1].Clear;
    for i:=2 to 4 do EDEdits[i].Hide;
    // ComboBox
    EDComboBox.Hide;
    // Button
    EDButtons[1].Top := 50;
    EDButtons[2].Top := 50;
    // Show Input Dialog
    EDInputQuery.Show;
  end;

  { IS_SUP }
  procedure ShowSupDialog;
  var
    i: byte;
  begin
    IS_EDDIALOG := 'IS_SUP';

    EDInputQuery.Height := 115;
    EDInputQuery.Caption := 'Укажите верхний индекс';
    EDPanel.Height := 40;
    // Label
    EDLabels[1].Show;
    EDLabels[1].Caption := 'Значение:';
    EDLabels[1].Top := (EDPanel.Height div 2) - 9;
    for i:=2 to 4 do EDLabels[i].Hide;
    // Edit
    EDEdits[1].Show;
    EDEdits[1].OnKeyPress := RTEditKeyPressNIL; // пустышка
    EDEdits[1].OnChange   := RTEditOnChangeNIL; // пустышка
    EDEdits[1].Top := (EDPanel.Height div 2) - 11;
    EDEdits[1].Hint := 'Может содержать любое слово.';
    EDEdits[1].Clear;
    for i:=2 to 4 do EDEdits[i].Hide;
    // ComboBox
    EDComboBox.Hide;
    // Button
    EDButtons[1].Top := 50;
    EDButtons[2].Top := 50;
    // Show Input Dialog
    EDInputQuery.Show;
  end;

  { IS_HYPERLINK }
  procedure ShowHyperlinkDialog;
  var
  i: byte;
  begin
    IS_EDDIALOG := 'IS_HYPERLINK';

    EDInputQuery.Height := 196;
    EDInputQuery.Caption := 'Создать ссылку';
    //  Panel
    EDPanel.Height := 120;
    //  Label
    EDLabels[1].Caption := 'Текст:';
    EDLabels[2].Caption := 'Ссылка:';
    EDLabels[3].Caption := 'В окне:';
    EDLabels[4].Caption := 'Заголовок:';
    for i := 1 to 4 do
    begin
      EDLabels[i].Top := (EDPanel.Height div 5) * i-8;
      EDLabels[i].Show;
    end;
    //  Edit
    EDEdits[1].OnKeyPress := RTEditKeyPressNIL;
    EDEdits[1].OnChange   := RTEditOnChangeNIL;
    EDEdits[1].Top := (EDPanel.Height div 5) -   10;
    EDEdits[2].Top := (EDPanel.Height div 5) * 2-10;
    EDEdits[4].Top := (EDPanel.Height div 5) * 4-10;
    EDEdits[1].Clear;
    EDEdits[1].Show;
    EDEdits[2].Clear;
    EDEdits[2].Show;
    EDEdits[4].Clear;
    EDEdits[4].Show;
    EDEdits[2].Text := 'http://';
    // ComboBox
    EDComboBox.Top := (EDPanel.Height div 5) * 3-10;
    EDComboBox.Items.Clear;
    EDComboBox.Items.Add('');
    EDComboBox.Items.Add('_blank');
    EDComboBox.Items.Add('_parent');
    EDComboBox.Items.Add('_self');
    EDComboBox.Items.Add('_top');
    EDComboBox.Show;
    //  Button
    EDButtons[1].Top := 130;
    EDButtons[2].Top := 130;
    //  Show Input Dialog
    EDInputQuery.Show;
  end;

  { IS_IMAGE }
  procedure ShowImageDialog;
  var
  i: byte;
  begin
    IS_EDDIALOG := 'IS_IMAGE';

    EDInputQuery.Height := 196;
    EDInputQuery.Caption := 'Вставить картинку';
    //  Panel
    EDPanel.Height := 120;
    //  Label
    EDLabels[1].Caption := 'Ссылка:';
    EDLabels[2].Caption := 'Ширина:';
    EDLabels[3].Caption := 'Высота:';
    EDLabels[4].Caption := 'Подсказка:';
    for i := 1 to 4 do
    begin
      EDLabels[i].Top := (EDPanel.Height div 5) * i-8;
      EDLabels[i].Show;
    end;
    //  Edit
    EDEdits[1].OnKeyPress := RTEditKeyPressNIL;
    EDEdits[1].OnChange   := RTEditOnChangeNIL;
    for i := 1 to 4 do
    begin
      EDEdits[i].Top := (EDPanel.Height div 5) * i-10;
      EDEdits[i].Clear;
      EDEdits[i].Show;
    end;
    EDEdits[1].Text := 'http://';
    // ComboBox
    EDComboBox.Hide;
    //  Button
    EDButtons[1].Top := 130;
    EDButtons[2].Top := 130;
    //  Show Input Dialog
    EDInputQuery.Show;
  end;

begin
  case DlgType of
    1: ShowFontSizeDialog;
    2: ShowSubDialog;
    3: ShowSupDialog;
    4: ShowHyperlinkDialog;
    5: ShowImageDialog;
  end;
end;


{ *****************************************************************************
  ->> События
 ***************************************************************************** }
procedure TWBEdit.RTButtonClick(Sender: TObject);
begin
  case (Sender as TButton).Tag of
    { Кнопка "Применить"  }
    1: begin
         if length(EDEdits[1].Text) < 1 then Exit;
         if IS_EDDIALOG = 'IS_FONTSIZE' then
           EDExecCommand('FontSize', false, EDEdits[1].Text);

         if IS_EDDIALOG = 'IS_SUB' then
           EDInsertHTMLcode('<sub>' + EDEdits[1].Text + '</sub>');

         if IS_EDDIALOG = 'IS_SUP' then
           EDInsertHTMLcode('<sup>' + EDEdits[1].Text + '</sup>');

         if IS_EDDIALOG = 'IS_HYPERLINK' then
             ;

         if IS_EDDIALOG = 'IS_IMAGE' then
             ;
         EDInputQuery.Hide;
       end;
    { Кнопка "Отмена" }
    2: EDInputQuery.Hide;
  end;
end;

procedure TWBEdit.RTEditKeyPress(Sender: TObject; var Key: Char);
begin
  case (Sender as TEdit).Tag of
    1: case key of
         '1'..'7': NumKey := key;
       end;
  end;
end;

procedure TWBEdit.RTEditOnChange(Sender: TObject);
begin
  case (Sender as TEdit).Tag of
    1: if NumKey <> EDEdits[1].Text then
       begin
         EDEdits[1].Clear;
         Beep;
       end;
  end;
end;

procedure TWBEdit.RTEditKeyPressNil(Sender: TObject; var Key: Char);
begin
end;

procedure TWBEdit.RTEditOnChangeNil(Sender: TObject);
begin
end;


{ *****************************************************************************
  =>> Замена текста
 ***************************************************************************** }
function TWBEdit.Replace(Str, S1, S2: String): String;
var
  i: integer;
  s,t: String;
begin
  s := '';
  t := str;
  repeat
    i := pos(AnsiLowerCase(s1), AnsiLowerCase(t));
    if i>0 then begin
      s := s+Copy(t,1,i-1)+s2;
      t := Copy(t, i+Length(s1), MaxInt);
    end else s := s+t;
  until i<=0;
  result := s;
end;


procedure TWBEdit.FileReplaceString(const FileName, searchstring, replacestring: string);
var 
  fs: TFileStream;
  S: string;
begin 
  fs := TFileStream.Create(FileName, fmOpenread or fmShareDenyNone); 
  try 
    SetLength(S, fs.Size); 
    fs.ReadBuffer(S[1], fs.Size); 
  finally 
    fs.Free; 
  end; 
  S  := StringReplace(S, SearchString, replaceString, [rfReplaceAll, rfIgnoreCase]); 
  fs := TFileStream.Create(FileName, fmCreate); 
  try 
    fs.WriteBuffer(S[1], Length(S)); 
  finally 
    fs.Free; 
  end; 
end;


{ *****************************************************************************
  =>> Симуляция нажатия клавиш

 Прараметры:
  key       : Виртуальный код клавиши в ANSI коде (Ord(character)).

  shift     : Клавишы модификаторы (shift, control, alt, mouse buttons)
              Тип TShiftState продекларирован в Classes Unit.

  specialkey: Обычно имеет значение False. Устанавливается в true
              при использовании цифровой клавиатуры.
****************************************************************************** }
procedure TWBEdit.PostKeyEx32(key: Word; const shift: TShiftState; specialkey: Boolean);
type
  TShiftKeyInfo = record
    shift: Byte;
    vkey: Byte;
  end;
  byteset = set of 0..7;
const
  shiftkeys: array [1..3] of TShiftKeyInfo =(
  (shift: Ord(ssCtrl);  vkey: VK_CONTROL),
  (shift: Ord(ssShift); vkey: VK_SHIFT),
  (shift: Ord(ssAlt);   vkey: VK_MENU) );
var
  flag: DWORD;
  bShift: ByteSet absolute shift;
  i: Integer;
begin
  for i := 1 to 3 do
    begin
      if shiftkeys[i].shift in bShift then
       keybd_event(shiftkeys[i].vkey, MapVirtualKey(shiftkeys[i].vkey, 0), 0, 0);
  end; { For }
  if specialkey then
    flag := KEYEVENTF_EXTENDEDKEY
  else
    flag := 0;
  keybd_event(key, MapvirtualKey(key, 0), flag, 0);
  flag := flag or KEYEVENTF_KEYUP;
  keybd_event(key, MapvirtualKey(key, 0), flag, 0);
  for i := 3 downto 1 do
  begin
    if shiftkeys[i].shift in bShift then
      keybd_event(shiftkeys[i].vkey, MapVirtualKey(shiftkeys[i].vkey, 0),
      KEYEVENTF_KEYUP, 0);
  end; { For }
end;


{ *****************************************************************************
  =>> Функция преобразования TColor в HTMLcolor
 ***************************************************************************** }
function TWBEdit.RGBtoHTMLColor(cl: TColor; ResultType: Word): string;
var
  rgbColor: TColorRef;
  codeColor: String[6];
begin
  rgbColor := ColorToRGB(cl);
  codeColor := Format('%.2x%.2x%.2x',[GetRValue(rgbColor),
                                      GetGValue(rgbColor),
                                      GetBValue(rgbColor)]);
  case ResultType of
    1: Result := 'color="#' + codeColor + '"';
    2: Result := 'bgcolor="#' + codeColor + '"';
    3: Result := '#' + codeColor;
  end;
end;


{ *****************************************************************************
  =>> Передача фокуса ввода
 ***************************************************************************** }
{procedure TWBEdit.EDSetFocus(WB: TWebBrowser);
begin
  repeat
    Application.ProcessMessages;
  until
    WB.ReadyState >= READYSTATE_COMPLETE;
  if WB.Document <> nil then
   (WB.Document as IHTMLDocument2).ParentWindow.Focus;
end;


{ *****************************************************************************
  =>> Загрузить HTML исходник
 ***************************************************************************** }
procedure TWBEdit.EDLoadHTMLsource(WB: TWebBrowser; HTMLCode: String);
var
  sl: TStringList;
  ms: TMemoryStream;
begin
  WB.Navigate('about:blank');
  while WB.ReadyState < READYSTATE_INTERACTIVE do
    Application.ProcessMessages;

  if Assigned(WB.Document) then
  begin
    sl := TStringList.Create;
    try
      ms := TMemoryStream.Create;
      try
        sl.Text := HTMLCode;
        sl.SaveToStream(ms);
        ms.Seek(0, 0);
        (WB.Document as IPersistStreamInit).Load(TStreamAdapter.Create(ms));
      finally
        ms.Free;
      end;
    finally
      sl.Free;
    end;
  end;
end;


{ *****************************************************************************
  =>> Получить HTML исходник
 ***************************************************************************** }
function TWBEdit.EDGetHTMLsource(const FileName: String; WB: TWebBrowser; Filter: Boolean): String;
var
  PersistStream: IPersistStreamInit;
  FileStream: TFileStream;
  Stream: IStream;
  SaveResult: HRESULT;

  function GetFormattedHTMLcodeFromFile(AFile: String): String;
  var
    sl: TStringList;
    i: Integer;
  begin
    Result := '';
    if not FileExists(AFile) then Exit;
    FileReplaceString(AFile, '<BODY>', '');
    //FileReplaceString(AFile, 'ZOOM: 0.9;', 'ZOOM: 1.0;');
    FileReplaceString(AFile, '</BODY></HTML>', '');
    sl := TStringList.Create;
    try
      sl.LoadFromFile(AFile);
      //sl.Insert(5, '<!-- Generated program by Localserver software, © 2007. -->');
      for i:= 4 to sl.Count-1 do
        Result := Result + sl.Strings[i];
    finally
      sl.Free;
    end;
  end;

  function GetHTMLcodeFromFile(AFile: String): String;
  var
    sl: TStringList;
    i: Integer;
  begin
    sl := TStringList.Create;
    try
      sl.LoadFromFile(AFile);
      //sl.Insert(5, '<!-- Generated program by Localserver software, © 2007. -->');
      for i:= 0 to sl.Count-1 do
        Result := Result + sl.Strings[i];
    finally
      sl.Free;
    end;
  end;

begin
  PersistStream := WB.Document as IPersistStreamInit;
  FileStream := TFileStream.Create(FileName, fmCreate);
  try
    Stream := TStreamAdapter.Create(FileStream, soReference) as IStream;
    SaveResult := PersistStream.Save(Stream, True);
    if FAILED(SaveResult) then
      raise Exception.Create('Ошибка при сохранении HTML кода!');
  finally
    { В ответ на уничтожение объекта TFileStream, передаём
    soReference в конструктор TStreamAdapter. }
    FileStream.Free;
    if Filter = true then // фильтруем теги html, body
      Result := GetFormattedHTMLcodeFromFile(FileName)
    else
      Result := GetHTMLcodeFromFile(FileName);
  end;
end;


{ *****************************************************************************
  =>> Режим редактирования
 ***************************************************************************** }
procedure TWBEdit.EDDesignMode(Mode: Boolean);
begin
  CurrentWB := Disp as IWebBrowser;
  Editor:=(CurrentWB.Document as IHTMLDocument2);
  if Mode then
    Editor.DesignMode := 'On'
  else
    Editor.DesignMode := 'Off';
end;


{ *****************************************************************************
  =>> Интерпритатор команд
 ***************************************************************************** }
procedure TWBEdit.EDExecCommand(Command: WideString; ShowUI: WordBool; Value: OleVariant);
var
  CtrlRange: IHTMLControlRange;
  TextRange: IHTMLTxtRange;
begin
  if editor.selection.type_='Control' then
    begin
      CtrlRange:=(editor.selection.createRange as IHTMLControlRange);
      if not CtrlRange.queryCommandEnabled(Command) then
        Application.MessageBox('Не поддерживается!','')
      else
        CtrlRange.execCommand(Command, ShowUI, Value) end
  else
    begin
      TextRange:=(editor.selection.createRange as IHTMLTxtRange);
      TextRange.execCommand(Command, ShowUI, Value)
    end;
end;


{ *****************************************************************************
  =>> Вставить HTML код
 ***************************************************************************** }
procedure TWBEdit.EDInsertHTMLcode(HTMLcode: WideString);
var
  Range: IHTMLTxtRange;
begin
  Range:=(editor.selection.createRange as IHTMLTxtRange);
  Range.pasteHTML(HTMLcode);
end;


{ *****************************************************************************
  =>> Стиль TWebBrowser (ScrollBar, Zoom)
 ***************************************************************************** }
procedure TWBEdit.EDStyle(WB: TWebBrowser{, Zoom: byte});
begin
  with WB do
  begin
    OleObject.document.body.Style.scrollbarArrowColor := '#CDC9C9';
    OleObject.document.body.Style.scrollbar3DLIGHTCOLOR := '#EEE9E9';
    OleObject.document.body.Style.scrollbarDarkShadowColor := '#FFFFFF';
    OleObject.document.body.Style.scrollbarFaceColor := '#FFFFFF';
    OleObject.document.body.Style.scrollbarHighlightColor := '#FFFFFF';
    OleObject.Document.body.Style.scrollbarShadowColor := '#EEE9E9';
    OleObject.Document.body.Style.scrollbarTrackColor := '#FFFFFF';
    OleObject.Document.Body.Style.Zoom := 0.90;
    //EditDescription.OleObject.Document.Body.Style.OverflowX := 'hidden';
    //EditDescription.OleObject.Document.Body.Style.OverflowY := 'hidden';
  end;
end;


{ Импортируемые функции }


{ *****************************************************************************
  ->> Удалить
 ***************************************************************************** }
procedure TWBedit.Edit_Delete;
begin
  EDExecCommand('Delete', false, emptyparam);
end;


{ *****************************************************************************
  ->> Вырезать
 ***************************************************************************** }
procedure TWBedit.Edit_Cut;
begin
  EDExecCommand('Cut', false, emptyparam);
  // Посылаем комбинацию клавиш Ctrl + X
  //PostKeyEx32(Ord('X'), [ssctrl], False);
end;


{ *****************************************************************************
  ->> Копировать
 ***************************************************************************** }
procedure TWBedit.Edit_Copy;
begin
  EDExecCommand('Copy', false, emptyparam);
  // Посылаем комбинацию клавиш Ctrl + C
  //PostKeyEx32(Ord('C'), [ssctrl], False);
end;


{ *****************************************************************************
  ->> Вырезать
 ***************************************************************************** }
procedure TWBedit.Edit_Paste;
begin
  EDExecCommand('Paste', false, emptyparam);
  // Посылаем комбинацию клавиш Ctrl + V
  //PostKeyEx32(Ord('V'), [ssctrl], False);
end;


{ *****************************************************************************
  ->> Создать гиперссылку
 ***************************************************************************** }
procedure TWBEdit.CreateLink;
begin
  ShowEdDialog(IS_HYPERLINK);
  //EDExecCommand('CreateLink', false, emptyparam);
end;


{ *****************************************************************************
  ->> Выделить всё
 ***************************************************************************** }
procedure TWBEdit.SelectAll;
begin
  //EDSetFocus(WB);
  // Посылаем комбинацию клавиш Ctrl + A
  PostKeyEx32(Ord('A'), [ssctrl], False);
end;


{ *****************************************************************************
  ->> Найти текст
 ***************************************************************************** }
procedure TWBEdit.FindText;
begin
  // Посылаем комбинацию клавиш Ctrl + F
  PostKeyEx32(Ord('F'), [ssctrl], False);
end;


{ *****************************************************************************
  ->> Равнение по левому краю
 ***************************************************************************** }
procedure TWBEdit.FormatAlignLeft;
begin
  EDExecCommand('JustifyLeft', false, emptyparam);
end;


{ *****************************************************************************
  ->> Равнение по центру
 ***************************************************************************** }
procedure TWBEdit.FormatAlignCenter;
begin
  EDExecCommand('JustifyCenter', false, emptyparam);
end;


{ *****************************************************************************
  ->> Равнение по правому краю
 ***************************************************************************** }
procedure TWBEdit.FormatAlignRight;
begin
  EDExecCommand('JustifyRight', false, emptyparam);
end;


{ *****************************************************************************
  ->> Очистить форматирование
 ***************************************************************************** }
procedure TWBEdit.FormatRemove;
begin
  EDExecCommand('RemoveFormat', false, emptyparam);
end;


{ *****************************************************************************
  ->> Задать шрифт
 ***************************************************************************** }
procedure TWBEdit.FormatFontDialog;
begin
  if EDFontDialog.Execute then with EDFontDialog do
  begin
    //EDExecCommand('RemoveFormat', false, emptyparam); // обычный

    if (Font.Style = [fsBold]) then  // жирный
      EDExecCommand('Bold', false, emptyparam);

    if (Font.Style = [fsItalic]) then // курсив
      EDExecCommand('Italic', false, emptyparam);

    if (Font.Style = [fsBold, fsItalic]) then  // жирный курсив
      begin
        EDExecCommand('Bold', false, emptyparam);
        EDExecCommand('Italic', false, emptyparam);
      end;

    if (Font.Style = [fsUnderline]) then  // обычный, подчёркнутый
      EDExecCommand('Underline', false, emptyparam);

    if (Font.Style = [fsBold, fsUnderline]) then  // жирный, подчёркнутый
      begin
        EDExecCommand('Bold', false, emptyparam);
        EDExecCommand('Underline', false, emptyparam);
      end;

    if (Font.Style = [fsItalic, fsUnderline]) then // курсив, подчёркнутый
      begin
        EDExecCommand('Italic', false, emptyparam);
        EDExecCommand('Underline', false, emptyparam);
      end;

    if (Font.Style = [fsBold, fsItalic, fsUnderline]) then  // жирный курсив, подчёркнутый
      begin
        EDExecCommand('Bold', false, emptyparam);
        EDExecCommand('Italic', false, emptyparam);
        EDExecCommand('Underline', false, emptyparam);
      end;

    EDExecCommand('FontName',  false, 'face="'+ Font.Name +','+ Font.Name +','+ Font.Name +'"');
    EDExecCommand('FontSize',  false, Font.Size div 3);
    EDExecCommand('ForeColor', false, RGBtoHTMLColor(Font.Color, IS_CODECOLOR));
  end;
end;

procedure TWBEdit.FormatFont(FName, FStyle: String; FColor: TColor; FSize: byte; ClearFormat: Boolean);
begin
  if ClearFormat then
    EDExecCommand('RemoveFormat', false, emptyparam); // очищаем формат

  EDExecCommand(FStyle, false, emptyparam);
  EDExecCommand('FontName',  false, 'face="'+ FName +','+ FName +','+ FName +'"');
  EDExecCommand('FontSize',  false, FSize div 3);
  EDExecCommand('ForeColor', false, RGBtoHTMLColor(FColor, IS_CODECOLOR));
end;


{ *****************************************************************************
  ->> Задать имя шрифта
 ***************************************************************************** }
procedure TWBEdit.FormatFontName(FontName: String);
begin
  EDExecCommand('FontName', false, 'face="'+ FontName +
                                        ','+ FontName +
                                        ','+ FontName + '"');
end;


{ *****************************************************************************
  ->> Задать размер шрифта
 ***************************************************************************** }
procedure TWBEdit.FormatFontSize;
begin
  ShowEdDialog(IS_FONTSIZE);
end;


{ *****************************************************************************
  ->> Задать цвет объекту
 ***************************************************************************** }
procedure TWBEdit.FormatColor(Color: TColor);
begin
  EDExecCommand('ForeColor', false, RGBtoHTMLColor(Color, IS_CODECOLOR));
end;


procedure TWBEdit.FormatColorDialog;
begin
  if EDColorDialog.Execute then
    EDExecCommand('ForeColor', false, RGBtoHTMLColor(EDColorDialog.Color, IS_CODECOLOR));
end;


{ *****************************************************************************
  ->> Форматирование (Bold, Italic, Underline)
 ***************************************************************************** }
procedure TWBEdit.FormatBold;
begin
  EDExecCommand('Bold', false, emptyparam);
end;

procedure TWBEdit.FormatItalic;
begin
  EDExecCommand('Italic', false, emptyparam);
end;

procedure TWBEdit.FormatUnderline;
begin
  EDExecCommand('Underline', false, emptyparam);
end;


{ *****************************************************************************
  ->> Форматированный список
 ***************************************************************************** }
procedure TWBEdit.FormatSortList;
begin
//
end;


{ *****************************************************************************
  ->> Вставка тегов Sub (нижний индекс) , Sup (верхний индекс)
 ***************************************************************************** }
procedure TWBEdit.Insert_SUB;
begin
  ShowEdDialog(IS_SUB);
end;

procedure TWBEdit.Insert_SUP;
begin
  ShowEdDialog(IS_SUP);
end;


{ *****************************************************************************
  ->> Вставка картинки
 ***************************************************************************** }
procedure TWBEdit.InsertImage;
begin
  ShowEdDialog(IS_IMAGE);
end;


{ *****************************************************************************
  ->> Вставка кода
 ***************************************************************************** }
procedure TWBEdit.InsertHTMLcode(Code: WideString);
begin
  EDInsertHTMLcode(Code);
end;


{ *****************************************************************************
  ->> Вставка горизонтальной линии (тег <HR>)
 ***************************************************************************** }
procedure TWBEdit.Insert_HR;
begin
  EDInsertHTMLcode('<HR>');
end;


{ *****************************************************************************
  ->> Вставка 2-х <BR> тегов
 ***************************************************************************** }
procedure TWBEdit.Insert_2BR;
begin
  EDInsertHTMLcode('<BR><BR>');
end;


{ *****************************************************************************
  ->> Вставка спец. символов
 ***************************************************************************** }
procedure TWBEdit.Insert_copy;
begin
  EDInsertHTMLcode('&copy;');
end;

procedure TWBEdit.Insert_reg;
begin
  EDInsertHTMLcode('&reg;');
end;

procedure TWBEdit.Insert_nbsp;
begin
  EDInsertHTMLcode('&nbsp;');
end;

procedure TWBEdit.Insert_tm;
begin
  EDInsertHTMLcode('&#8482;');
end;


{ *****************************************************************************
  ->> Загрузить код в WB
 ***************************************************************************** }
procedure TWBEdit.LoadHTMLsrc(WB: TWebBrowser; Code: String);
begin
  EDLoadHTMLsource(WB, Code);
end;


{ *****************************************************************************
  ->> Получить исходный код из WB
 ***************************************************************************** }
function  TWBEdit.GetHTMLsrc(const FileName: string; WB: TWebBrowser; Filter: Boolean): String;
begin
  Result := EDGetHTMLsource(FileName, WB, Filter);
end;


{ *****************************************************************************
  ->> Режим редактирования
 ***************************************************************************** }
procedure TWBEdit.DesignMode(Mode: Boolean);
begin
  EDDesignMode(Mode);
end;


{ *****************************************************************************
  ->> Стиль компонента WB
 ***************************************************************************** }
procedure TWBEdit.Style(WB: TWebBrowser);
begin
  EDStyle(WB);
end;


initialization
  OleInitialize(nil);  // необходимо, для правильной работы команд cut, copy...

finalization
  OleUninitialize;


end.
