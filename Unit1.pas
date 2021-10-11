unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, Menus, ImgList, OleCtrls, SHDocVw,

  WBEdit, About;

type
  TForm1 = class(TForm)
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ImageList1: TImageList;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    ToolButton14: TToolButton;
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    ToolButton17: TToolButton;
    ToolButton18: TToolButton;
    ToolButton19: TToolButton;
    ToolButton20: TToolButton;
    ToolButton21: TToolButton;
    ToolButton22: TToolButton;
    ToolButton23: TToolButton;
    ToolButton24: TToolButton;
    ToolButton25: TToolButton;
    ToolButton26: TToolButton;
    ToolButton27: TToolButton;
    ToolButton28: TToolButton;
    WebBrowser1: TWebBrowser;
    N6: TMenuItem;
    N7: TMenuItem;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    procedure N7Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure WebBrowser1DocumentComplete(Sender: TObject;
      const pDisp: IDispatch; var URL: OleVariant);
    procedure N3Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure ToolButton12Click(Sender: TObject);
    procedure ToolButton13Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure ToolButton6Click(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure ToolButton8Click(Sender: TObject);
    procedure ToolButton10Click(Sender: TObject);
    procedure ToolButton15Click(Sender: TObject);
    procedure ToolButton16Click(Sender: TObject);
    procedure ToolButton18Click(Sender: TObject);
    procedure ToolButton19Click(Sender: TObject);
    procedure ToolButton20Click(Sender: TObject);
    procedure ToolButton25Click(Sender: TObject);
    procedure ToolButton26Click(Sender: TObject);
    procedure ToolButton27Click(Sender: TObject);
    procedure ToolButton22Click(Sender: TObject);
    procedure ToolButton23Click(Sender: TObject);
    procedure ToolButton28Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    WBEdit: TWBEdit;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}


procedure TForm1.N7Click(Sender: TObject);
begin
  AboutForm.ShowModal;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  WBEdit := TWBEdit.Create;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  WBEdit.Free;
end;

// Открыть HTML файл
procedure TForm1.N2Click(Sender: TObject);
begin
  try
    WBEdit.DesignMode(false);
  except
  end;
  if OpenDialog1.Execute then
    WebBrowser1.Navigate(OpenDialog1.FileName);
end;

// Сохранить в HTML
procedure TForm1.N3Click(Sender: TObject);
begin
  if SaveDialog1.Execute then
    WBEdit.GetHTMLsrc(SaveDialog1.FileName, WebBrowser1, False);
end;

// Выход
procedure TForm1.N5Click(Sender: TObject);
begin
  Application.Terminate;
end;

// переводим в режим редактора
procedure TForm1.WebBrowser1DocumentComplete(Sender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
  ToolBar1.Enabled := true;
  WBEdit.Disp := pDisp;
  WBEdit.DesignMode(true);
end;

{ -----------------------------------------------------------------------------
   -> Инструменты
 ----------------------------------------------------------------------------- }

// Выделит всё
procedure TForm1.ToolButton12Click(Sender: TObject);
begin
  WBEdit.SelectAll;
end;

// Очистить формат
procedure TForm1.ToolButton13Click(Sender: TObject);
begin
  WBEdit.FormatRemove;
end;

procedure TForm1.ToolButton5Click(Sender: TObject);
begin
  WBEdit.Edit_Cut;
end;

procedure TForm1.ToolButton6Click(Sender: TObject);
begin
  WBEdit.Edit_Copy;
end;

procedure TForm1.ToolButton7Click(Sender: TObject);
begin
  WBEdit.Edit_Paste;
end;

procedure TForm1.ToolButton8Click(Sender: TObject);
begin
  WBEdit.Edit_Delete;
end;

// найти
procedure TForm1.ToolButton10Click(Sender: TObject);
begin
  WBEdit.FindText;
end;

// диалог выбора цвета
procedure TForm1.ToolButton15Click(Sender: TObject);
begin
  WBEdit.FormatColorDialog;
end;

// шрифта
procedure TForm1.ToolButton16Click(Sender: TObject);
begin
  WBEdit.FormatFontDialog;
end;

// начертание шрифта
procedure TForm1.ToolButton18Click(Sender: TObject);
begin
  WBEdit.FormatBold;
end;

procedure TForm1.ToolButton19Click(Sender: TObject);
begin
  WBEdit.FormatItalic;
end;

procedure TForm1.ToolButton20Click(Sender: TObject);
begin
  WBEdit.FormatUnderline;
end;

// равнение по краям и центру
procedure TForm1.ToolButton25Click(Sender: TObject);
begin
  WBEdit.FormatAlignLeft;
end;

procedure TForm1.ToolButton26Click(Sender: TObject);
begin
  WBEdit.FormatAlignCenter;
end;

procedure TForm1.ToolButton27Click(Sender: TObject);
begin
  WBEdit.FormatAlignRight;
end;

// индексы, верхний и нижний
procedure TForm1.ToolButton22Click(Sender: TObject);
begin
  WBEdit.Insert_Sub;
end;

procedure TForm1.ToolButton23Click(Sender: TObject);
begin
  WBEdit.Insert_Sup
end;

// вставка линии
procedure TForm1.ToolButton28Click(Sender: TObject);
begin
  WBEdit.Insert_HR;
end;

// пробел
procedure TForm1.N10Click(Sender: TObject);
begin
  WBEdit.Insert_nbsp;
end;

// копирайт
procedure TForm1.N11Click(Sender: TObject);
begin
  WBEdit.Insert_copy;
end;

// ®
procedure TForm1.N12Click(Sender: TObject);
begin
  WBEdit.Insert_reg;
end;

procedure TForm1.N13Click(Sender: TObject);
begin
  WBEdit.Insert_tm;
end;

end.
