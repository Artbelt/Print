unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, RzGrids, StdCtrls, ComCtrls, RzListVw, AdvObj, BaseGrid,
  AdvGrid, ExtCtrls, GDIPAPI, GDIPOBJ, IniFiles, Menus, Printers, RzStatus,
  RzForms, RzPanel, DateUtils, Buttons;

type
  TForm1 = class(TForm)
    pnl1: TPanel;
    grid1: TAdvStringGrid;
    img1: TImage;
    btn1: TButton;
    btn3: TButton;
    btn2: TButton;
    mm1: TMainMenu;
    mniN1: TMenuItem;
    pnl2: TPanel;
    bar: TProgressBar;
    lbl1: TLabel;
    findPanel: TPanel;
    findEdit: TEdit;
    lbl2: TLabel;
    Panel1: TPanel;
    btn4: TSpeedButton;
    btn5: TSpeedButton;
    procedure ScanFile(Path: string);
    procedure btn3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure Paint(DC: HDC);
    procedure btn2Click(Sender: TObject);
    procedure ReadIniParameters();
    procedure mniN1Click(Sender: TObject);
    procedure grid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormResize(Sender: TObject);
    procedure findEditChange(Sender: TObject);
    procedure btn4Click(Sender: TObject);
    procedure btn5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  IMageAddress:string;
  EXT:string;
  Path:string;
  Printer:string;
  Datastring: string;
  CoordX, CoordY :Real;
  FontName:string;
  FontSize:Integer;

implementation

{$R *.dfm}

var Date_Stamp : Boolean;

function DecodeString(Stroka:string):string;
var
  x:Integer;
begin
  for x := 1 to Length(Stroka) do
  begin
    if Stroka[x]='Y' then
    begin
      Stroka := Copy(Stroka,0,x-1)+
                IntToStr(YearOf(Now))[3]+
                IntToStr(YearOf(Now))[4]+
                Copy(Stroka,x+1,Length(Stroka)-x);
    end;
    if Stroka[x]='W' then
    begin
       Stroka := Copy(Stroka,0,x-1)+
                IntToStr(WeekOf(Now))+
                Copy(Stroka,x+1,Length(Stroka)-x);
    end;
           
  end;
  Result :=Stroka;
end;


function SlashInsert(Path,FName: string): string;
 begin
   if Path[Length(Path)]<>'\'
    then  Result:= Path+ '\'+ Fname
    else  Result:= Path+ FName;
 end;

procedure TForm1.ScanFile(Path: string);
var
 rec: TSearchRec;
 FileName: string;
 str:string;
 attr:Integer;
begin
    // \*.* - искать файлы с любым расширением
 if FindFirst(Path+'\*.*',faAnyFile,rec)= 0 then
   begin
     repeat
        // если имя . или .. то игнорировать
       if (rec.Name='.') or (rec.Name='..') then Continue;
          Filename:= SlashInsert(Path,rec.Name);
        // если директория то открываем папку (рекурсия)
       if rec.Attr = faDirectory then
        begin
          ScanFile(FileName);
          Continue;
        end;

              {Copy(rec.Name,1,Length(Edit1.Text)) - выделяем из найденной строки столько же
                  символов сколько в заданной строке от точки}
              if AnsiUpperCase(EXT)=AnsiUpperCase(Copy(rec.Name,
              Pos('.',rec.Name),Length(EXT)))
              then
              begin

                if (rec.Attr and SysUtils.faHidden > 0)  then
                begin
                   //ShowMessage('!');
                end else
                begin
                  //ShowMessage('ENTER NO HIDDEN');
                  grid1.AddRow;
                  str:= pchar(Rec.Name);
                  Delete(str,Length(str)-3,4);
                  grid1.Cells[0,grid1.RowCount-1]:=str;
                  grid1.Cells[1,grid1.RowCount-1]:=(pchar( Path + '\' + Rec.Name));
                end;  

              end;

     until FindNext(rec)<>0;
   end;
    FindClose(rec);
end;

procedure TForm1.ReadIniParameters;
var
  ini:TINIFile;
  ini_path:string;
begin

    ini_path:=ExtractFilePath(Application.Exename)+'settings.ini';
    ini:=TIniFile.Create(ini_path);
    EXT:= ini.ReadString('FILES','EXT','*.*');
    Path:= ini.ReadString('Path','Path',''{ExtractFilePath(Application.Exename)});
    Printer:= ini.ReadString('PRINTER','Printer','');

    Datastring := Decodestring(ini.ReadString('HEADER','STRING','NULL'));
    CoordX :=ini.ReadFloat('HEADER','CoordX',1);
    CoordY :=ini.ReadFloat('HEADER','CoordY',1);
    FontName := ini.ReadString('HEADER','FONTNAME','Arial');
    FontSize := ini.ReadInteger('HEADER','FONTSIZE',5);
    ini.Free;
end;


Procedure TFOrm1.Paint(DC: HDC);
var
  graphics : TGPGraphics;
  Image: TGPImage;
  FontFamily: TGPFontFamily;
  Font: TGPFont;
  Context: TGPGraphics;
  Matrix: TGPMatrix;
  Brush: TGPSolidBrush;

begin
  IMageAddress:=grid1.Cells[1, grid1.Selection.Top];
  graphics := TGPGraphics.Create(DC);
  Image:= TGPImage.Create(IMageAddress);
  graphics.DrawImage(Image,0,0);
  FontFamily:=TGPFontFamily.Create(FontName);
  Font:=TGPFont.Create(FontFamily,FontSize);
  Brush:=TGPSolidBrush.Create(MakeColor(0,0,0));
  Graphics.DrawString(Datastring,-1,Font,MakePoint(CoordX,CoordY),Brush);
  Image.Free;
  graphics.Free;
  Brush.Free;
  Font.Free;
  FontFamily.Free;
end;

procedure TForm1.btn3Click(Sender: TObject);
begin
  ScanFile(path);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  grid1.ColWidths[0]:=200;
  grid1.ColWidths[1]:=00;
  grid1.ColWidths[2]:=70;
  grid1.Cells[0,0]:='Этикетка';
  grid1.Cells[2,0]:='Кол-во';
  ReadIniParameters;
  //ShowMessage(Path);
  btn3.Click;
  //grid1.FixedRows:=1;
  with Form1 do
  begin
    Height := Screen.Height-25;
    Top := Screen.Height - Height -30;
    Left := Screen.Width - Width ;
  end;
end;

procedure TForm1.btn1Click(Sender: TObject);

begin
   Paint(img1.Canvas.Handle);
end;

procedure TForm1.btn2Click(Sender: TObject);
var
  hdcPrint: HDC;
  docInfo: TDOCINFO;
  graphics: TGPGraphics;
  Image: TGPImage;
  x,y:Integer;
  copies: Integer;
  picscount:Integer;
    FontFamily: TGPFontFamily;
    Font: TGPFont;
    Brush: TGPSolidBrush;
begin
  pnl2.Visible:= True;
      bar.Position := 0;
      picscount := 0;
      for x := 1 to grid1.RowCount-1 do
      begin
        if grid1.Cells[2,x] <> '' then
        begin
           picscount := picscount + StrToInt(grid1.Cells[2,x]);
        end;
      end;
      bar.Max := picscount;

      for x := 0 to grid1.RowCount-1 do
      begin
         if (grid1.Cells[2,x] <> '0') and (grid1.Cells[2,x] <> '') and (grid1.Cells[2,x] <> 'Кол-во') then
         begin
           IMageAddress:=grid1.Cells[1,x{grid1.Selection.Top}];
           // Get a device context for the printer.
           hdcPrint := CreateDC(nil, PChar(Printer), nil, nil);
           ZeroMemory(@docInfo, sizeof(DOCINFO));
           docInfo.cbSize := sizeof(DOCINFO);
           docInfo.lpszDocName := 'GdiplusPrint';
           StartDoc(hdcPrint, docInfo);

             {отправляем заданное количество раз картинку в принтер}
             for y:= 0 to StrToInt(grid1.Cells[2,x])-1 do
             begin
                 Application.ProcessMessages;
                 StartPage(hdcPrint);
                    graphics := TGPGraphics.Create(hdcPrint);
                    Image:= TGPImage.Create(IMageAddress);
                    graphics.DrawImage(Image,0,0);
                    if Date_Stamp = True then
                    begin
                      {сюда вставляем надпись}
                        FontFamily:=TGPFontFamily.Create(FontName);
                        Font:=TGPFont.Create(FontFamily,FontSize);
                        Brush:=TGPSolidBrush.Create(MakeColor(0,0,0));
                        Graphics.DrawString(Datastring,-1,Font,MakePoint(CoordX,CoordY),Brush);
                        Brush.Free;
                        Font.Free;
                        FontFamily.Free;
                      {конец блока добаления надписи}
                    end;
                    Image.Free;
                    graphics.Free;
                 EndPage(hdcPrint);
                bar.Position := bar.Position+1;
                lbl1.Caption := 'Идет отправка файлов в принтер:'+
                                  Inttostr(bar.Position)+' из '+
                                  inttostr(picscount);
              end;

         end;
         EndDoc(hdcPrint);
         DeleteDC(hdcPrint);
      end;
  pnl2.Visible:= False;
end;

procedure TForm1.mniN1Click(Sender: TObject);
var
  x:Integer;
begin
  btn2.Click;
  findEdit.Text := '';
  for x:= 1 to grid1.RowCount-1 do
  begin
    grid1.Cells[2,x]:='';
  end;

end;

procedure TForm1.grid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

  case Key of
    48: ;
    49: ;
    50: ;
    51: ;
    52: ;
    53: ;
    54: ;
    55: ;
    56: ;
    57: ;
    96: ;
    97: ;
    98: ;
    99: ;
    100: ;
    101: ;
    102: ;
    103: ;
    104: ;
    105: ;
    115: begin

            findEdit.SetFocus;

         end;
    112: begin
           ShowMessage('Количество этикеток: ' + IntToStr(grid1.RowCount));
         end;  
    38: ;
    40: ;
    13: ;
  else
    begin
      grid1.Cells[grid1.Selection.Left,grid1.Selection.Top] :='';
      //ShowMessage(IntToStr(Key));
    end;
  end;



 end;

procedure TForm1.FormResize(Sender: TObject);
begin
  pnl2.Width := Form1.Width;
  bar.Width := pnl2.Width;
  pnl2.Top:= form1.Height-pnl2.Height-60 ;
  pnl2.Left:= form1.Width div 2-pnl2.Width div 2;
end;

procedure TForm1.findEditChange(Sender: TObject);
begin
   grid1.FilterActive := False;
   grid1.Filter.Clear;
   with grid1.Filter.Add do
   begin
     Condition :='*' + findEdit.Text + '*';
     Column := 0;
     Operation := foNONE;
   end;
   grid1.FilterActive := True;
end;

procedure TForm1.btn4Click(Sender: TObject);
var
  x:Integer;
begin
  Date_Stamp := True;
  btn2.Click;
  findEdit.Text := '';
  for x:= 1 to grid1.RowCount-1 do
  begin
    grid1.Cells[2,x]:='';
  end;
  Date_Stamp := True;
end;

procedure TForm1.btn5Click(Sender: TObject);
var
  x:Integer;
begin
  Date_Stamp := False;
  btn2.Click;
  findEdit.Text := '';
  for x:= 1 to grid1.RowCount-1 do
  begin
    grid1.Cells[2,x]:='';
  end;
  Date_Stamp := False;

end;

end.
