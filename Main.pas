unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBTables, Grids, DBGrids, StdCtrls, Buttons, Excel2000,
  OleServer, ComCtrls;

type
  TForm1 = class(TForm)
    DataSource1: TDataSource;
    BitBtn1: TBitBtn;
    ExApp: TExcelApplication;
    Exbook: TExcelWorkbook;
    date1: TDateTimePicker;
    date2: TDateTimePicker;
    Query1: TQuery;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Database1: TDatabase;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn5: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.BitBtn1Click(Sender: TObject);
var
     s,n, st, st1: string;
     nach, kon : string;
     i: integer;
begin
  try
    Database1.Connected:=false;
//    Database1.AliasName:='LocalServer';
    Database1.LoginPrompt:=false;
    Database1.Params.Values['User Name'] := 'qwer';
    Database1.Params.Values['Password']  := '1234';
    Database1.Connected:=true;


    GetDir(0,s);
    if FileExists('Forma17.xls') then
         DeleteFile('Forma17.xls');
    n:=s+'\Shablon\Shablon.xls';
    ExApp.Workbooks.Add(n,0);
    Exbook.ConnectTo(ExApp.ActiveWorkbook );

    ExApp.Cells.Item[4,2].Value:= 'про результати роботи медичного відділу Чернігівського цеху за період з '+DateToStr(Date1.Date)+' по '+DateToStr(Date2.Date);

    // Изменение даты
    st:= DateToStr(Date1.Date);
    nach:=st;
    nach[1]:=st[4];
    nach[2]:=st[5];
    nach[4]:=st[1];
    nach[5]:=st[2];

    st:= DateToStr(Date2.Date);
    kon:=st;
    kon[1]:=st[4];
    kon[2]:=st[5];
    kon[4]:=st[1];
    kon[5]:=st[2];

    Query1.ExecSQL;
    Query1.Close;

    // Заполнение полей ВСЬОГО
    Query1.SQL.Clear;
    Query1.SQL.Add('SELECT     a.vid_f17, COUNT(b.T11F06K) AS Expr1 ');
    Query1.SQL.Add('FROM       vid2 a INNER JOIN ');
    Query1.SQL.Add('      RS_T11 b ON a.S10F00 = b.S10F00 INNER JOIN ');
    Query1.SQL.Add('      RS_T01 c ON c.N_KAR = b.N_KAR INNER JOIN ');
    Query1.SQL.Add('      S03 d ON d.S03F00 = c.S03F00 INNER JOIN ' );
    Query1.SQL.Add('      T11_temp e ON e.N_KAR = b.N_KAR ');
    Query1.SQL.Add('WHERE (b.N_ZAV = 3) AND (b.T11F01 >='+chr(39)+Nach+chr(39)+' ) AND (b.T11F01 <='+chr(39)+Kon+chr(39)+' ) AND (c.TCREATED >='+chr(39)+Nach+chr(39)+' ) AND (c.TCREATED <='+chr(39)+Kon+chr(39)+' ) ');
    Query1.SQL.Add('GROUP BY a.vid_f17');
    Query1.Open;

    While not Query1.Eof do
    begin
        ExApp.Cells.Item[9+Query1.Fields[0].AsInteger,3].Value:=
           Query1.Fields[1].AsInteger;
        Query1.Next;

    end;
    Query1.Close;

    Query1.SQL.Clear;
    // Заполнение полей ВСЬОГО  2
    Query1.SQL.Clear;
    Query1.SQL.Add('SELECT     a.vid_f17, COUNT(b.T11F06K) AS Expr1 ');
    Query1.SQL.Add('FROM       vid2 a INNER JOIN ');
    Query1.SQL.Add('      RS_T11 b ON a.S10F00 = b.S10F00 INNER JOIN ');
    Query1.SQL.Add('      RS_T01 c ON c.N_KAR = b.N_KAR INNER JOIN ');
    Query1.SQL.Add('      S03 d ON d.S03F00 = c.S03F00 INNER JOIN ' );
    Query1.SQL.Add('      T11_temp e ON e.N_KAR = b.N_KAR ');
    Query1.SQL.Add('WHERE (b.N_ZAV = 3) AND (b.T11F01 >='+chr(39)+Nach+chr(39)+' ) AND (b.T11F01 <='+chr(39)+Kon+chr(39)+' ) AND (c.TCREATED <'+chr(39)+Nach+chr(39)+' ) ');
    Query1.SQL.Add('GROUP BY a.vid_f17');

    Query1.Open;

    While not Query1.Eof do
    begin
        ExApp.Cells.Item[9+Query1.Fields[0].AsInteger,4].Value:=
           Query1.Fields[1].AsInteger;
        Query1.Next;

    end;
    Query1.Close;

    // Заполнение полей Д18
    Query1.SQL.Clear;
    Query1.SQL.Add('SELECT     a.vid_f17, COUNT(b.T11F06K) AS Expr1 ');
    Query1.SQL.Add('FROM       vid2 a INNER JOIN ');
    Query1.SQL.Add('      RS_T11 b ON a.S10F00 = b.S10F00 INNER JOIN ');
    Query1.SQL.Add('      RS_T01 c ON c.N_KAR = b.N_KAR INNER JOIN ');
    Query1.SQL.Add('      S03 d ON d.S03F00 = c.S03F00 INNER JOIN ' );
    Query1.SQL.Add('      T11_temp e ON e.N_KAR = b.N_KAR ');
    Query1.SQL.Add('WHERE (b.N_ZAV = 3) AND (b.T11F01 >='+chr(39)+Nach+chr(39)+' ) AND (b.T11F01 <='+chr(39)+Kon+chr(39)+' ) AND (d.S03F00 = 25) AND (c.TCREATED >='+chr(39)+Nach+chr(39)+' ) AND (c.TCREATED <='+chr(39)+Kon+chr(39)+' ) ');
    Query1.SQL.Add('GROUP BY a.vid_f17');
    Query1.Open;

    While not Query1.Eof do
    begin
        ExApp.Cells.Item[9+Query1.Fields[0].AsInteger,5].Value:=
           Query1.Fields[1].AsInteger;
        Query1.Next;

    end;
    Query1.Close;

    // Заполнение полей Д18 2
    Query1.SQL.Clear;
    Query1.SQL.Add('SELECT     a.vid_f17, COUNT(b.T11F06K) AS Expr1 ');
    Query1.SQL.Add('FROM       vid2 a INNER JOIN ');
    Query1.SQL.Add('      RS_T11 b ON a.S10F00 = b.S10F00 INNER JOIN ');
    Query1.SQL.Add('      RS_T01 c ON c.N_KAR = b.N_KAR INNER JOIN ');
    Query1.SQL.Add('      S03 d ON d.S03F00 = c.S03F00 INNER JOIN ' );
    Query1.SQL.Add('      T11_temp e ON e.N_KAR = b.N_KAR ');
    Query1.SQL.Add('WHERE (b.N_ZAV = 3) AND (b.T11F01 >='+chr(39)+Nach+chr(39)+' ) AND (b.T11F01 <='+chr(39)+Kon+chr(39)+' ) AND (d.S03F00 = 25) AND (c.TCREATED <'+chr(39)+Nach+chr(39)+' ) ');
    Query1.SQL.Add('GROUP BY a.vid_f17');
    Query1.Open;

    While not Query1.Eof do
    begin
        ExApp.Cells.Item[9+Query1.Fields[0].AsInteger,6].Value:=
           Query1.Fields[1].AsInteger;
        Query1.Next;

    end;
    Query1.Close;

    // Заполнение полей Інвалід війни
    Query1.SQL.Clear;
    Query1.SQL.Add('SELECT     a.vid_f17, COUNT(b.T11F06K) AS Expr1 ');
    Query1.SQL.Add('FROM       vid2 a INNER JOIN ');
    Query1.SQL.Add('      RS_T11 b ON a.S10F00 = b.S10F00 INNER JOIN ');
    Query1.SQL.Add('      RS_T01 c ON c.N_KAR = b.N_KAR INNER JOIN ');
    Query1.SQL.Add('      S03 d ON d.S03F00 = c.S03F00 INNER JOIN ' );
    Query1.SQL.Add('      T11_temp e ON e.N_KAR = b.N_KAR ');
    Query1.SQL.Add('WHERE (b.N_ZAV = 3) AND (b.T11F01 >='+chr(39)+Nach+chr(39)+' ) AND (b.T11F01 <='+chr(39)+Kon+chr(39)+' ) AND (d.S03F00 = 1) AND (c.TCREATED >='+chr(39)+Nach+chr(39)+' ) AND (c.TCREATED <='+chr(39)+Kon+chr(39)+' ) ');
    Query1.SQL.Add('GROUP BY a.vid_f17');
    Query1.Open;

    While not Query1.Eof do
    begin
        ExApp.Cells.Item[9+Query1.Fields[0].AsInteger,7].Value:=
           Query1.Fields[1].AsInteger;
        Query1.Next;

    end;
    Query1.Close;

    // Заполнение полей Інвалід війни 2
    Query1.SQL.Clear;
    Query1.SQL.Add('SELECT     a.vid_f17, COUNT(b.T11F06K) AS Expr1 ');
    Query1.SQL.Add('FROM       vid2 a INNER JOIN ');
    Query1.SQL.Add('      RS_T11 b ON a.S10F00 = b.S10F00 INNER JOIN ');
    Query1.SQL.Add('      RS_T01 c ON c.N_KAR = b.N_KAR INNER JOIN ');
    Query1.SQL.Add('      S03 d ON d.S03F00 = c.S03F00 INNER JOIN ' );
    Query1.SQL.Add('      T11_temp e ON e.N_KAR = b.N_KAR ');
    Query1.SQL.Add('WHERE (b.N_ZAV = 3) AND (b.T11F01 >='+chr(39)+Nach+chr(39)+' ) AND (b.T11F01 <='+chr(39)+Kon+chr(39)+' ) AND (d.S03F00 = 1) AND (c.TCREATED <'+chr(39)+Nach+chr(39)+' ) ');
    Query1.SQL.Add('GROUP BY a.vid_f17');
    Query1.Open;

    While not Query1.Eof do
    begin
        ExApp.Cells.Item[9+Query1.Fields[0].AsInteger,8].Value:=
           Query1.Fields[1].AsInteger;
        Query1.Next;

    end;
    Query1.Close;

    ////////////////// Заполнение второй части //////////////////
    // Заполнение полей ВСЬОГО
    Query1.SQL.Clear;
    Query1.SQL.Add('SELECT     a.vid_f17, SUM(b.T11F06K) AS Expr1 ');
    Query1.SQL.Add('FROM       vid2 a INNER JOIN ');
    Query1.SQL.Add('      RS_T11 b ON a.S10F00 = b.S10F00 INNER JOIN ');
    Query1.SQL.Add('      RS_T01 c ON c.N_KAR = b.N_KAR INNER JOIN ');
    Query1.SQL.Add('      S03 d ON d.S03F00 = c.S03F00 ');
    Query1.SQL.Add('WHERE (b.N_ZAV = 3) AND (b.T11F01 >='+chr(39)+Nach+chr(39)+' ) AND (b.T11F01 <='+chr(39)+Kon+chr(39)+' ) ');
    Query1.SQL.Add('GROUP BY a.vid_f17');
    Query1.Open;

    While not Query1.Eof do
    begin
        ExApp.Cells.Item[9+Query1.Fields[0].AsInteger,11].Value:=
           Query1.Fields[1].AsInteger;
        Query1.Next;
    end;
    Query1.Close;

    // Заполнение полей Д18
    Query1.SQL.Clear;
    Query1.SQL.Add('SELECT     a.vid_f17, SUM(b.T11F06K) AS Expr1 ');
    Query1.SQL.Add('FROM       vid2 a INNER JOIN ');
    Query1.SQL.Add('      RS_T11 b ON a.S10F00 = b.S10F00 INNER JOIN ');
    Query1.SQL.Add('      RS_T01 c ON c.N_KAR = b.N_KAR INNER JOIN ');
    Query1.SQL.Add('      S03 d ON d.S03F00 = c.S03F00 ');
    Query1.SQL.Add('WHERE (b.N_ZAV = 3) AND (b.T11F01 >='+chr(39)+Nach+chr(39)+' ) AND (b.T11F01 <='+chr(39)+Kon+chr(39)+' ) AND (d.S03F00 = 25) ');
    Query1.SQL.Add('GROUP BY a.vid_f17');
    Query1.Open;

    While not Query1.Eof do
    begin
        ExApp.Cells.Item[9+Query1.Fields[0].AsInteger,12].Value:=
           Query1.Fields[1].AsInteger;
        Query1.Next;

    end;
    Query1.Close;

    // Заполнение полей Інвалід війни
    Query1.SQL.Clear;
    Query1.SQL.Add('SELECT     a.vid_f17, SUM(b.T11F06K) AS Expr1 ');
    Query1.SQL.Add('FROM       vid2 a INNER JOIN ');
    Query1.SQL.Add('      RS_T11 b ON a.S10F00 = b.S10F00 INNER JOIN ');
    Query1.SQL.Add('      RS_T01 c ON c.N_KAR = b.N_KAR      INNER JOIN ');
    Query1.SQL.Add('      S03 d ON d.S03F00 = c.S03F00 ');
    Query1.SQL.Add('WHERE (b.N_ZAV = 3) AND (b.T11F01 >='+chr(39)+Nach+chr(39)+' ) AND (b.T11F01 <='+chr(39)+Kon+chr(39)+' ) AND (d.S03F00 = 1) ');
    Query1.SQL.Add('GROUP BY a.vid_f17');
    Query1.Open;

    While not Query1.Eof do
    begin
        ExApp.Cells.Item[9+Query1.Fields[0].AsInteger,13].Value:=
           Query1.Fields[1].AsInteger;
        Query1.Next;

    end;
    Query1.Close;

    Exbook.SaveAs(s+'\Forma17.xls',xlWorkbookNormal,'','',false,false,
                  xlNoChange,1,true,1,1,0);
    ShowMessage('Звіт виконано! Файл знаходиться: '+s+'\Forma17.xls');
  finally
    Exbook.Close;
    ExApp.Quit;

  end;


end;

//   ЗВІТ ПО ЗАМОВЛЕННЯМ (ВЗУТТЯ - ПО ПІВПАРАМ) ЗА ПЕРІОД ...
procedure TForm1.BitBtn2Click(Sender: TObject);
var
     s,n, st, st1: string;
     nach, kon : string;
     i: integer;
begin
  try
    Database1.Connected:=false;
//    Database1.AliasName:='LocalServer';
    Database1.LoginPrompt:=false;
    Database1.Params.Values['User Name'] := 'qwer';
    Database1.Params.Values['Password']  := '1234';
    Database1.Connected:=true;


    GetDir(0,s);
    if FileExists('ZvitPoPivparam.xls') then
         DeleteFile('ZvitPoPivparam.xls');
    n:=s+'\Shablon\Shablon2.xls';
    ExApp.Workbooks.Add(n,0);
    Exbook.ConnectTo(ExApp.ActiveWorkbook );

   // ExApp.Cells.Item[4,2].Value:= 'про результати роботи медичного відділу Чернігівського цеху за період з '+DateToStr(Date1.Date)+' по '+DateToStr(Date2.Date);

    // Изменение даты
    st:= DateToStr(Date1.Date);
    nach:=st;
    nach[1]:=st[4];
    nach[2]:=st[5];
    nach[4]:=st[1];
    nach[5]:=st[2];

    st:= DateToStr(Date2.Date);
    kon:=st;
    kon[1]:=st[4];
    kon[2]:=st[5];
    kon[4]:=st[1];
    kon[5]:=st[2];


    // Заполнение полей ВСЬОГО
    Query1.SQL.Clear;
    Query1.SQL.Add('    SELECT     a.T01F01 + '+chr(39)+chr(32)+chr(39)+' + a.T01F02 + '+chr(39)+chr(32)+chr(39)+' + a.T01F03 AS Expr1, a.T01F05, b.N_ZAK, b.R_SHIFR1, CONVERT(DATETIME, CONVERT(VARCHAR(15), b.T11F04, 101)) AS Expr3, b.T11F01, b.T11F07, ISNULL(g.S42F05, '+chr(39)+ chr(39)+' ) + '+chr(39)+chr(32)+chr(39)+' + a.T01F07 AS Expr2, b.T11F20, b.T11F06, b.T11F071, b.T11F072, b.R_KX1, b.R_KX2  ');
    Query1.SQL.Add('FROM         RS_T11 b LEFT OUTER JOIN ');
    Query1.SQL.Add('      RS_T01 a ON a.N_KAR = b.N_KAR LEFT OUTER JOIN ');
    Query1.SQL.Add('      S42 g ON a.S42F00 = g.S42F00 ');
    Query1.SQL.Add('WHERE     (b.T11F04 >= '+chr(39)+Nach+chr(39)+') AND (b.T11F04 < '+chr(39)+Kon+chr(39)+') ');
    Query1.SQL.Add('ORDER BY b.T11F20');
    Query1.Open;

    i:=2;
    While not Query1.Eof do
    begin
        if (Query1.Fields[9].AsString = 'півпар.') and (Query1.Fields[10].AsString <> Query1.Fields[11].AsString)
           then
           // Вставка двух строк
             begin
                ExApp.Cells.Item[i,1].Value:= Query1.Fields[0].AsString;
                ExApp.Cells.Item[i,2].Value:= Query1.Fields[1].AsString;
                ExApp.Cells.Item[i,3].Value:= Query1.Fields[2].AsString+'(1)';
                ExApp.Cells.Item[i,4].Value:= Query1.Fields[3].AsString+'.'+Query1.Fields[12].AsString;
                ExApp.Cells.Item[i,5].Value:= Query1.Fields[4].AsString;
                ExApp.Cells.Item[i,6].Value:= Query1.Fields[5].AsString;
                ExApp.Cells.Item[i,8].Value:= Query1.Fields[7].AsString;
                ExApp.Cells.Item[i,9].Value:= Query1.Fields[8].AsString;
                if Query1.Fields[11].AsString = '0'
                   then ExApp.Cells.Item[i,7].Value:= Query1.Fields[10].AsFloat/2
                   else ExApp.Cells.Item[i,7].Value:= Query1.Fields[10].AsString;
                
                i:=i+1;

                ExApp.Cells.Item[i,1].Value:= Query1.Fields[0].AsString;
                ExApp.Cells.Item[i,2].Value:= Query1.Fields[1].AsString;
                ExApp.Cells.Item[i,3].Value:= Query1.Fields[2].AsString+'(2)';
                ExApp.Cells.Item[i,5].Value:= Query1.Fields[4].AsString;
                ExApp.Cells.Item[i,6].Value:= Query1.Fields[5].AsString;
                ExApp.Cells.Item[i,8].Value:= Query1.Fields[7].AsString;
                ExApp.Cells.Item[i,9].Value:= Query1.Fields[8].AsString;
                if Query1.Fields[11].AsString = '0'
                   then
                   begin
                   ExApp.Cells.Item[i,7].Value:= Query1.Fields[10].AsFloat/2;
                   ExApp.Cells.Item[i,4].Value:= Query1.Fields[3].AsString+'.'+Query1.Fields[12].AsString;
                   end
                   else
                   begin
                   ExApp.Cells.Item[i,7].Value:= Query1.Fields[11].AsString;
                   ExApp.Cells.Item[i,4].Value:= Query1.Fields[3].AsString+'.'+Query1.Fields[13].AsString;
                   end;

             end
           else
             begin
                ExApp.Cells.Item[i,1].Value:= Query1.Fields[0].AsString;
                ExApp.Cells.Item[i,2].Value:= Query1.Fields[1].AsString;
                ExApp.Cells.Item[i,3].Value:= Query1.Fields[2].AsString;
                ExApp.Cells.Item[i,4].Value:= Query1.Fields[3].AsString;
                ExApp.Cells.Item[i,5].Value:= Query1.Fields[4].AsString;
                ExApp.Cells.Item[i,6].Value:= Query1.Fields[5].AsString;
                ExApp.Cells.Item[i,7].Value:= Query1.Fields[6].AsString;
                ExApp.Cells.Item[i,8].Value:= Query1.Fields[7].AsString;
                ExApp.Cells.Item[i,9].Value:= Query1.Fields[8].AsString;
             end;
        Query1.Next;

       i:=i+1;
    end;
    Query1.Close;

    Exbook.SaveAs(s+'\ZvitPoPivparam.xls',xlWorkbookNormal,'','',false,false,
                  xlNoChange,1,true,1,1,0);
    ShowMessage('Звіт виконано! Файл знаходиться: '+s+'\ZvitPoPivparam.xls');
  finally
    Exbook.Close;
    ExApp.Quit;

  end;

end;

procedure TForm1.BitBtn3Click(Sender: TObject);
var
     s,n, st, st1: string;
     nach, kon : string;
     i: integer;
begin
  try
    Database1.Connected:=false;
//    Database1.AliasName:='LocalServer';
    Database1.LoginPrompt:=false;
    Database1.Params.Values['User Name'] := 'qwer';
    Database1.Params.Values['Password']  := '1234';
    Database1.Connected:=true;


    GetDir(0,s);
    if FileExists('ZvitPoPivparamDataVidachi.xls') then
         DeleteFile('ZvitPoPivparamDataVidachi.xls');
    n:=s+'\Shablon\Shablon3.xls';
    ExApp.Workbooks.Add(n,0);
    Exbook.ConnectTo(ExApp.ActiveWorkbook );

    // Изменение даты
    st:= DateToStr(Date1.Date);
    nach:=st;
    nach[1]:=st[4];
    nach[2]:=st[5];
    nach[4]:=st[1];
    nach[5]:=st[2];

    st:= DateToStr(Date2.Date);
    kon:=st;
    kon[1]:=st[4];
    kon[2]:=st[5];
    kon[4]:=st[1];
    kon[5]:=st[2];


    //
    Query1.SQL.Clear;
    Query1.SQL.Add('    SELECT     a.T01F01 + '+chr(39)+chr(32)+chr(39)+' + a.T01F02 + '+chr(39)+chr(32)+chr(39)+' + a.T01F03 AS Expr1, b.N_ART, b.N_ZAK, b.R_SHIFR1, CONVERT(DATETIME, CONVERT(VARCHAR(15), b.T11F04, 101)) AS Expr3, b.T11F01, b.T11F07, ISNULL(g.S42F05, '+chr(39)+ chr(39)+' ) + '+chr(39)+chr(32)+chr(39)+' + a.T01F07 AS Expr2, b.T11F20, b.T11F06, b.T11F071, b.T11F072, b.R_KX1, b.R_KX2  ');
    Query1.SQL.Add('FROM         RS_T11 b LEFT OUTER JOIN ');
    Query1.SQL.Add('      RS_T01 a ON a.N_KAR = b.N_KAR LEFT OUTER JOIN ');
    Query1.SQL.Add('      S42 g ON a.S42F00 = g.S42F00 ');
    Query1.SQL.Add('WHERE     (b.T11F01 >= '+chr(39)+Nach+chr(39)+') AND (b.T11F01 < '+chr(39)+Kon+chr(39)+') ');
    Query1.SQL.Add('ORDER BY b.T11F20');
    Query1.Open;

    i:=2;
    While not Query1.Eof do
    begin
        if (Query1.Fields[9].AsString = 'півпар.') and (Query1.Fields[10].AsString <> Query1.Fields[11].AsString)
           then
           // Вставка двух строк
             begin
                ExApp.Cells.Item[i,1].Value:= Query1.Fields[0].AsString;
                ExApp.Cells.Item[i,2].Value:= Query1.Fields[1].AsString;
                ExApp.Cells.Item[i,3].Value:= Query1.Fields[2].AsString+'(1)';
                ExApp.Cells.Item[i,4].Value:= Query1.Fields[3].AsString+'.'+Query1.Fields[12].AsString;
                ExApp.Cells.Item[i,5].Value:= Query1.Fields[4].AsString;
                ExApp.Cells.Item[i,6].Value:= Query1.Fields[5].AsString;
                ExApp.Cells.Item[i,8].Value:= Query1.Fields[7].AsString;
                ExApp.Cells.Item[i,9].Value:= Query1.Fields[8].AsString;
                if Query1.Fields[11].AsString = '0'
                   then ExApp.Cells.Item[i,7].Value:= Query1.Fields[10].AsFloat/2
                   else ExApp.Cells.Item[i,7].Value:= Query1.Fields[10].AsString;
                
                i:=i+1;

                ExApp.Cells.Item[i,1].Value:= Query1.Fields[0].AsString;
                ExApp.Cells.Item[i,2].Value:= Query1.Fields[1].AsString;
                ExApp.Cells.Item[i,3].Value:= Query1.Fields[2].AsString+'(2)';
                ExApp.Cells.Item[i,5].Value:= Query1.Fields[4].AsString;
                ExApp.Cells.Item[i,6].Value:= Query1.Fields[5].AsString;
                ExApp.Cells.Item[i,8].Value:= Query1.Fields[7].AsString;
                ExApp.Cells.Item[i,9].Value:= Query1.Fields[8].AsString;
                if Query1.Fields[11].AsString = '0'
                   then
                   begin
                   ExApp.Cells.Item[i,7].Value:= Query1.Fields[10].AsFloat/2;
                   ExApp.Cells.Item[i,4].Value:= Query1.Fields[3].AsString+'.'+Query1.Fields[12].AsString;
                   end
                   else
                   begin
                   ExApp.Cells.Item[i,7].Value:= Query1.Fields[11].AsString;
                   ExApp.Cells.Item[i,4].Value:= Query1.Fields[3].AsString+'.'+Query1.Fields[13].AsString;
                   end;

             end
           else
             begin
                ExApp.Cells.Item[i,1].Value:= Query1.Fields[0].AsString;
                ExApp.Cells.Item[i,2].Value:= Query1.Fields[1].AsString;
                ExApp.Cells.Item[i,3].Value:= Query1.Fields[2].AsString;
                ExApp.Cells.Item[i,4].Value:= Query1.Fields[3].AsString;
                ExApp.Cells.Item[i,5].Value:= Query1.Fields[4].AsString;
                ExApp.Cells.Item[i,6].Value:= Query1.Fields[5].AsString;
                ExApp.Cells.Item[i,7].Value:= Query1.Fields[6].AsString;
                ExApp.Cells.Item[i,8].Value:= Query1.Fields[7].AsString;
                ExApp.Cells.Item[i,9].Value:= Query1.Fields[8].AsString;
             end;
        Query1.Next;

       i:=i+1;
    end;
    Query1.Close;

    Exbook.SaveAs(s+'\ZvitPoPivparamDataVidachi.xls',xlWorkbookNormal,'','',false,false,
                  xlNoChange,1,true,1,1,0);
    ShowMessage('Звіт виконано! Файл знаходиться: '+s+'\ZvitPoPivparamDataVidachi.xls');
  finally
    Exbook.Close;
    ExApp.Quit;

  end;

end;

//   ЗВІТ  VIKORISTANI MATERIALY (ПО ЗАМОВЛЕННЯМ ) ЗА ПЕРІОД ...
procedure TForm1.BitBtn4Click(Sender: TObject);
var
     s,n, st, st1: string;
     nach, kon : string;
     i, npp : integer;
     tmpShifr : string;
begin
  try
    Database1.Connected:=false;
//    Database1.AliasName:='LocalServer';
    Database1.LoginPrompt:=false;
    Database1.Params.Values['User Name'] := 'qwer';
    Database1.Params.Values['Password']  := '1234';
    Database1.Connected:=true;


    GetDir(0,s);
    if FileExists('SpisokMaterialov.xls') then
         DeleteFile('SpisokMaterialov.xls');
    n:=s+'\Shablon\SpisokMaterialovShablon.xls';
    ExApp.Workbooks.Add(n,0);
    Exbook.ConnectTo(ExApp.ActiveWorkbook );

   // ExApp.Cells.Item[4,2].Value:= 'про результати роботи медичного відділу Чернігівського цеху за період з '+DateToStr(Date1.Date)+' по '+DateToStr(Date2.Date);

    // Изменение даты
    st:= DateToStr(Date1.Date);
    nach:=st;
    nach[1]:=st[4];
    nach[2]:=st[5];
    nach[4]:=st[1];
    nach[5]:=st[2];

    st:= DateToStr(Date2.Date);
    kon:=st;
    kon[1]:=st[4];
    kon[2]:=st[5];
    kon[4]:=st[1];
    kon[5]:=st[2];


    // Заполнение полей ВСЬОГО
    Query1.SQL.Clear;
    //  ОБУВЬ
    Query1.SQL.Add('SELECT    t12.T12F02, t12.T12F03 AS CENA, SUM(t12.T12F04) AS KOLVO, ');               // 0,1,2
    Query1.SQL.Add('          SUM(t12.T12F03 * t12.T12F04) AS SUMMA, n.N_MAT01, n.eizm ');  // 3,4,5
    Query1.SQL.Add('FROM      RS_T12 t12 LEFT OUTER JOIN ');
    Query1.SQL.Add('          RS_T11 t11 ON t11.N_ZAK = t12.N_ZAK LEFT OUTER JOIN ');
    Query1.SQL.Add('          N_MAT n ON t12.T12F02 = n.N_MAT00 ');
    Query1.SQL.Add('WHERE     (T11.N_ZAK LIKE '+chr(39)+'В3%'+chr(39)+') AND (t12.N_RAB00 = - 1) AND (t11.T11F05 >= '+chr(39)+Nach+chr(39)+') AND (t11.T11F05 <= '+chr(39)+Kon+chr(39)+') AND (n.N_ZAV = 3) ');
    Query1.SQL.Add('GROUP BY  t12.T12F02, t12.T12F03, n.N_MAT01, n.eizm');

    Query1.SQL.Add('UNION ALL');
    
    //  ПРОТЕЗИ
    Query1.SQL.Add('SELECT    t12.T12F02, t12.T12F03 AS CENA, SUM(t12.T12F04 * t11.T11F06K) AS KOLVO, '); // 0,1,2
    Query1.SQL.Add('          SUM(t12.T12F03 * t12.T12F04 * t11.T11F06K) AS SUMMA, n.N_MAT01, n.eizm ');  // 3,4,5
    Query1.SQL.Add('FROM      RS_T12 t12 LEFT OUTER JOIN ');
    Query1.SQL.Add('          RS_T11 t11 ON t11.N_ZAK = t12.N_ZAK LEFT OUTER JOIN ');
    Query1.SQL.Add('          N_MAT n ON t12.T12F02 = n.N_MAT00 ');
    Query1.SQL.Add('WHERE     (T11.N_ZAK NOT LIKE '+chr(39)+'В3%'+chr(39)+') AND (t12.N_RAB00 = - 1) AND (t11.T11F05 >= '+chr(39)+Nach+chr(39)+') AND (t11.T11F05 <= '+chr(39)+Kon+chr(39)+') AND (n.N_ZAV = 3) ');
    Query1.SQL.Add('GROUP BY  t12.T12F02, t12.T12F03, n.N_MAT01, n.eizm ');
    Query1.SQL.Add('ORDER BY  t12.T12F02');
    Query1.Open;

    i:=4;
    npp:=1;
    tmpShifr := '';
    ExApp.Cells.Item[2,1].Value:= 'за період з '+DateToStr(Date1.Date)+' по '+DateToStr(Date2.Date);
    While not Query1.Eof do
    begin
        ExApp.Cells.Item[i,1].Value:= npp;
        ExApp.Cells.Item[i,2].Value:= Query1.Fields[0].AsString;
        ExApp.Cells.Item[i,3].Value:= Query1.Fields[4].AsString;
        ExApp.Cells.Item[i,4].Value:= Query1.Fields[5].AsString;
        ExApp.Cells.Item[i,5].Value:= Query1.Fields[1].AsString;
        ExApp.Cells.Item[i,6].Value:= Query1.Fields[2].AsString;

       Query1.Next;
       i:=i+1;
       npp:=npp+1;
    end;
    Query1.Close;

     // Заполнение полей ПО ШИФРАМ
    Query1.SQL.Clear;
    //  ОБУВЬ
    Query1.SQL.Add('SELECT    t12.T12F02, t12.T12F03 AS CENA, SUM(t12.T12F04) AS KOLVO, ');               // 0,1,2
    Query1.SQL.Add('          SUM(t12.T12F03 * t12.T12F04) AS SUMMA, t11.R_SHIFR1, n.N_MAT01, n.eizm ');  // 3,4,5,6
    Query1.SQL.Add('FROM      RS_T12 t12 LEFT OUTER JOIN ');
    Query1.SQL.Add('          RS_T11 t11 ON t11.N_ZAK = t12.N_ZAK LEFT OUTER JOIN ');
    Query1.SQL.Add('          N_MAT n ON t12.T12F02 = n.N_MAT00 ');
    Query1.SQL.Add('WHERE     (T11.N_ZAK LIKE '+chr(39)+'В3%'+chr(39)+') AND (t12.N_RAB00 = - 1) AND (t11.T11F05 >= '+chr(39)+Nach+chr(39)+') AND (t11.T11F05 <= '+chr(39)+Kon+chr(39)+') AND (n.N_ZAV = 3) ');
    Query1.SQL.Add('GROUP BY t12.T12F02, t12.T12F03, t11.R_SHIFR1, n.N_MAT01, n.eizm');

    Query1.SQL.Add('UNION ALL');
    
    //  ПРОТЕЗИ
    Query1.SQL.Add('SELECT    t12.T12F02, t12.T12F03 AS CENA, SUM(t12.T12F04 * t11.T11F06K) AS KOLVO, ');               // 0,1,2
    Query1.SQL.Add('          SUM(t12.T12F03 * t12.T12F04 * t11.T11F06K) AS SUMMA, t11.R_SHIFR1, n.N_MAT01, n.eizm ');  // 3,4,5,6
    Query1.SQL.Add('FROM      RS_T12 t12 LEFT OUTER JOIN ');
    Query1.SQL.Add('          RS_T11 t11 ON t11.N_ZAK = t12.N_ZAK LEFT OUTER JOIN ');
    Query1.SQL.Add('          N_MAT n ON t12.T12F02 = n.N_MAT00 ');
    Query1.SQL.Add('WHERE     (T11.N_ZAK NOT LIKE '+chr(39)+'В3%'+chr(39)+') AND (t12.N_RAB00 = - 1) AND (t11.T11F05 >= '+chr(39)+Nach+chr(39)+') AND (t11.T11F05 <= '+chr(39)+Kon+chr(39)+') AND (n.N_ZAV = 3) ');
    Query1.SQL.Add('GROUP BY t12.T12F02, t12.T12F03, t11.R_SHIFR1, n.N_MAT01, n.eizm');
    Query1.SQL.Add('ORDER BY t11.R_SHIFR1, t12.T12F02');
    Query1.Open;

    i:=i+2;
    tmpShifr := '';
    While not Query1.Eof do
    begin

        if tmpShifr <> Query1.Fields[4].AsString
           // smena shifra - vyvesti zagolovok
           then
             begin
                i:=i+2;
                ExApp.Cells.Item[i,3].Value:= 'По шифру '+Query1.Fields[4].AsString;
                i:=i+1;
                npp:=1;
             end;
        ExApp.Cells.Item[i,1].Value:= npp;
        ExApp.Cells.Item[i,2].Value:= Query1.Fields[0].AsString;
        ExApp.Cells.Item[i,3].Value:= Query1.Fields[5].AsString;
        ExApp.Cells.Item[i,4].Value:= Query1.Fields[6].AsString;
        ExApp.Cells.Item[i,5].Value:= Query1.Fields[1].AsString;
        ExApp.Cells.Item[i,6].Value:= Query1.Fields[2].AsString;

       tmpShifr := Query1.Fields[4].AsString;
       Query1.Next;
       i:=i+1;
       npp:=npp+1;
    end;
    Query1.Close;


    Exbook.SaveAs(s+'\SpisokMaterialov.xls',xlWorkbookNormal,'','',false,false,
                  xlNoChange,1,true,1,1,0);
    ShowMessage('Звіт виконано! Файл знаходиться: '+s+'\SpisokMaterialov.xls');
  finally
    Exbook.Close;
    ExApp.Quit;

  end;

end;

procedure TForm1.BitBtn5Click(Sender: TObject);
var
     s,n, st, st1: string;
     nach, kon : string;
     i, npp : integer;
     tmpShifr : string;
begin
  try
    Database1.Connected:=false;
//    Database1.AliasName:='LocalServer';
    Database1.LoginPrompt:=false;
    Database1.Params.Values['User Name'] := 'qwer';
    Database1.Params.Values['Password']  := '1234';
    Database1.Connected:=true;


    GetDir(0,s);
    if FileExists('SpisokMaterialovArtikul.xls') then
         DeleteFile('SpisokMaterialovArtikul.xls');
    n:=s+'\Shablon\SpisokMaterialovArtikulShablon.xls';
    ExApp.Workbooks.Add(n,0);
    Exbook.ConnectTo(ExApp.ActiveWorkbook );

    // Изменение даты
    st:= DateToStr(Date1.Date);
    nach:=st;
    nach[1]:=st[4];
    nach[2]:=st[5];
    nach[4]:=st[1];
    nach[5]:=st[2];

    st:= DateToStr(Date2.Date);
    kon:=st;
    kon[1]:=st[4];
    kon[2]:=st[5];
    kon[4]:=st[1];
    kon[5]:=st[2];


    // Заполнение полей ВСЬОГО
    Query1.SQL.Clear;
    //  ОБУВЬ
    Query1.SQL.Add('SELECT    t12.T12F02, t12.T12F03 AS CENA, SUM(t12.T12F04) AS KOLVO, ');               // 0,1,2
    Query1.SQL.Add('          SUM(t12.T12F03 * t12.T12F04) AS SUMMA, n.N_MAT01, n.eizm ');  // 3,4,5
    Query1.SQL.Add('FROM      RS_T12 t12 LEFT OUTER JOIN ');
    Query1.SQL.Add('          RS_T11 t11 ON t11.N_ZAK = t12.N_ZAK LEFT OUTER JOIN ');
    Query1.SQL.Add('          N_MAT n ON t12.T12F02 = n.N_MAT00 ');
    Query1.SQL.Add('WHERE     (T11.N_ZAK LIKE '+chr(39)+'В3%'+chr(39)+') AND (t12.N_RAB00 = - 1) AND (t11.T11F05 >= '+chr(39)+Nach+chr(39)+') AND (t11.T11F05 <= '+chr(39)+Kon+chr(39)+') AND (n.N_ZAV = 3) ');
    Query1.SQL.Add('GROUP BY  t12.T12F02, t12.T12F03, n.N_MAT01, n.eizm');

    Query1.SQL.Add('UNION ALL');
    
    //  ПРОТЕЗИ
    Query1.SQL.Add('SELECT    t12.T12F02, t12.T12F03 AS CENA, SUM(t12.T12F04 * t11.T11F06K) AS KOLVO, '); // 0,1,2
    Query1.SQL.Add('          SUM(t12.T12F03 * t12.T12F04 * t11.T11F06K) AS SUMMA, n.N_MAT01, n.eizm ');  // 3,4,5
    Query1.SQL.Add('FROM      RS_T12 t12 LEFT OUTER JOIN ');
    Query1.SQL.Add('          RS_T11 t11 ON t11.N_ZAK = t12.N_ZAK LEFT OUTER JOIN ');
    Query1.SQL.Add('          N_MAT n ON t12.T12F02 = n.N_MAT00 ');
    Query1.SQL.Add('WHERE     (T11.N_ZAK NOT LIKE '+chr(39)+'В3%'+chr(39)+') AND (t12.N_RAB00 = - 1) AND (t11.T11F05 >= '+chr(39)+Nach+chr(39)+') AND (t11.T11F05 <= '+chr(39)+Kon+chr(39)+') AND (n.N_ZAV = 3) ');
    Query1.SQL.Add('GROUP BY  t12.T12F02, t12.T12F03, n.N_MAT01, n.eizm ');
    Query1.SQL.Add('ORDER BY  t12.T12F02');
    Query1.Open;

    i:=4;
    npp:=1;
    tmpShifr := '';
    ExApp.Cells.Item[2,1].Value:= 'за період з '+DateToStr(Date1.Date)+' по '+DateToStr(Date2.Date);
    While not Query1.Eof do
    begin
        ExApp.Cells.Item[i,1].Value:= npp;
        ExApp.Cells.Item[i,2].Value:= Query1.Fields[0].AsString;
        ExApp.Cells.Item[i,3].Value:= Query1.Fields[4].AsString;
        ExApp.Cells.Item[i,4].Value:= Query1.Fields[5].AsString;
        ExApp.Cells.Item[i,5].Value:= Query1.Fields[1].AsString;
        ExApp.Cells.Item[i,6].Value:= Query1.Fields[2].AsString;

       Query1.Next;
       i:=i+1;
       npp:=npp+1;
    end;
    Query1.Close;

     // Заполнение полей ПО АРТИКУЛАМ
    Query1.SQL.Clear;
    //  ОБУВЬ
    Query1.SQL.Add('SELECT    t12.T12F02, t12.T12F03 AS CENA, SUM(t12.T12F04) AS KOLVO, ');               // 0,1,2
    Query1.SQL.Add('          SUM(t12.T12F03 * t12.T12F04) AS SUMMA, t11.N_ART, n.N_MAT01, n.eizm ');  // 3,4,5,6
    Query1.SQL.Add('FROM      RS_T12 t12 LEFT OUTER JOIN ');
    Query1.SQL.Add('          RS_T11 t11 ON t11.N_ZAK = t12.N_ZAK LEFT OUTER JOIN ');
    Query1.SQL.Add('          N_MAT n ON t12.T12F02 = n.N_MAT00 ');
    Query1.SQL.Add('WHERE     (T11.N_ZAK LIKE '+chr(39)+'В3%'+chr(39)+') AND (t12.N_RAB00 = - 1) AND (t11.T11F05 >= '+chr(39)+Nach+chr(39)+') AND (t11.T11F05 <= '+chr(39)+Kon+chr(39)+') AND (n.N_ZAV = 3) ');
    Query1.SQL.Add('GROUP BY t12.T12F02, t12.T12F03, t11.N_ART, n.N_MAT01, n.eizm');

    Query1.SQL.Add('UNION ALL');
    
    //  ПРОТЕЗИ
    Query1.SQL.Add('SELECT    t12.T12F02, t12.T12F03 AS CENA, SUM(t12.T12F04 * t11.T11F06K) AS KOLVO, ');               // 0,1,2
    Query1.SQL.Add('          SUM(t12.T12F03 * t12.T12F04 * t11.T11F06K) AS SUMMA, t11.N_ART, n.N_MAT01, n.eizm ');  // 3,4,5,6
    Query1.SQL.Add('FROM      RS_T12 t12 LEFT OUTER JOIN ');
    Query1.SQL.Add('          RS_T11 t11 ON t11.N_ZAK = t12.N_ZAK LEFT OUTER JOIN ');
    Query1.SQL.Add('          N_MAT n ON t12.T12F02 = n.N_MAT00 ');
    Query1.SQL.Add('WHERE     (T11.N_ZAK NOT LIKE '+chr(39)+'В3%'+chr(39)+') AND (t12.N_RAB00 = - 1) AND (t11.T11F05 >= '+chr(39)+Nach+chr(39)+') AND (t11.T11F05 <= '+chr(39)+Kon+chr(39)+') AND (n.N_ZAV = 3) ');
    Query1.SQL.Add('GROUP BY t12.T12F02, t12.T12F03, t11.N_ART, n.N_MAT01, n.eizm');
    Query1.SQL.Add('ORDER BY t11.N_ART, t12.T12F02');
    Query1.Open;

    i:=i+2;
    tmpShifr := '';
    While not Query1.Eof do
    begin

        if tmpShifr <> Query1.Fields[4].AsString
           // smena artikula - vyvesti zagolovok
           then
             begin
                i:=i+2;
                ExApp.Cells.Item[i,3].Value:= 'По артикулу '+Query1.Fields[4].AsString;
                i:=i+1;
                npp:=1;
             end;
        ExApp.Cells.Item[i,1].Value:= npp;
        ExApp.Cells.Item[i,2].Value:= Query1.Fields[0].AsString;
        ExApp.Cells.Item[i,3].Value:= Query1.Fields[5].AsString;
        ExApp.Cells.Item[i,4].Value:= Query1.Fields[6].AsString;
        ExApp.Cells.Item[i,5].Value:= Query1.Fields[1].AsString;
        ExApp.Cells.Item[i,6].Value:= Query1.Fields[2].AsString;

       tmpShifr := Query1.Fields[4].AsString;
       Query1.Next;
       i:=i+1;
       npp:=npp+1;
    end;
    Query1.Close;


    Exbook.SaveAs(s+'\SpisokMaterialovArtikul.xls',xlWorkbookNormal,'','',false,false,
                  xlNoChange,1,true,1,1,0);
    ShowMessage('Звіт виконано! Файл знаходиться: '+s+'\SpisokMaterialovArtikul.xls');
  finally
    Exbook.Close;
    ExApp.Quit;

  end;

end;

procedure TForm1.FormCreate(Sender: TObject);
begin
     date1.Date:= now();
     date2.Date:= now();
end;

End.
