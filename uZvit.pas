unit uZvit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Buttons, ComCtrls,
  DMudo, Excel97, OleServer;

type
  TZvitF = class(TForm)
    GroupBox1: TGroupBox;
    dtpBeg: TDateTimePicker;
    dtpEnd: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    GroupBox2: TGroupBox;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    ExApp: TExcelApplication;
    Exbook: TExcelWorkbook;


    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ZvitF: TZvitF;
  SQL2 : TStrings;   //Сколько пришло документов
                     // Сколько выполненых документов

implementation

uses uProcess, main;

{$R *.DFM}

procedure ClearMas;
var
   i,j,k : integer;
begin
     For i:=1 to NumUpr do
       For j:=1 to NumType do
         For k:=3 to NumInstanse do
           With Rez[i,j,k] do
             begin
               Prib := 0;
               Vip  := 0;
               NotVip := 0;
             end;
end;

procedure TZvitF.BitBtn1Click(Sender: TObject);
var
    sBeg, sEnd : String;
    n : OLeVariant;
    i,j,k,h: integer;
    s: string;

begin
    TempCursor:=Cursor;
    MainFRM.Cursor:=crSQLWait;
    Visible:=False;
    Processing.Show;
    Processing.Animate1.Active:=True;
    Processing.PBar.Position:=5;

    sBeg:=DateToStr(dtpBeg.Date);
    sEnd:=DateToStr(dtpEnd.Date);
    ClearMas;

    SQL2:=TStringList.Create;
    try
      // Определение поступивших документов
      SQL2.Text:='select Fullfil.Instanse_ID, Docum.Type_ID, Fullfil.Upravl_ID '+
                 'from Fullfil, Docum '+
                 'where ('+#39+sBeg+#39+' <= Docum.Date_Post) '+
                 'and (Docum.Date_Post <= '+#39+sEnd+#39+' ) '+
                 'and (Fullfil.Docum_ID = Docum.ID)';

      Data.IBQuery1.SQL.Assign(SQL2);
      Data.IBQuery1.Open;
      Processing.PBar.Position:=10;
      // Цикл обработки
      Data.IBQuery1.First;
      With  Data.IBQuery1 do
        While not Eof do
        begin
          inc(Rez[Fields[2].AsInteger,Fields[1].AsInteger,Fields[0].AsInteger].Prib);
          Next;
        end;

    finally
      Data.IBQuery1.Close;
      SQL2.Free;
    end;

    Processing.PBar.Position:=20;
    SQL2:=TStringList.Create;
    try
      // Определение выполненых/невыполненых документов
      SQL2.Text:='select F.Instanse_ID, D.Type_ID, F.Upravl_ID, F.Date_Vipoln, F.DateOtv, F.OtmOVipoln '+
      'from Fullfil F, Docum D where ('+ #39+sBeg+#39+' <= F.Date_Vipoln) and (F.Date_Vipoln <= '+#39+sEnd+#39+' )';
      Data.IBQuery1.SQL.Assign(SQL2);
      Data.IBQuery1.Open;
      Processing.PBar.Position:=25;

      // Цикл обработки
      Data.IBQuery1.First;
      With  Data.IBQuery1 do
        While not Eof do
        begin
           if (Fields[4].AsDateTime > Fields[3].AsDateTime) or
              (Fields[4].AsDateTime = StrToDate('01.01.00'))
              // Не выполнен документ
              then
                inc(Rez[Fields[2].AsInteger,Fields[1].AsInteger,Fields[0].AsInteger].NotVip)
              // Выполнен документ
              else
                inc(Rez[Fields[2].AsInteger,Fields[1].AsInteger,Fields[0].AsInteger].Vip);

           Next;
        end;

    finally
      Data.IBQuery1.Close;
      SQL2.Free;
    end;
    Processing.PBar.Position:=30;

    // Формирование отчета в формате таблицы Excel
  try
    GetDir(0,s);
    if FileExists('Report.xls') then
         DeleteFile('Report.xls');
    n:=s+'\Tabl_1.xls';
    ExApp.Workbooks.Add(n,0);
    Processing.PBar.Position:=35;
    Exbook.ConnectTo(ExApp.ActiveWorkbook );
    Processing.PBar.Position:=45;

    for i:=1 to NumUpr do begin
      h:=8;
      for j:=1 to NumType do begin
        for k:=3 to NumInstanse do begin
           ExApp.Cells.Item[h,(i-1)*3+2].Value:=Rez[i,j,k].Prib;
           ExApp.Cells.Item[h,(i-1)*3+3].Value:=Rez[i,j,k].Vip;
           ExApp.Cells.Item[h,(i-1)*3+4].Value:=Rez[i,j,k].NotVip;
           h:=h+1;
        end;
        h:=h+8;
      end;
    end;
    Processing.PBar.Position:=75;
    MainFRM.Cursor:=TempCursor;
    Exbook.SaveAs(s+'\report.xls',xlWorkbookNormal,'','',false,false,
                  xlNoChange,1,true,1,1,0);
    Processing.PBar.Position:=90;
    Processing.Animate1.Active:=false;
    
  finally
    Processing.PBar.Position:=100;
    Exbook.Close;
    ExApp.Quit;
    Processing.Close;
    Close;
  end;
end;

procedure TZvitF.BitBtn2Click(Sender: TObject);
begin
     Close;
end;

end.



