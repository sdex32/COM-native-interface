unit Unit1;

interface

uses
  ActiveX,
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

const
   MY_FACTORY : TGUID = '{99BB4CB8-EFF9-4017-B4A8-B8B3B6C8BA63}';
   MY_OBJECT : TGUID = '{FDB0CC23-EA12-4B8B-9455-DE5C9011B596}';

type
  IMyOBJ = interface(IUnknown)
//    ['{FDB0CC23-EA12-4B8B-9455-DE5C9011B596}'] // we did not need this decoration
    function Ping(a:longword):longword; stdcall;
  end;


procedure TForm1.Button1Click(Sender: TObject);
var res :HResult;
    i:longword;
    s:shortstring;
    fac : IMyOBJ;
begin
 //  res := CoInitializeEx(nil, 2 {COINIT_APARTMENTTHREADED});
   res := CoInitialize(nil);
   if res = 1 then
   begin
      fac := nil;
      res := CoCreateInstance (MY_FACTORY,
                               nil,
                               CLSCTX_INPROC_SERVER,
                               MY_OBJECT,
                               fac);


      if res = 0 then
      begin
         res := fac.ping(1);   // must return 11
         str(res,s);
         Label1.Caption := string(s);
      end;
      CoUninitialize;
   end;
end;




end.
