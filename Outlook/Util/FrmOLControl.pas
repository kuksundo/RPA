unit FrmOLControl;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  AdvOfficePager,
  OtlCommon, OtlComm,
  mormot.core.base, mormot.core.variants,
  FrameOLEmailList4Ole,
  UnitOutLookDataType, UnitOLEmailRecord2;

type
  TOLControlF = class(TForm)
    Panel1: TPanel;
    AdvOfficePager1: TAdvOfficePager;
    LogPage: TAdvOfficePage;
    Splitter1: TSplitter;
    Edit1: TEdit;
    Button1: TButton;
    Button2: TButton;
    Memo1: TMemo;
    Button3: TButton;
    EmailPage: TAdvOfficePage;
    OLEmailListFr: TOutlookEmailListFr;
    Button4: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure OLEmailListFrBitBtn1Click(Sender: TObject);
  private
  protected
    procedure InitVar();
    procedure DestroyVar();
  public
    procedure Log(AMsg: string);
    procedure ShowEmailListFromSrchRec(ASearchRec: TOLEmailSrchRec);
  end;

var
  OLControlF: TOLControlF;

implementation

{$R *.dfm}

uses UnitSynLog2, UnitNextGridUtil2;

{ TOLControlF }

procedure TOLControlF.Button1Click(Sender: TObject);
begin
  OLEmailListFr.SendCmd2WorkerThrd(olckGetFolderList, TOmniValue.CastFrom(''));
end;

procedure TOLControlF.Button3Click(Sender: TObject);
begin
  OLEmailListFr.SendCmd2WorkerThrd(olckInitVar, TOmniValue.CastFrom(''));
end;

procedure TOLControlF.Button4Click(Sender: TObject);
begin
  OLEmailListFr.SendCmd2WorkerThrd(olckGetSelectedMailItemFromExplorer, TOmniValue.CastFrom(''));
end;

procedure TOLControlF.DestroyVar;
begin
end;

procedure TOLControlF.FormActivate(Sender: TObject);
begin
  OLEmailListFr.SetLogProc(Log);
end;

procedure TOLControlF.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DestroyVar();
end;

procedure TOLControlF.FormCreate(Sender: TObject);
begin
  InitVar();
end;

procedure TOLControlF.InitVar;
begin
  InitSynLog();
end;

procedure TOLControlF.Log(AMsg: string);
begin
  if Memo1.Lines.Count > MEMO_LOG_MAX_LINE_COUNT then
    Memo1.Lines.Clear;

  Memo1.Lines.Add(AMsg);

  DoLog(AMsg, False, sllInfo);
end;

procedure TOLControlF.OLEmailListFrBitBtn1Click(Sender: TObject);
begin
  Close;
end;

procedure TOLControlF.ShowEmailListFromSrchRec(ASearchRec: TOLEmailSrchRec);
var
  LUtf8: RawUtf8;
  LVar: variant;
begin
  LUtf8 := GetEmailList2JSONArrayFromSearchRec(ASearchRec);
  LVar := _JSON(LUtf8);
  GetListFromVariant2NextGrid(OLEmailListFr.grid_Mail, LVar, True, True, True, True);
//  AddNextGridRowsFromVariant2(OLEmailListFr.grid_Mail, LVar);
end;

end.


