unit FrmOLEmailList;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  AdvOfficePager, NxColumns, NxColumnClasses, NxCustomGridControl, Vcl.Menus,
  OtlCommon, OtlComm,
  mormot.core.base, mormot.core.variants,
  FrameOLEmailList4Ole, UnitCopyData,
  UnitOutLookDataType, UnitOLEmailRecord2, JvComponentBase, JvCaptionButton;

type
  TOLEmailListF = class(TForm)
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
    Button4: TButton;
    OLEmailListFr: TOutlookEmailListFr;
    ShowDefaultFolderName1: TMenuItem;
    JvCaptionButton1: TJvCaptionButton;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure OLEmailListFrBitBtn1Click(Sender: TObject);
    procedure ShowDefaultFolderName1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure OLEmailListFrgrid_MailCellDblClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure JvCaptionButton1Click(Sender: TObject);
    procedure OLEmailListFrButton1Click(Sender: TObject);
  private
  protected
    procedure InitVar();
    procedure DestroyVar();

    procedure MailGridDblClick(Sender: TObject; ACol, ARow: Integer);
  public
    FOwnerFormHandle: THandle;
    FHiconisASManageMode: Boolean;

    procedure Log(AMsg: string);
  end;

function CreateNShowOLEmailListForm(AOLEmailSrchRec: TOLEmailSrchRec): integer;

var
  OLEmailListF: TOLEmailListF;

implementation

{$R *.dfm}

uses UnitSynLog2, UnitNextGridUtil2, UnitDynamicFormManager;

function CreateNShowOLEmailListForm(AOLEmailSrchRec: TOLEmailSrchRec): integer;
var
  LOLEmailListF: TOLEmailListF;
begin
  if not Assigned(g_GPFormManager) then
    g_GPFormManager := TGPFormManager.Create;

  LOLEmailListF := g_GPFormManager.CreateNewForm(nil, TOLEmailListF, False) as TOLEmailListF;

  with LOLEmailListF do
  begin
    FOwnerFormHandle := AOLEmailSrchRec.FOwnerFormHandle;
    FHiconisASManageMode := AOLEmailSrchRec.FHiconisASManageMode;
    OLEmailListFr.AutoMoveCB.Checked := AOLEmailSrchRec.AutoMoveCBCheck;
    OLEmailListFr.SubFolderCB.Checked := AOLEmailSrchRec.AutoMoveCBCheck;
    OLEmailListFr.AeroButton1.Enabled := AOLEmailSrchRec.SaveToDBButtonEnable;
    OLEmailListFr.BitBtn1.Enabled := AOLEmailSrchRec.CloseButtonEnable;
    OLEmailListFr.InitVarFromOwner(AOLEmailSrchRec);
    OLEmailListFr.grid_Mail.Options := OLEmailListFr.grid_Mail.Options + [goMultiSelect];

    Show;
  end;
end;

{ TOLControlF }

procedure TOLEmailListF.Button1Click(Sender: TObject);
begin
  OLEmailListFr.SendCmd2WorkerThrd(olckGetFolderList, TOmniValue.CastFrom(''));
end;

procedure TOLEmailListF.Button3Click(Sender: TObject);
begin
  OLEmailListFr.SendCmd2WorkerThrd(olckInitVar, TOmniValue.CastFrom(''));
end;

procedure TOLEmailListF.Button4Click(Sender: TObject);
begin
  OLEmailListFr.SendCmd2WorkerThrd(olckGetSelectedMailItemFromExplorer, TOmniValue.CastFrom(''));
end;

procedure TOLEmailListF.DestroyVar;
begin
end;

procedure TOLEmailListF.FormActivate(Sender: TObject);
begin
  OLEmailListFr.SetLogProc(Log);
end;

procedure TOLEmailListF.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DestroyVar();
end;

procedure TOLEmailListF.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if Assigned(g_GPFormManager) then
  begin
    g_GPFormManager.DestroyForm(Handle);

  if not g_GPFormManager.HandleIsExistInList(FOwnerFormHandle) then
    SendMessage(FOwnerFormHandle, MSG_OLEMAILLISTF_CLOSE, 0, 0);
  end;

end;

procedure TOLEmailListF.FormCreate(Sender: TObject);
begin
  InitVar();
end;

procedure TOLEmailListF.InitVar;
begin
  InitSynLog();
  OLEmailListFr.FDefaultMoveFolder := '\\jhpark@hyundai-gs.com(2024)\HiCONIS(2024)';
end;

procedure TOLEmailListF.JvCaptionButton1Click(Sender: TObject);
begin
  if JvCaptionButton1.Down then
    JvCaptionButton1.ImageIndex := 48
  else
    JvCaptionButton1.ImageIndex := 47;

  if JvCaptionButton1.Down then
    FormStyle := fsStayOnTop
  else
    FormStyle := fsNormal;
end;

procedure TOLEmailListF.Log(AMsg: string);
begin
  if Memo1.Lines.Count > MEMO_LOG_MAX_LINE_COUNT then
    Memo1.Lines.Clear;

  Memo1.Lines.Add(AMsg);

  DoLog(AMsg, False, sllInfo);
end;

procedure TOLEmailListF.MailGridDblClick(Sender: TObject; ACol, ARow: Integer);
var
  LStr: string;
begin
  LStr := OLEmailListFr.grid_Mail.CellsByName['HullNo', ARow] + ';' + OLEmailListFr.grid_Mail.CellsByName['ClaimNo', ARow];
  SendCopyData2(FOwnerFormHandle, LStr, ARow, 0);
  Log('MailGridDblClick : ' + LStr);
end;

procedure TOLEmailListF.OLEmailListFrBitBtn1Click(Sender: TObject);
begin
  Close;
end;

procedure TOLEmailListF.OLEmailListFrButton1Click(Sender: TObject);
begin
  OLEmailListFr.Button1Click(Sender);

end;

procedure TOLEmailListF.OLEmailListFrgrid_MailCellDblClick(Sender: TObject;
  ACol, ARow: Integer);
begin
  if FHiconisASManageMode then
  begin
    MailGridDblClick(Sender, ACol, ARow);
    Log('MailGridDblClick');
  end
  else
  begin
    OLEmailListFr.grid_MailCellDblClick(Sender, ACol, ARow);
  end;
end;

procedure TOLEmailListF.ShowDefaultFolderName1Click(Sender: TObject);
begin
  ShowMessage(OLEmailListFr.FDefaultMoveFolder);
end;

end.


