unit OLAddin4HiconisAS_IMPL;

interface

uses
  SysUtils, ComObj, ComServ, ActiveX, Variants, Outlook2000, Office2000, adxAddIn, OLAddin4HiconisAS_TLB,
  Vcl.ExtCtrls, AdvAlertWindow, adxHostAppEvents, System.Classes, System.Generics.Collections,
  Winapi.Windows, System.SyncObjs, adxolFormsManager, Messages, Vcl.Dialogs,
  IdMessage, UnitSynLog2,
  mormot.core.base, mormot.rest.server, mormot.orm.core, mormot.rest.http.server,
  mormot.soa.server, mormot.core.datetime, mormot.rest.memserver, mormot.soa.core,
  mormot.core.interfaces, mormot.core.buffers, mormot.core.unicode, mormot.core.os,
  mormot.core.data, mormot.core.variants, mormot.core.json,
  StompClient, StompTypes,
  OtlCommon, OtlComm, OtlTaskControl, OtlContainerObserver, otlTask,
  Cromis.Comm.Custom, Cromis.Comm.IPC, Cromis.Threading, Cromis.AnyValue,
  OLMailWSCallbackInterface2, UnitCommonWSInterface2, UnitClientInfoClass2, CommonData2,
  UnitOLDataType
  ;

type
  TCoOLAddin4HiconisAS = class(TadxAddin, ICoOLAddin4HiconisAS)
  end;

  TAddInModule = class(TadxCOMAddInModule)
    adxContextMenu1: TadxContextMenu;
    adxContextMenu2: TadxContextMenu;
    adxOutlookAppEvents1: TadxOutlookAppEvents;
    AdvAlertWindow1: TAdvAlertWindow;
    Timer1: TTimer;
    procedure adxCOMAddInModuleCreate(Sender: TObject);
  private
  protected
  public
  end;

implementation

{$R *.dfm}

procedure TAddInModule.adxCOMAddInModuleCreate(Sender: TObject);
begin
  ShowMessage('adxCOMAddInModuleCreate');
end;

initialization
  TadxFactory.Create(ComServer, TCoOLAddin4HiconisAS, CLASS_CoOLAddin4HiconisAS, TAddInModule);

end.
