unit OLMailWSCallbackInterface2;

interface

uses SysUtils,
  mormot.core.base, mormot.soa.core, mormot.core.interfaces;
  //UnitCommonWSInterface2;

type
  TCommMode = (cmRESTful, cmWebSocket, cmCromisIPC);
  TCommModes = Set of TCommMode;

  IOLMailCallback = interface(IInvokable)
    ['{29AF8173-73F9-4CFF-94D6-7A0E9CC882E4}']
    procedure ClientExecute(const command, msg: string);
  end;

  IOLMailService = interface(IServiceWithCallbackReleased)
    ['{7CF29E82-7FF4-4578-BA74-8AEF4D2E7E1B}']
    procedure Join(const pseudo: string; const callback: IOLMailCallback);
    procedure CallbackReleased(const callback: IInvokable; const interfaceName: RawUTF8);
    function ServerExecute(const Acommand: string): RawUTF8;
    function GetOLEmailInfo(ACommand: string): RawUTF8;
    function GetOLEmailAccountInfo: RawUTF8;
  end;

const
  OL_ROOT_NAME_4_WS = 'root';
  OL_PORT_NAME_4_WS = '704';
  OL_APPLICATION_NAME_4_WS = 'OL_RestService_WebSocket';
  OL_DEFAULT_IP = '10.22.43.55';
  MQ_SERVER_IP = '10.100.23.63';
  MQ_USER_ID = 'pjh';
  MQ_PASSWORD = 'pjh';
  OL4WS_TRANSMISSION_KEY = 'OL_PrivateKey';


implementation

{$IFDEF USE_MORMOT_WS}
initialization
//  TInterfaceFactory.RegisterInterfaces([
//    TypeInfo(IOLMailService),TypeInfo(IOLMailCallback)]);
{$ENDIF}

end.
