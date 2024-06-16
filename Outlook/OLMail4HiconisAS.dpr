library OLMail4HiconisAS;

uses
  ComServ,
  OLMail4InqManage_TLB in '..\..\..\..\..\project\util\OutLookAddIn\OLMail4InqManage\OLMail4InqManage_TLB.pas',
  OLMail4HiconisAS_IMPL in 'OLMail4HiconisAS_IMPL.pas' {AddInModule: TAddInModule} {CoOLMail4InqManage: CoClass},
  UnitClientInfoClass2 in 'common\UnitClientInfoClass2.pas',
  OLMailWSCallbackInterface2 in 'OLMailWSCallbackInterface2.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
