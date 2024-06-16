library OLAddin4HiconisAS;

uses
  ComServ,
  OLAddin4HiconisAS_TLB in 'OLAddin4HiconisAS_TLB.pas',
  OLAddin4HiconisAS_IMPL in 'OLAddin4HiconisAS_IMPL.pas' {AddInModule: TAddInModule} {CoOLAddin4HiconisAS: CoClass};

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
