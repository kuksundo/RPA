program HiMTech_Dev;

uses
  Vcl.Forms,
  FrmHiMTech in 'FrmHiMTech.pas' {HiMTechF},
  UnitHiMTechData in 'UnitHiMTechData.pas',
  FrmHiMTechDM in 'FrmHiMTechDM.pas' {DataModule1: TDataModule},
  UnitHiMTechExcelUtil in 'UnitHiMTechExcelUtil.pas',
  UnitRegAppUtil in '..\..\..\NoGitHub\RegCodeManager2\Common\UnitRegAppUtil.pas',
  EasterEgg in '..\..\..\..\..\..\project\common\EasterEgg.pas',
  FormAboutDefs in '..\..\..\..\..\..\project\common\Forms\TFormAbout\FormAboutDefs.pas',
  UnitHiMTechCLO in 'UnitHiMTechCLO.pas';

{$R *.res}

//Created by pjh on 2025-05-24
//{1CBD7C54-5AA3-4C63-A652-1592F9FB389C}=Prod Code -> Version Info -> InternalName에 Encrypted로 저장됨
//UnitCryptUtil2.EncryptString_Syn3()를 이용하여 암호화 함
//Encrypted: lfS525wdRPCfT2ubBLZ2t+n7eq6SmwDpQdnJlrleQwB0ZJoa1jkUM7icsYtxly3T

begin
  if UnitRegAppUtil.TgpAppSigInfo.CheckRegByAppSigUsingRegistry('') <> -1 then
  begin
    exit;
  end;

  TCLO4HiMTech.FRegAppInfoB64 := TgpAppSigInfo.GetAppSigInfo2Base64ByRegPath('');

  ReportMemoryLeaksOnShutdown := DebugHook <> 0;
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(THiMTechF, HiMTechF);
  Application.CreateForm(TDataModule1, DataModule1);
  Application.Run;
end.
