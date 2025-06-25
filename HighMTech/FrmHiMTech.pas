unit FrmHiMTech;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Buttons, AdvEdit, AdvEdBtn, Vcl.Menus,
  Vcl.Mask, JvExMask, JvToolEdit, JvCombobox, Vcl.StdCtrls, AeroButtons,
  Vcl.ExtCtrls, Vcl.ComCtrls, AdvGroupBox, AdvOfficeButtons, AdvToolBtn,
  JvExControls, JvLabel, CurvyControls, NxColumns, NxColumnClasses,
  NxScrollControl, NxCustomGridControl, NxCustomGrid, NxGrid, AdvOfficeTabSet,
  DragDropInternet,DropSource,DragDropFile,DragDropFormats, DragDrop, DropTarget,
  mormot.core.base, mormot.core.os, mormot.core.data, mormot.core.text, mormot.core.unicode,
  EasterEgg, FormAboutDefs,
  UnitHiMTechData, UnitHiMTechCLO;

type
  THiMTechF = class(TForm)
    CurvyPanel1: TCurvyPanel;
    JvLabel5: TJvLabel;
    JvLabel4: TJvLabel;
    JvLabel8: TJvLabel;
    JvLabel10: TJvLabel;
    Panel1: TPanel;
    btn_Search: TAeroButton;
    btn_Close: TAeroButton;
    AeroButton1: TAeroButton;
    ClaimNoEdit: TEdit;
    HullNoEdit: TAdvEditBtn;
    OrderNoEdit: TAdvEditBtn;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    ShippingNoEdit: TEdit;
    TaskTab: TAdvOfficeTabSet;
    NextGrid1: TNextGrid;
    NxIncrementColumn1: TNxIncrementColumn;
    DropEmptyTarget1: TDropEmptyTarget;
    DataFormatAdapter1: TDataFormatAdapter;
    DataFormatAdapterOutlook: TDataFormatAdapter;
    DataFormatAdapterTarget: TDataFormatAdapter;
    JvLabel9: TJvLabel;
    DataTypeRG: TRadioGroup;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Help1: TMenuItem;
    Close1: TMenuItem;
    About1: TMenuItem;
    FormAbout1: TFormAbout;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    OpenDialog1: TOpenDialog;
    JvLabel1: TJvLabel;
    DateTimePicker1: TDateTimePicker;
    N5: TMenuItem;

    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure DropEmptyTarget1Drop(Sender: TObject; ShiftState: TShiftState;
      APoint: TPoint; var Effect: Integer);
    procedure btn_CloseClick(Sender: TObject);
    procedure AeroButton1Click(Sender: TObject);
    procedure NextGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Close1Click(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
  private
    FEgg: TEasternEgg;

    procedure OnEasterEgg(msg: string);
  public
    { Public declarations }
  end;

var
  HiMTechF: THiMTechF;

implementation

uses UnitDragUtil, UnitHiMTechExcelUtil, UnitExcelUtil,
  FrmHiMTechDM;

{$R *.dfm}

procedure THiMTechF.About1Click(Sender: TObject);
begin
  FormAbout1.LicenseText.Text := '이 프로그램은 HiMTech 에서만 사용 가능합니다';
  FormAbout1.Show(False);
end;

procedure THiMTechF.AeroButton1Click(Sender: TObject);
begin
  MakeHiMTechReport2ExcelByDataTypeFromGrid(NextGrid1, DataTypeRG.ItemIndex+1, DateTimePicker1.Date);
end;

procedure THiMTechF.btn_CloseClick(Sender: TObject);
begin
  Close;
end;

procedure THiMTechF.Close1Click(Sender: TObject);
begin
  Close;
end;

procedure THiMTechF.DropEmptyTarget1Drop(Sender: TObject; ShiftState: TShiftState;
  APoint: TPoint; var Effect: Integer);
var
  LTargetStream: TStream;
  LRawByte: RawByteString;
  LFromOutlook: Boolean;
  LFileName, LFileExt: string;
  LJson: RawUtf8;
begin
  LFileName := '';
  LFromOutlook := False;
  if (DataFormatAdapter1.DataFormat <> nil) then
  begin
    LFileName := (DataFormatAdapter1.DataFormat as TFileDataFormat).Files.Text;

    // OutLook에서 Drag한 경우에는 LFileName = '' 임
    if LFileName = '' then
    begin
      // OutLook에서 첨부파일을 Drag 했을 경우
      if (TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat).FileNames.Count > 0) then
      begin
        LFileName := TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat).FileNames[0];
        LFileExt := UpperCase(ExtractFileExt(LFileName));

        if (LFileExt = '.XLSX') or (LFileExt = '.XLS') or (LFileExt = '.XLSM') then
        begin
          LTargetStream := GetStreamFromDropDataFormat(TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat));
          try
            if not Assigned(LTargetStream) then
              ShowMessage('Not Assigned');

            LRawByte := StreamToRawByteString(LTargetStream);
            LFromOutlook := True;
          finally
            if Assigned(LTargetStream) then
              LTargetStream.Free;
          end;
        end;
      end;
    end
    else// 윈도우 탐색기에서 Drag 했을 경우 LFileName에 Drag한 File Name이 존재함
    begin
      LFileExt := UpperCase(ExtractFileExt(LFileName));

      if (LFileExt = '.XLSX') or (LFileExt = '.XLS') or (LFileExt = '.XLSM') then
      begin
        LRawByte := StringFromFile(LFileName);
      end;
    end;
  end;

  if LFileName <> '' then
  begin
    if Pos(XLS_NAME_WORKTIMETAG, LFileName) > 0 then
    begin
      DataTypeRG.ItemIndex := Ord(hmtdtworkTimeTag) - 1;
      ImportWorkTimeTagData2GridFromXlsFile(LFileName, NextGrid1);
    end
    else
    if Pos(XLS_NAME_PAYROLLSHEET, LFileName) > 0 then
    begin
      DataTypeRG.ItemIndex := Ord(hmtdtPayRollSheet) - 1;
      ImportPaySlipData2GridFromXlsFile(LFileName, NextGrid1);
    end;
  end;
end;

procedure THiMTechF.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  FEgg.Free;
end;

procedure THiMTechF.FormShow(Sender: TObject);
begin
  FEgg := TEasternEgg.Create('Reg', [ssCtrl], 'HIMTECH', Self, OnEasterEgg);
end;

procedure THiMTechF.N3Click(Sender: TObject);
begin
  OpenDialog1.Filter := 'xlsx (*.xlsx)|*.xlsx|xlsm (*.xlsm)|*.xlsm|xls (*.xls)|*.xls|모든 파일 (*.*)|*.*';

  if OpenDialog1.Execute() then
  begin
    DataTypeRG.ItemIndex := Ord(hmtdtworkTimeTag) - 1;
    ImportWorkTimeTagData2GridFromXlsFile(OpenDialog1.FileName, NextGrid1);
  end;
end;

procedure THiMTechF.N4Click(Sender: TObject);
begin
  OpenDialog1.Filter := 'xlsm (*.xlsm)|*.xlsm|xlsx (*.xlsx)|*.xlsx|xls (*.xls)|*.xls|모든 파일 (*.*)|*.*';

  if OpenDialog1.Execute() then
  begin
    DataTypeRG.ItemIndex := Ord(hmtdtPayRollSheet) - 1;
    ImportPaySlipData2GridFromXlsFile(OpenDialog1.FileName, NextGrid1);
  end;
end;

procedure THiMTechF.N5Click(Sender: TObject);
begin
  MakeHiMTechReport2ExcelByDataTypeFromGrid(NextGrid1, DataTypeRG.ItemIndex+1, DateTimePicker1.Date, True);
end;

procedure THiMTechF.NextGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  FEgg.CheckKeydown(Key, Shift);
end;

procedure THiMTechF.OnEasterEgg(msg: string);
begin
//  FormAbout1.LicenseText.Text := '이 프로그램은 HiMTech 에서만 사용 가능합니다';
  FileFromString(TCLO4HiMTech.FRegAppInfoB64, 'c:\temp\gpappinfo.txt');
  About1Click(nil);
end;

end.
