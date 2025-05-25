unit UnitHiMTechExcelUtil;

interface

uses Sysutils, Dialogs, Classes, Forms, Variants,
  NxColumnClasses, NxColumns, NxGrid, NxCells,
  mormot.core.base, mormot.core.datetime, mormot.core.variants, mormot.core.data,
  mormot.core.unicode, mormot.core.text, mormot.core.os,
  UnitStringUtil, UnitFileSearchUtil, UnitExcelUtil, UnitHiMTechData;

function GetOutFileNameByDataType(ADataType: integer): string;

function ImportWorkTimeTagData2GridFromXlsFile(AFileName: string; AGrid: TNextGrid): integer;
function ImportWorkTimeTagData2GridFromString(AStr: RawByteString; AGrid: TNextGrid): integer;
function ImportPaySlipData2GridFromXlsFile(AFileName: string; AGrid: TNextGrid): integer;

function MakeHiMTechReport2ExcelByDataTypeFromGrid(AGrid: TNextGrid; ADataType: integer; ADate: TDate): integer;
function MakeDailyWorkReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid): integer;
function MakePaySlipReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid; ADate: TDate): integer;

implementation

uses UnitmORMotUtil2, UnitDateUtil2, UnitNextGridUtil2;

function GetOutFileNameByDataType(ADataType: integer): string;
begin
  Result := ExtractFilePath(Application.ExeName) + g_HiMTechOriginalRptName.ToString(ADataType);
end;

function ImportWorkTimeTagData2GridFromXlsFile(AFileName: string; AGrid: TNextGrid): integer;
var
  LJsonAry: string;
  LDocList: IDocList;
//  LVar: variant;
begin
  if FileExists(AFileName) then
  begin
    LJsonAry := GetJsonAryBySheetNameFromExcelFile(AFileName);
    LDocList := DocList(StringToUtf8(LJsonAry));
    NextGridFromDocList(AGrid, LDocList);
  end
  else
  begin
    ShowMessage('File [' + AFileName + ']�� �������� �ʽ��ϴ�');
    exit;
  end;
end;

function ImportWorkTimeTagData2GridFromString(AStr: RawByteString; AGrid: TNextGrid): integer;
var
  LStream: TStream;
  LTmpXlsFileName: string;
begin
  Result := -1;

  try
    LStream := RawByteStringToStream(AStr);
    //'c:\Temp\Temp.xls' �� LStream�� ������
    LTmpXlsFileName := GetFileNameFromStream(LStream);
    Result := ImportWorkTimeTagData2GridFromXlsFile(LTmpXlsFileName, AGrid);
  finally
    LStream.Free
  end;
end;

function ImportPaySlipData2GridFromXlsFile(AFileName: string; AGrid: TNextGrid): integer;
var
  LJsonAry: string;
  LDocList: IDocList;
begin
  if FileExists(AFileName) then
  begin
    LJsonAry := GetJsonAryBySheetNameFromExcelFile(AFileName, SHEET_NAME_PAYROLLSHEET, 'B', '5', 'B', '7', True);
    LDocList := DocList(StringToUtf8(LJsonAry));
    NextGridFromDocList(AGrid, LDocList);
  end
  else
  begin
    ShowMessage('File [' + AFileName + ']�� �������� �ʽ��ϴ�');
    exit;
  end;
end;

function MakeHiMTechReport2ExcelByDataTypeFromGrid(AGrid: TNextGrid; ADataType: integer; ADate: TDate): integer;
var
  LOutFileName, LTempFileName: string;
  LFileCopySuccess: Boolean;
begin
  LOutFileName := GetOutFileNameByDataType(ADataType);
  EnsureDirectoryExists('c:\temp\');
  LTempFileName := 'c:\temp\' + ChangeFileExt(ExtractFileName(LOutFileName), '-' + FormatDateTime('yyyymmddhhmiss' , now) + '.xlsx');
  LTempFileName := StringReplace(LTempFileName, '-����', '', [rfReplaceAll]);
  LFileCopySuccess := CopyFile(LOutFileName, LTempFileName, False);

  if LFileCopySuccess then
  begin
    case g_HiMTechDataType.ToType(ADataType) of
      hmtdtworkTimeTag: begin
        MakeDailyWorkReport2ExcelFromGrid(LTempFileName, AGrid);
      end;
      hmtdtPayRollSheet: begin
        MakePaySlipReport2ExcelFromGrid(LTempFileName, AGrid, ADate);
      end;
    end; //case
  end
  else
    ShowMessage('���� ���� ���� : ' + LTempFileName);
end;

function MakeDailyWorkReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid): integer;
var
  LRange: OleVariant;
  LWorksheet: OleVariant;

  i, LRow, LCol, LFindFromRow: integer;
  LEmployeeName, LWorkBegin, LWorkEnd, LWorkPeriod, LWorkOT, LAttendance,
  LRangeStr: string;
  LVarAry: variant;
begin
  Result := -1;

  LWorksheet := GetWorkSheetByNameFromExcelFile(AOutFileName, '');

  if VarIsNull(LWorksheet) then
  begin
    ShowMessage('Worksheet�� �������� �ʽ��ϴ�');
    exit;
  end;

  try
    LRow := GetLastRowNumFromExcel(LWorksheet);

    for i := 7 to LRow do
    begin
      LRange := LWorksheet.range['F'+IntToStr(i)];
      LEmployeeName := LRange.FormulaR1C1;

      if LEmployeeName = '' then
        Break;
    end;

    if i >= LRow then
      exit;

    LRangeStr := 'F7:M' + IntToStr(i-1);
    LVarAry := ReadExcelRangeToVarArrayFromWorkSheet(LWorksheet, LRangeStr);

    for LRow := VarArrayLowBound(LVarAry, 1) to VarArrayHighBound(LVarAry, 1) do
    begin
      LEmployeeName := VarToStr(LVarAry[LRow, 1]);

      LCol := GetColIdxByColCaptionFromGrid(AGrid, '����');
      LFindFromRow := 0;
      LFindFromRow := GetRowIndexFromFindNext(AGrid, LEmployeeName, LCol, LFindFromRow);

      if LFindFromRow = -1 then
        Continue;

      LWorkBegin := GetCellDataByColCaptionFromGrid(AGrid, '�ٹ� ����', LFindFromRow);
      LWorkEnd := GetCellDataByColCaptionFromGrid(AGrid, '�ٹ� ����', LFindFromRow);
      LWorkPeriod := GetCellDataByColCaptionFromGrid(AGrid, '�ٷ� �ð�', LFindFromRow);
      LWorkOT := GetCellDataByColCaptionFromGrid(AGrid, '�ʰ� �ٹ�', LFindFromRow);
      LAttendance := GetCellDataByColCaptionFromGrid(AGrid, '����', LFindFromRow);

      LVarAry[LRow, 4] := LAttendance;
//      LVarAry[LRow, 3] := '';
      LVarAry[LRow, 6] := LWorkBegin;
      LVarAry[LRow, 7] := LWorkEnd;
      LVarAry[LRow, 8] := AddTimeStrings(LWorkPeriod, LWorkOT);

//      for LCol := VarArrayLowBound(LVarAry, 2) to VarArrayHighBound(LVarAry, 2) do
//      begin
//        VarToStr(LVarAry[LRow, LCol]);
//      end;
    end;//for

    LWorksheet.Range[LRangeStr].Value2 := LVarAry;
    ShowMessage('���Ͼ�����Ȳ ���� �Ϸ�');
  finally
//    LExcel.WorkBooks.Close;
//    LExcel.quit;
//    LExcel:=unassigned;
  end;
end;

function MakePaySlipReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid; ADate: TDate): integer;
var
  LRange: OleVariant;
  LWorksheet: OleVariant;
  LExcel: OleVariant;
  LWorkBook: OleVariant;

  i, LRow, LCol, LFindFromRow: integer;
  LEmployeeName, LEntryDate, LWagePerHour, LPaidDayOff, LAnnualLeave, LWeeklyLeave,
  LOverTime, LWorkHour,
  LRangeStr, LDateStr: string;
begin
  Result := -1;

  if not FileExists(AOutFileName) then
  begin
    ShowMessage('File(' + AOutFileName + ')�� �������� �ʽ��ϴ�');
    exit;
  end;

  LExcel := GetActiveExcelOleObject(True);
  LWorkBook := LExcel.Workbooks.Open(AOutFileName);
  LExcel.Visible := true;
//  LWorksheet := LExcel.ActiveSheet;

  LDateStr := FormatDateTime('yyyy�� mm�� �޿���ǥ', ADate);

  for i := 0 to AGrid.RowCount - 1 do
  begin
    LEmployeeName := GetCellDataByColCaptionFromGrid(AGrid, '����', i);
    LWorksheet := CopySheet2WorkBookByName(LWorkBook, '1', LEmployeeName);

    if VarIsNull(LWorksheet) then
    begin
      Continue;
    end;

    LEntryDate := GetCellDataByColCaptionFromGrid(AGrid, '�Ի�����', i);
    LWagePerHour := GetCellDataByColCaptionFromGrid(AGrid, '�ñ�', i);
    LPaidDayOff := GetCellDataByColCaptionFromGrid(AGrid, '��������', i);
    LAnnualLeave := GetCellDataByColCaptionFromGrid(AGrid, '����', i);
    LWeeklyLeave := GetCellDataByColCaptionFromGrid(AGrid, '����', i);
    LOverTime := GetCellDataByColCaptionFromGrid(AGrid, '����', i);
    LWorkHour := GetCellDataByColCaptionFromGrid(AGrid, '�հ�', i);

    LRange := LWorksheet.range['D2'];
    LRange.FormulaR1C1 := LDateStr;
    LRange := LWorksheet.range['F5'];
    LRange.FormulaR1C1 := LEmployeeName;
    LRange := LWorksheet.range['O5'];
    LRange.FormulaR1C1 := LEntryDate;
    LRange := LWorksheet.range['E7'];
    LRange.FormulaR1C1 := LWagePerHour;
    LRange := LWorksheet.range['K9'];
    LRange.FormulaR1C1 := LPaidDayOff;
    LRange := LWorksheet.range['M9'];
    LRange.FormulaR1C1 := LAnnualLeave;
    LRange := LWorksheet.range['O9'];
    LRange.FormulaR1C1 := LWeeklyLeave;
    LRange := LWorksheet.range['G9'];
    LRange.FormulaR1C1 := LOverTime;
    LRange := LWorksheet.range['A9'];
    LRange.FormulaR1C1 := LWorkHour;
  end;

//  try
//      LWorkBegin := GetCellDataByColCaptionFromGrid(AGrid, '�ٹ� ����', LFindFromRow);
//      LWorkEnd := GetCellDataByColCaptionFromGrid(AGrid, '�ٹ� ����', LFindFromRow);
//      LWorkPeriod := GetCellDataByColCaptionFromGrid(AGrid, '�ٷ� �ð�', LFindFromRow);
//      LWorkOT := GetCellDataByColCaptionFromGrid(AGrid, '�ʰ� �ٹ�', LFindFromRow);
//      LAttendance := GetCellDataByColCaptionFromGrid(AGrid, '����', LFindFromRow);
//
//    ShowMessage('�޿���ǥ ���� �Ϸ�');
//  finally
////    LExcel.WorkBooks.Close;
////    LExcel.quit;
////    LExcel:=unassigned;
//  end;
end;

end.
