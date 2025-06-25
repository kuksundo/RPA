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

function MakeHiMTechReport2ExcelByDataTypeFromGrid(AGrid: TNextGrid; ADataType: integer; ADate: TDate; ASelectedOnly: Boolean=False): integer;
function MakeDailyWorkReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid): integer;
function MakePaySlipReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid; ADate: TDate; ASelectedOnly: Boolean): integer;

function SetPayData2ExcelRptBySheetNameFromGridRow(AGrid: TNextGrid; ARow: integer; ADate: TDate; AWorkSheet: OLEVariant): integer;

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
    ShowMessage('File [' + AFileName + ']이 존재하지 않습니다');
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
    //'c:\Temp\Temp.xls' 에 LStream을 저장함
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
    ShowMessage('File [' + AFileName + ']이 존재하지 않습니다');
    exit;
  end;
end;

function MakeHiMTechReport2ExcelByDataTypeFromGrid(AGrid: TNextGrid; ADataType: integer;
  ADate: TDate; ASelectedOnly: Boolean): integer;
var
  LOutFileName, LTempFileName: string;
  LFileCopySuccess: Boolean;
  LRow: integer;
begin
  LOutFileName := GetOutFileNameByDataType(ADataType);
  EnsureDirectoryExists('c:\temp\');
  LTempFileName := 'c:\temp\' + ChangeFileExt(ExtractFileName(LOutFileName), '-' + FormatDateTime('yyyymmddhhnnss' , now) + '.xlsx');
  LTempFileName := StringReplace(LTempFileName, '-원본', '', [rfReplaceAll]);
  LFileCopySuccess := CopyFile(LOutFileName, LTempFileName, False);

  if LFileCopySuccess then
  begin
    case g_HiMTechDataType.ToType(ADataType) of
      hmtdtworkTimeTag: begin
        MakeDailyWorkReport2ExcelFromGrid(LTempFileName, AGrid);
      end;
      hmtdtPayRollSheet: begin
        MakePaySlipReport2ExcelFromGrid(LTempFileName, AGrid, ADate, ASelectedOnly);
      end;
    end; //case
  end
  else
    ShowMessage('파일 복사 실패 : [' + LOutFileName + '] --> [' + LTempFileName + ']');
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
    ShowMessage('Worksheet가 존재하지 않습니다');
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

      LCol := GetColIdxByColCaptionFromGrid(AGrid, '성명');
      LFindFromRow := 0;
      LFindFromRow := GetRowIndexFromFindNext(AGrid, LEmployeeName, LCol, LFindFromRow);

      if LFindFromRow = -1 then
        Continue;

      LWorkBegin := GetCellDataByColCaptionFromGrid(AGrid, '근무 시작', LFindFromRow);
      LWorkEnd := GetCellDataByColCaptionFromGrid(AGrid, '근무 종료', LFindFromRow);
      LWorkPeriod := GetCellDataByColCaptionFromGrid(AGrid, '근로 시간', LFindFromRow);
      LWorkOT := GetCellDataByColCaptionFromGrid(AGrid, '초과 근무', LFindFromRow);
      LAttendance := GetCellDataByColCaptionFromGrid(AGrid, '근태', LFindFromRow);

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
    ShowMessage('일일업무현황 생성 완료');
  finally
//    LExcel.WorkBooks.Close;
//    LExcel.quit;
//    LExcel:=unassigned;
  end;
end;

function MakePaySlipReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid;
  ADate: TDate;  ASelectedOnly: Boolean): integer;
var
  LWorksheet: OleVariant;
  LExcel: OleVariant;
  LWorkBook: OleVariant;

  i: integer;
  LEmployeeName: string;
begin
  Result := -1;

  if not FileExists(AOutFileName) then
  begin
    ShowMessage('File(' + AOutFileName + ')이 존재하지 않습니다');
    exit;
  end;

  LExcel := GetActiveExcelOleObject(True);
  LWorkBook := LExcel.Workbooks.Open(AOutFileName);
  LExcel.Visible := true;
//  LWorksheet := LExcel.ActiveSheet;

  for i := 0 to AGrid.RowCount - 1 do
  begin
    if ASelectedOnly then
      if not AGrid.Row[i].Selected then
        Continue;

    LEmployeeName := GetCellDataByColCaptionFromGrid(AGrid, '성명', i);
    LWorksheet := CopySheet2WorkBookByName(LWorkBook, '1', LEmployeeName);

    if VarIsNull(LWorksheet) then
    begin
      Continue;
    end;

    SetPayData2ExcelRptBySheetNameFromGridRow(AGrid, i, ADate, LWorkSheet);
  end;//for

//  try
//      LWorkBegin := GetCellDataByColCaptionFromGrid(AGrid, '근무 시작', LFindFromRow);
//      LWorkEnd := GetCellDataByColCaptionFromGrid(AGrid, '근무 종료', LFindFromRow);
//      LWorkPeriod := GetCellDataByColCaptionFromGrid(AGrid, '근로 시간', LFindFromRow);
//      LWorkOT := GetCellDataByColCaptionFromGrid(AGrid, '초과 근무', LFindFromRow);
//      LAttendance := GetCellDataByColCaptionFromGrid(AGrid, '근태', LFindFromRow);
//
//    ShowMessage('급여명세표 생성 완료');
//  finally
////    LExcel.WorkBooks.Close;
////    LExcel.quit;
////    LExcel:=unassigned;
//  end;
end;

function SetPayData2ExcelRptBySheetNameFromGridRow(AGrid: TNextGrid; ARow: integer;
  ADate: TDate; AWorkSheet: OLEVariant): integer;
var
  LRange, LRange2: OleVariant;
  i, ItemCount, LCol, LCol2, LRangeRow: integer;
  LEmployeeName, LEntryDate, LWagePerHour, LPaidDayOff, LAnnualLeave, LWeeklyLeave,
  LOverTime, LWorkHour, LWorkHour_Night, LWorkHour_Holiday, LWorkHour_Add,
  LRangeStr, LDateStr, LValue, LRangeStr2: string;
begin
  LDateStr := FormatDateTime('yyyy년 mm월 급여명세표', ADate);

  LEmployeeName := GetCellDataByColCaptionFromGrid(AGrid, '성명', ARow);
  LEntryDate := GetCellDataByColCaptionFromGrid(AGrid, '입사일자', ARow);
  LWagePerHour := GetCellDataByColCaptionFromGrid(AGrid, '시급', ARow);
  LPaidDayOff := GetCellDataByColCaptionFromGrid(AGrid, '유급휴일', ARow);
  LAnnualLeave := GetCellDataByColCaptionFromGrid(AGrid, '년차', ARow);
  LWeeklyLeave := GetCellDataByColCaptionFromGrid(AGrid, '주차', ARow);
  LOverTime := GetCellDataByColCaptionFromGrid(AGrid, '연장시간', ARow);
  LWorkHour := GetCellDataByColCaptionFromGrid(AGrid, '근로시간', ARow);
  LWorkHour_Night := GetCellDataByColCaptionFromGrid(AGrid, '야간근로시간', ARow);
  LWorkHour_Holiday := GetCellDataByColCaptionFromGrid(AGrid, '휴일근로시간', ARow);
  LWorkHour_Add := GetCellDataByColCaptionFromGrid(AGrid, '가급시간', ARow);

  LRange := AWorkSheet.range['D2'];
  LRange.FormulaR1C1 := LDateStr;
  LRange := AWorkSheet.range['F5'];
  LRange.FormulaR1C1 := LEmployeeName;
  LRange := AWorkSheet.range['O5'];
  LRange.FormulaR1C1 := LEntryDate;
  LRange := AWorkSheet.range['E7'];
  LRange.FormulaR1C1 := LWagePerHour;
  LRange := AWorkSheet.range['A9'];
  LRange.FormulaR1C1 := LWorkHour;
  LRange := AWorkSheet.range['C9'];
  LRange.FormulaR1C1 := LWorkHour_Night;
  LRange := AWorkSheet.range['E9'];
  LRange.FormulaR1C1 := LWorkHour_Holiday;
  LRange := AWorkSheet.range['G9'];
  LRange.FormulaR1C1 := LOverTime;
  LRange := AWorkSheet.range['I9'];
  LRange.FormulaR1C1 := LWorkHour_Add;
  LRange := AWorkSheet.range['K9'];
  LRange.FormulaR1C1 := LPaidDayOff;
  LRange := AWorkSheet.range['M9'];
  LRange.FormulaR1C1 := LAnnualLeave;
  LRange := AWorkSheet.range['O9'];
  LRange.FormulaR1C1 := LWeeklyLeave;

  //수당 항목(가변적임) 채우기
  LCol := GetColIdxByColCaptionFromGrid(AGrid, '시간총액');

  if LCol > 0 then
  begin
    //Header Caption이 '시간총액' 이후 '수당합계' 까지 수당 항목임
    Inc(LCol);

    //수당 항목의 마지막 Column Index 가져옴
    LCol2 := GetColIdxByColCaptionFromGrid(AGrid, '수당합계');

    //Header Caption '수당합계' 가 '시간총액' 뒤에 있어야 함
    if (LCol2 > 0) and (LCol < LCol2) then
    begin
      Dec(LCol2);
      ItemCount := LCol2 - LCol + 1;

      //수당 항목이 10개 이상이면 엑셀 Row 추가
      if ItemCount > 10 then
      begin
        for i := 10 to ItemCount do
          XlsRangeCopyNInsert2WS(AWorkSheet, 'B19:P19', 'B20:P20');
      end;

      LRangeRow := 11;

      for i := LCol to LCol2 do
      begin
        //수당 항목 기입
        LRangeStr := 'L' + IntToStr(LRangeRow);
        LRange := AWorkSheet.range[LRangeStr];
        LRange.FormulaR1C1 := AGrid.Columns.Item[i].Header.Caption;

        //수당 금액 기입
        LRangeStr := 'P' + IntToStr(LRangeRow);
        LRange := AWorkSheet.range[LRangeStr];
        LRange.FormulaR1C1 := AGrid.Cells[i, ARow];

        Inc(LRangeRow);
      end;
    end;
  end;

  //공제 항목 채우기
  LCol := GetColIdxByColCaptionFromGrid(AGrid, '총급여액');

  if LCol > 0 then
  begin
    //Header Caption이 '총급여액' 이후 Grid의 마지막 Column 까지 공제 항목임
    Inc(LCol);
    //25행부터 시작해서 '국민연금'이 있는 행을 검색함
    LRangeRow := GetRowidxByCellValueFromWS(AWorkSheet, '국민연금', 25);
    LRangeStr2 := 'A' + IntToStr(LRangeRow + 4) + ':T' + IntToStr(LRangeRow + 4);

    LCol2 := 0;

    for i := LCol to AGrid.Columns.Count - 1 do
    begin
      LDateStr := AGrid.Columns.Item[i].Header.Caption;
      LValue := AGrid.Cells[i, ARow];

      if (LDateStr = '국민연금') or (LDateStr = '건강보험') or
        (LDateStr = '고용보험') or (LDateStr = '갑근세') or (LDateStr = '주민세') then
        LRangeStr := 'F';

      if LRangeRow <> -1 then
      begin
        if LDateStr = '국민연금' then
          ItemCount := LRangeRow
        else if LDateStr = '건강보험' then
          ItemCount := LRangeRow+1
        else if LDateStr = '고용보험' then
          ItemCount := LRangeRow+2
        else if LDateStr = '갑근세' then
          ItemCount := LRangeRow+3
        else if LDateStr = '주민세' then
          ItemCount := LRangeRow+4
        else
        begin
          if LValue = '' then
            Continue;

          if LDateStr = '공제액계' then
            Break;

          ItemCount := LRangeRow + LCol2;

          //엑셀파일 내 공제내용은 기본이 5행임, 고로 5보다 크면 행을 추가 해야함
          if LCol2 > 4 then
          begin
            LRangeStr := 'A' + IntToStr(LRangeRow + LCol2-1) + ':T' + IntToStr(LRangeRow + LCol2-1);
            LRangeStr2 := 'A' + IntToStr(LRangeRow + LCol2) + ':T' + IntToStr(LRangeRow + LCol2);
            XlsRangeCopyNInsert2WS(AWorkSheet, LRangeStr, LRangeStr2);
            LRange := AWorkSheet.range[LRangeStr2];
            LRange.FormulaR1C1 := '';  //새로 추가한 셀의 내용을 비움
          end;

          //공제 항목을 엑셀에 추가함
          LRangeStr := 'K' + IntToStr(ItemCount);
          LRange := AWorkSheet.range[LRangeStr];
          LRange.FormulaR1C1 := LDateStr;

          LRangeStr := 'P';

          Inc(LCol2);
        end;

        LRangeStr := LRangeStr + IntToStr(ItemCount);
        LRange := AWorkSheet.range[LRangeStr];
        LRange.FormulaR1C1 := LValue;
      end;
    end;//for

    LRange := AWorkSheet.range[LRangeStr2];
    //xlEdgeBottom, xlContinuous, xlThick
    SetExcelCellRangeSingleBorder(LRange, 9, 1, 4);
  end;
end;

end.
