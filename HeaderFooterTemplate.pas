unit HeaderFooterTemplate;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes,
  System.Variants,
  FMX.Types, FMX.Graphics, FMX.Controls, FMX.Forms, FMX.Dialogs,
  FMX.StdCtrls,
  System.Rtti, FMX.Grid.Style, FMX.Grid, FMX.ScrollBox,
  FMX.Controls.Presentation, ComObj, DateUtils, FMX.Memo.Types, FMX.Memo,
  FMX.Edit, FMX.Objects, System.Math, System.ImageList, FMX.ImgList, FMX.Media,
  System.Actions, FMX.ActnList, FMX.StdActns, FMX.MediaLibrary, FMX.MultiView, System.IniFiles,
  FMX.Layouts, FMX.ExtCtrls;

type
  THeaderFooterForm = class(TForm)
    Header: TToolBar;
    Footer: TToolBar;
    HeaderLabel: TLabel;
    PanelMain: TPanel;
    SGMain: TStringGrid;
    StringColumn1: TStringColumn;
    StringColumn2: TStringColumn;
    TimeColumn: TTimeColumn;
    Button1: TButton;
    PanelEd: TPanel;
    SGEd: TStringGrid;
    StringColumn3: TStringColumn;
    SpdBHome: TSpeedButton;
    SpdBRefresh: TSpeedButton;
    PanelPlan: TPanel;
    SGPlan: TStringGrid;
    StringColumn4: TStringColumn;
    CheckColumn1: TCheckColumn;
    Timer1: TTimer;
    StringColumn5: TStringColumn;
    SpdBSave: TSpeedButton;
    PanelSettings: TPanel;
    SpdBSettings: TSpeedButton;
    SpdBSavePath: TSpeedButton;
    EditBasePath: TEdit;
    GroupBox1: TGroupBox;
    PanelTimer: TPanel;
    LabelTimer: TLabel;
    Memo1: TMemo;
    StyleBook1: TStyleBook;
    PanelClock: TPanel;
    PanelMedia: TPanel;
    MPMotivate: TMediaPlayer;
    MPMotivateControl: TMediaPlayerControl;
    GroupBox2: TGroupBox;
    TrackBVolS: TTrackBar;
    SGVideo: TStringGrid;
    SpeedButton1: TSpeedButton;
    OpenDialog1: TOpenDialog;
    StringColumn6: TStringColumn;
    StringColumn7: TStringColumn;
    StringColumn8: TStringColumn;
    Button2: TButton;
    LabelVideoName: TLabel;
    TimerVideo: TTimer;
    MultiView1: TMultiView;
    ButtonSkip: TButton;
    Label2: TLabel;
    Label1: TLabel;
    TrackBVolV: TTrackBar;
    SwitchVideo: TSwitch;
    Label3: TLabel;
    LabelNextTask: TLabel;
    TimerHistorySave: TTimer;
    ImageControl1: TImageControl;
    LabelNextTime: TLabel;
    LabelNowTime: TLabel;
    SwitchLoadMissed: TSwitch;
    Label4: TLabel;
    GroupBox3: TGroupBox;
    PanelQuote: TPanel;
    LabelQuote: TLabel;
    TimerCreate: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure LoadExcelDataToGrid(const FileName: string; SGMain: TStringGrid);
    procedure SGMainCellClick(const Column: TColumn; const Row: Integer);
    procedure SpdBRefreshClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure CheckPlan;
    procedure ExportToExcel;
    procedure SpdBSaveClick(Sender: TObject);
    procedure SpdBSettingsClick(Sender: TObject);
    procedure SpdBSavePathClick(Sender: TObject);
    procedure TimerGenerator;
    procedure SGPlanCellClick(const Column: TColumn; const Row: Integer);
    procedure SpeedButton1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure SGVideoCellClick(const Column: TColumn; const Row: Integer);
    procedure CheckAndPlayVideo;
    procedure SpdBHomeClick(Sender: TObject);
    procedure TimerVideoTimer(Sender: TObject);
    procedure ButtonSkipClick(Sender: TObject);
    procedure TimerHistorySaveTimer(Sender: TObject);
    procedure LoadSettingsFromIni(const FileName: string);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure LoadRandomQuoteToLabel;
    procedure ImageControl1Click(Sender: TObject);
    procedure TimerCreateTimer(Sender: TObject);
  private
    RectangleArray: array [1 .. 30] of TRectangle;
  public
    { Public declarations }
  end;

var
  HeaderFooterForm: THeaderFooterForm;
  BaseName: string = 'res\Schedule.xlsx';
  HistoryBaseName: string = 'res\History.xlsx';
  NextTime: string = '';
  WorksNames: string = '';
  ExePath: string;
  TargetRow: integer = -1;
  MusicCheck: string;
  SoundFilePath, VideoFilePath: string;
  DeltaTime: TTime;
  FlagVideo: Boolean;
  saveHeader: string;
  IniFileName: string = 'res\settings.ini';
  NextTimeTimer: TTime;
  RectCount: Real;
  RectColored: Int8 = 30;
  TimerCreateBool: Boolean;
  TimerRefresh: Boolean = True;
  ExcelApp: OleVariant;


implementation

{$R *.fmx}

procedure THeaderFooterForm.Button1Click(Sender: TObject);
begin
  PanelPlan.Visible := True;
  // ExportToExcel;
end;

procedure THeaderFooterForm.Button2Click(Sender: TObject);
begin
//  PanelMedia.Visible:=True;
  Timer1.Enabled := False;
end;

procedure THeaderFooterForm.ButtonSkipClick(Sender: TObject);
begin
  PanelMedia.Visible := False;
  MPMotivate.Stop;
  TimerVideo.Enabled := False;
  ButtonSkip.Visible := False;
end;

procedure THeaderFooterForm.CheckAndPlayVideo;
var
  I, J, ColumnIndex: Integer;
  NonEmptyRowCount: Integer;
  SuitableColumns: array of Integer;
begin
  ExePath := ExtractFilePath(ParamStr(0));
  Randomize;
  SetLength(SuitableColumns, 0);
  NonEmptyRowCount := 0;
  if TargetRow = -1 then Exit;
  Memo1.Lines.Add('Tаймер запущен. TargetRow: ' + IntToStr(TargetRow));

  // 1. Проверка наличия всех заголовков таблицы SGVideo в SGPlan.Cells[0,2]
  for I := 0 to SGVideo.ColumnCount - 1 do
  begin
    if pos(SGVideo.Columns[I].Header, SGMain.Cells[1, TargetRow]) <> 0 then
    begin
      for J := 0 to SGVideo.RowCount - 1 do
      begin
        if Trim(SGVideo.Cells[i, j]) = '' then  Break;
          Inc(NonEmptyRowCount);
      end;
      if NonEmptyRowCount <> 0 then begin
        SetLength(SuitableColumns, Length(SuitableColumns) + 1);
        SuitableColumns[High(SuitableColumns)] := I;
      end;
    end;
  end;
  Memo1.Lines.Add('Columns: ' + IntToStr(Length(SuitableColumns)));

  //Если видео нет, то воспроизводим звук и выходим
  if Length(SuitableColumns) = 0 then
  begin
      MPMotivate.FileName := SoundFilePath;
      MPMotivate.Volume := TrackBVolS.Value;
      MPMotivate.Play;
    TimerVideo.Enabled := False;
    Exit; // Если подходящих столбцов нет, выходим из процедуры
  end;
  // 2. Случайный выбор одного из подходящих столбцов
  ColumnIndex := SuitableColumns[Random(Length(SuitableColumns))];

  Memo1.Lines.Add('RandomColumnIndex: ' + IntToStr(ColumnIndex));

  // 5. Выбор случайной не пустой строки в столбце

  repeat
    J := Random(NonEmptyRowCount-1);
  until SGVideo.Cells[ColumnIndex, J] <> '';

  Memo1.Lines.Add('RandomRowIndex: ' + IntToStr(J));
  // When the current media ends the player starts playing
  // the next media in playlist

  VideoFilePath := ExePath + 'res\Video\' + SGVideo.Cells[ColumnIndex, J];

  MusicCheck := 'NoSound';

  ButtonSkip.Visible := True;
  TimerVideo.Enabled := True;
end;


procedure THeaderFooterForm.CheckPlan;
var
  CurrentTime: TTime;
  TargetTime: TTime;
  I, j: Integer;
  HeaderText: string;
  lastRowMain: Integer;
  lastTaskTime: TTime;
  TargetWorks: array of string;
begin
  if SGMain.Cells[2,0] = '' then Exit; //Если расписание не загружно то выходим

  lastTaskTime := StrToTime ('00:30');
  lastRowMain := SGMain.RowCount - 1;

  if NextTime <> '' then
    if (StrToTime(NextTime) > (StrToTime(SGMain.Cells[2, lastRowMain]) + lastTaskTime)) and (TargetRow = lastRowMain) then
    begin
      HeaderLabel.Text:= 'Конец расписания';
      LabelNextTask.Text := 'Начало расписания в ' + SGMain.Cells[2, 0];
      TargetRow := SGMain.RowCount - 1;
      ImageControl1.Visible := True;
      Timer1.Enabled := False;
      Exit;
    end;

  if Time <= DeltaTime then
  begin
    Memo1.Lines.Add('DeltaTime: ' + TimeToStr(DeltaTime) + ' - Time: ' + TimeToStr(Time) + ' - 23:00 + 01:00');
    CurrentTime := (DeltaTime - Time);
    CurrentTime := StrToTime('23:00') - CurrentTime + StrToTime('01:00'); //если текущее время меньше дельты то высчитываем время назад от получночи
  end
  else
    CurrentTime := Time - DeltaTime;
  //Memo1.Lines.Add('Time: ' + TimeToStr(Time) + ' - DeltaTime: ' + TimeToStr(DeltaTime) + ' = ' + TimeToStr(CurrentTime));

  TargetRow := -1;
  // Инициализация переменной для строки, в которой найдено совпадение
  TargetTime := StrToTime('00:00');

  // Шаг 1: Вычисление строки, соответствующей текущему времени
  for I := 0 to SGMain.RowCount - 1 do
  begin
    if StrToTime(SGMain.Cells[2, I]) < DeltaTime then
    begin
      TargetTime := (DeltaTime - StrToTime(SGMain.Cells[2, I]));
      TargetTime := StrToTime('23:00')- TargetTime + StrToTime('01:00')
    end
    else
      TargetTime := StrToTime(SGMain.Cells[2, I]) - DeltaTime;

    if CurrentTime < TargetTime then
    begin
      Memo1.Lines.Add('TargetTime - Delta: ' + TimeToStr(TargetTime));
      TargetRow := I - 1;
      Break;
    end
    else if (CurrentTime  < TargetTime + lastTaskTime) and (I >= lastRowMain) then
      TargetRow := lastRowMain;
  end;

  Memo1.Lines.Add('Перед условиями');
  Memo1.Lines.Add('CurrentTime: ' + TimeToStr(CurrentTime));
  Memo1.Lines.Add('TargetTime: ' + TimeToStr(TargetTime));
  Memo1.Lines.Add('TargetRow: ' + IntToStr(TargetRow));


  if TargetRow = -1 then begin
      HeaderLabel.Text:= 'Здравствуйте';
      LabelNextTask.Text := 'Начало расписания в ' + SGMain.Cells[2, 0];
      ImageControl1.Visible := True;
      TargetRow := SGMain.RowCount -1;
      Timer1.Enabled := False;
      Exit;
  end;

  if TargetRow = lastRowMain then begin
    HeaderLabel.Text := SGMain.Cells[1, lastRowMain];
    LabelNextTask.Text := 'Далее: Конец расписания';
    NextTime := TimeToStr(StrToTime(SGMain.Cells[2, lastRowMain]) + lastTaskTime);
  end;

  if TargetRow <> SGMain.RowCount -1 then
  begin
    LabelNextTask.Text := 'Далее: ' + SGMain.Cells[1, TargetRow + 1];
    NextTime := SGMain.Cells[2, TargetRow + 1];
  end;

  // Записываем следующее время для сравнения в таймере
  Memo1.Lines.Add('NowTime: ' + SGMain.Cells[2, TargetRow]);
  Memo1.Lines.Add('NextTime: ' + NextTime);

  LabelNextTime.Text := 'NextTime: ' + NextTime;
  LabelNowTime.Text := 'NowTime: ' + SGMain.Cells[2, TargetRow];


  // Шаг 2: Поиск и обработка данных
  // В WorksNames Записаны названия задач которые сейчас активны, Если текущая задача продолжается, то итерация пропускается
  // По нажатию Save WorkNames очищается

  //Удаляем строку добавления задачи на лету
  if SGPlan.RowCount > 0 then SGPlan.RowCount := SGPlan.RowCount - 1;

  SetLength(TargetWorks, TargetRow + 1);

  if (Pos('Здравствуйте', HeaderLabel.Text) <> 0) and (SwitchLoadMissed.IsChecked = True) then
  begin
    for i := 0 to TargetRow do
      TargetWorks[i] := SGMain.Cells[1, i];
  end
  else
  begin
    TargetWorks[0] := SGMain.Cells[1, TargetRow];
  end;

  HeaderLabel.Text := SGMain.Cells[1, TargetRow];

  //Memo1.Lines.Add('TargetWorks: ' + string.Join(', ', TargetWorks));

  for j := 0 to SGEd.ColumnCount - 1 do
  begin
    HeaderText := SGEd.Columns[j].Header;
    //Memo1.Lines.Add('HeaderText: ' + HeaderText);
    // Проверка наличия HeaderText в массиве TargetWorks
    if (Pos(HeaderText, string.Join(', ', TargetWorks)) <> 0) then
    begin
      //Memo1.Lines.Add('(Pos(HeaderText, string.Join('', '', TargetWorks)) <> 0) ');
      //Memo1.Lines.Add('WorksNames: ' + WorksNames);
      //Если работа уже есть в списке, то пропускаем
      if (Pos(HeaderText, WorksNames) <> 0) then Break;

      //Memo1.Lines.Add('Found: ' + IntToStr(Found));
      WorksNames := WorksNames + ',/,' + HeaderText;
      //Memo1.Lines.Add(InttoStr(Found));
      // Копирование непустых строк из SGEd в SGPlan
      for I := 0 to SGEd.RowCount - 1 do
      begin
        if SGEd.Cells[j, I].Trim <> '' then
        begin
          // Добавление новой строки в SGPlan
          SGPlan.RowCount := SGPlan.RowCount + 1;
          SGPlan.Cells[2, SGPlan.RowCount - 1] := HeaderText;
          SGPlan.Cells[0, SGPlan.RowCount - 1] := SGEd.Cells[j, I];
        end
        else
        begin
          // Прекращение копирования, если значение становится пустым
          Break;
        end;
      end;
    end;

  end;


  for i := 0 to Length(TargetWorks) - 1  do
  begin
    if (Pos(TargetWorks[i], WorksNames) = 0) then
    begin
      //Memo1.Lines.Add('Pos(TargetWorks[i], WorksNames) = 0)');
      //Memo1.Lines.Add('TargetWorks[i]: ' + TargetWorks[i]);
      SGPlan.RowCount := SGPlan.RowCount + 1;
      SGPlan.Cells[0, SGPlan.RowCount - 1] := TargetWorks[i];
      WorksNames := WorksNames + ',/,' + TargetWorks[i];
    end;
  end;

  //Добавление задачи на лету

  for i := SGPlan.RowCount - 1 downto 0  do
    if Trim(SGPlan.Cells[0, i]) = '' then SGPlan.RowCount := SGPlan.RowCount - 1;


  SGPlan.RowCount := SGPlan.RowCount + 1;
  SGPlan.Cells[0, SGPlan.RowCount - 1] := '       ...Нажимите чтоб добавить задачу...';

  Memo1.Lines.Add(WorksNames);

  ImageControl1.Visible := False;

  if FlagVideo = True then CheckAndPlayVideo;  //в зависимости кто вызывает процедуру запускается видео

  //Преобразование времени для таймера
  NextTimeTimer := StrToTime(NextTime);

  Memo1.Lines.Add('NextTimeTimer: ' + TimeToStr(NextTimeTimer));
  Memo1.Lines.Add('DeltaTime: ' + TimeToStr(DeltaTime));

  if NextTimeTimer <= DeltaTime + lastTaskTime + StrToTime('00:01') then
  begin
    Memo1.Lines.Add('Таймер. NextTimeTimer <= DeltaTime ');
    Memo1.Lines.Add('DeltaTime: ' + TimeToStr(DeltaTime) + ' - Time: ' + TimeToStr(NextTimeTimer) + ' - 23:00 + 01:00');
    NextTimeTimer := (DeltaTime - NextTimeTimer);
    NextTimeTimer := StrToTime('23:00') - NextTimeTimer + StrToTime('01:00'); //если текущее время меньше дельты то высчитываем время назад от получночи
  end
  else
    NextTimeTimer := NextTimeTimer - DeltaTime;

  Memo1.Lines.Add('NextTimeTimer: ' + TimeToStr(NextTimeTimer));

  if FlagVideo = False then TimerCreateBool := False
  else TimerCreateBool := True;

  Memo1.Lines.Add('Timer1.Enabled := True;');

  Timer1.Enabled := True;
  Timer1Timer(Timer1);
end;

procedure THeaderFooterForm.ExportToExcel;
var
  ExcelApp: OleVariant;
  Workbook: OleVariant;
  Sheet: OleVariant;
  I, LastRow: Integer;
begin
  try
    // Создание экземпляра Excel
    Workbook := ExcelApp.Workbooks.Open(ExePath + HistoryBaseName);
    Sheet := Workbook.Worksheets[1]; // Лист 1

    // Находим последнюю заполненную строку на листе 1
    LastRow := Sheet.UsedRange.Rows.Count + 1;

    // Проход по строкам SGPlan
    for I := 0 to SGPlan.RowCount - 1 do
    begin
      if SGPlan.Cells[1, I] = 'True' then // Проверка чекбокса
      begin
        Sheet.Cells[LastRow, 1] := Date; // Текущая дата
        Sheet.Cells[LastRow, 2] := Time;
        // Текущее время
        Sheet.Cells[LastRow, 3] := SGPlan.Cells[2, I];
        // Данные из SGPlan столбец 2
        Sheet.Cells[LastRow, 4] := SGPlan.Cells[0, I];
        // Данные из SGPlan столбец 0
        inc(LastRow);
        //Memo1.Lines.Add('LastRow: ' + InttoStr(LastRow));
        //Memo1.Lines.Add('SGPlan.Cells[2, i]: ' + SGPlan.Cells[2, I] +
        //  'SGPlan.Cells[0, i]: ' + SGPlan.Cells[0, I]);
      end;
    end;

    TimerHistorySave.Enabled := True;
    //Memo1.Lines.Add('Данные записаны');
    // Сохранение и закрытие книги
    Workbook.Save;

  except
    on E: Exception do

  end;
end;

procedure THeaderFooterForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  SpdBSavePath.OnClick(SpdBSavePath);
  ExcelApp.Quit;
  ExcelApp := Unassigned;
end;

procedure THeaderFooterForm.FormCreate(Sender: TObject);
begin
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Visible := False;
  MusicCheck := 'NoSound';
  SoundFilePath := ExePath + 'res\Sound\check.wav';
  DeltaTime := StrToTime('00:00');
  ExePath := ExtractFilePath(ParamStr(0));
  LoadSettingsFromIni(ExePath + IniFileName);
  BaseName := ExePath + BaseName;
  TimerGenerator;
  SGMain.Columns[1].Width := SGMain.Width - 70 - 30 - 22;
  SGPlan.RowCount := 1;
  SGPlan.Cells[0, 0] := '       ...Нажимите чтоб добавить задачу...';
  ImageControl1.Visible := True;
  LoadRandomQuoteToLabel;
  // SGEd.Columns[0].Width:=SGEd.Width-20;
  // LoadExcelDataToGrid(BaseName, SGMain);
end;

procedure THeaderFooterForm.FormResize(Sender: TObject);
begin
  SGMain.Columns[1].Width := SGMain.Width - 70 - 30 - 22;
  // SGEd.Columns[0].Width:=SGEd.Width-20;
end;

procedure THeaderFooterForm.ImageControl1Click(Sender: TObject);
begin
  exit;
end;

procedure THeaderFooterForm.LoadExcelDataToGrid(const FileName: string;
  SGMain: TStringGrid);
var
  Workbook: OleVariant;
  Sheet: OleVariant;
  I, j, DayOfWeek: Integer;
  TotalRows: Integer;
  ColCount: Integer;
  Column: TStringColumn;
begin
  // Определение дня недели
  DayOfWeek := DayOfTheWeek(Now);
  // Текущий день недели (1 - Понедельник, ..., 7 - Воскресенье)

  // Создание экземпляра Excel
  try
    Workbook := ExcelApp.Workbooks.Open(FileName);
    Sheet := Workbook.Worksheets[1];

    // Определение общего количества строк на листе
    TotalRows := Sheet.UsedRange.Rows.Count;

    // Установка количества строк в StringGrid
    SGMain.RowCount := TotalRows - 4;
    // На основе вашего кода, начиная с 5 строки

    // Чтение данных из Excel и загрузка их в StringGrid
    for I := 5 to TotalRows do // Начиная с 5 строки
    begin
      SGMain.Cells[0, I - 5] := '>';
      SGMain.Cells[1, I - 5] := Sheet.Cells[I, DayOfWeek + 2].Value;
      // Расписание для текущего дня недели
      SGMain.Cells[2, I - 5] := FormatDateTime('hh:nn',
        Sheet.Cells[I, 2].Value);
      // Время в столбец 3
    end;

    ///////////// Загрузка плана ////////////////
    Sheet := Workbook.Worksheets[2]; // Лист 2

    // Определение количества столбцов на листе
    ColCount := Sheet.UsedRange.Columns.Count;

    SGEd.ClearColumns;

    for I := 0 to ColCount - 1 do
    begin
      Column := TStringColumn.Create(SGEd);
      Column.Header := 'Column ' + InttoStr(I + 1);
      Column.Width := 360;
      SGEd.AddObject(Column);
    end;

    // Создание столбцов в StringGrid и установка заголовков
    for j := 1 to ColCount do
    begin
      SGEd.Columns[j - 1].Header := Sheet.Cells[1, j].Value;
      // Названия столбцов из строки 1
    end;

    // Заполнение StringGrid данными из Excel
    for I := 2 to Sheet.UsedRange.Rows.Count do
      // Начиная со строки 2
      for j := 1 to ColCount do
        SGEd.Cells[j - 1, I - 2] := Sheet.Cells[I, j].Value;

    ///////////// Загрузка списка видео ////////////////
    Sheet := Workbook.Worksheets[4]; // Лист 4

    // Определение количества столбцов на листе
    ColCount := Sheet.UsedRange.Columns.Count;

    SGVideo.ClearColumns;

    for I := 0 to ColCount - 1 do
    begin
      Column := TStringColumn.Create(SGEd);
      Column.Width := 360;
      SGVideo.AddObject(Column);
    end;

    SGVideo.RowCount := Sheet.UsedRange.Rows.Count;

    // Создание столбцов в StringGrid и установка заголовков
    for j := 1 to ColCount do
    begin
      SGVideo.Columns[j - 1].Header := Sheet.Cells[1, j].Value;
      // Названия столбцов из строки 1
    end;

    // Заполнение StringGrid данными из Excel
    for I := 2 to Sheet.UsedRange.Rows.Count do
      // Начиная со строки 2
      for j := 1 to ColCount-1 do
        SGVideo.Cells[j - 1, i - 2] := Sheet.Cells[i, j].Value;


    // Закрытие и освобождение ресурсов
    Workbook.Close(False);


  except
    HeaderLabel.text := 'Не удалось загрузить расписание';
  end;

end;

procedure THeaderFooterForm.LoadRandomQuoteToLabel;
var
  QuoteList: TStringList;
  QuoteFileName: string;
begin
  QuoteFileName := ExePath + 'res\quotes.txt';

  // Проверяем, существует ли файл
  if not FileExists(QuoteFileName) then
  begin
    Exit;
  end;

  // Создаем объект TStringList и загружаем строки из файла
  QuoteList := TStringList.Create;
  try
    QuoteList.LoadFromFile(QuoteFileName);

    // Проверяем, есть ли строки в файле
    if QuoteList.Count > 0 then
    begin
      // Выбираем случайную строку и загружаем ее в Label
      LabelQuote.Text := QuoteList[Random(QuoteList.Count)];
    end
    else
    begin
      //ShowMessage('Файл с цитатами пуст: ' + QuoteFileName);
    end;
  finally
    // Освобождаем ресурсы TStringList
    QuoteList.Free;
  end;
end;

procedure THeaderFooterForm.LoadSettingsFromIni(const FileName: string);
var
  IniFile: TIniFile;
begin
  IniFile := TIniFile.Create(FileName);
  try
    // Загрузка настроек
    EditBasePath.Text := IniFile.ReadString('Settings', 'BasePath', '');
    TrackBVolS.Value := IniFile.ReadFloat('Settings', 'VolumeS', 0);
    TrackBVolV.Value := IniFile.ReadFloat('Settings', 'VolumeV', 0);
    SwitchVideo.IsChecked := IniFile.ReadBool('Settings', 'SwitchVideo', False);
    SwitchLoadMissed.IsChecked := IniFile.ReadBool('Settings', 'SwitchLoadMissed', True);
  finally
    IniFile.Free;
  end;
end;

procedure THeaderFooterForm.SGMainCellClick(const Column: TColumn;
  const Row: Integer);
begin
  if Column.Index = 0 then
   PanelEd.Visible := True;

end;

procedure THeaderFooterForm.SGPlanCellClick(const Column: TColumn;
  const Row: Integer);

var
  textAdd: string;
begin

  if (Row = SGPlan.RowCount - 1) and (Column.Index = 0) then
  begin
    textAdd := SGPlan.Cells[0, SGPlan.RowCount -1];
    SGPlan.Cells[0, SGPlan.RowCount -1] := '';
    SGPlan.RowCount := SGPlan.RowCount + 1;
    SGPlan.Cells[0, SGPlan.RowCount -1] := textAdd;
    Exit;
  end;

  if (Column.Index <> 1)  then Exit;
  if Row = SGPlan.RowCount - 1 then Exit;


  If SGPlan.Cells[1, Row] = 'True' Then
    SGPlan.Cells[1, Row] := 'False'
  else
    SGPlan.Cells[1, Row] := 'True';
end;

procedure THeaderFooterForm.SGVideoCellClick(const Column: TColumn;
  const Row: Integer);
var
  i:integer;
begin
  if Column.Index <> 2 then Exit;

  for I := Row to SGVideo.RowCount - 2 do
    begin
      SGVideo.Cells[0, I] := SGVideo.Cells[0, I + 1];
      SGVideo.Cells[1, I] := SGVideo.Cells[1, I + 1];
    end;

end;

procedure THeaderFooterForm.SpdBHomeClick(Sender: TObject);
begin
  PanelEd.Visible := False;
  PanelPlan.Visible := False;
  PanelSettings.Visible := False;
end;

procedure THeaderFooterForm.SpdBRefreshClick(Sender: TObject);
var
  FirstTime, LastTime : TTime;
begin
  Memo1.Text := '';
  WorksNames := ''; //добавил очистку списка работ при обновлении
  SGPlan.RowCount := 1;
  SGPlan.Cells[0, 0] := '       ...Нажимите чтоб добавить задачу...';
  FlagVideo := True;
  LoadExcelDataToGrid(BaseName, SGMain);
  //Добавить дельту времени если расписание после полуночи
  FirstTime := StrToTime(SGMain.Cells[2, 0]);
  LastTime := StrToTime(SGMain.Cells[2, SGMain.RowCount - 1]);
  if LastTime < FirstTime then
    DeltaTime := LastTime + StrToTime('01:00');
  Memo1.Lines.Add('DeltaTime: ' + TimeToStr(DeltaTime));
  CheckPlan;
end;

procedure THeaderFooterForm.SpdBSaveClick(Sender: TObject);
var
  i, checkedCount:integer;

begin
  if SGPlan.RowCount < 2 then Exit;

  checkedCount := 0;

  //Проверяем есть ли выполненные пункты плана
  for I := 0 to SGPlan.RowCount -2 do
    if SGPlan.Cells[1, i] = 'True' then inc(checkedCount);

  if checkedCount > 0 then ExportToExcel;


  FlagVideo := False;
  SGPlan.RowCount := 0;
  SGPlan.RowCount := 1;
  SGPlan.Cells[0, 0] := '       ...Нажимите чтоб добавить задачу...';
  WorksNames := '';
  CheckPlan;

  //Если нет галочек то сохранять нечего

end;

procedure THeaderFooterForm.SpdBSavePathClick(Sender: TObject);
var
  IniFile: TIniFile;
begin
  IniFile := TIniFile.Create(ExePath + IniFileName);
  try
    // Сохранение настроек
    IniFile.WriteString('Settings', 'BasePath', EditBasePath.Text);
    IniFile.WriteFloat('Settings', 'VolumeS', TrackBVolS.Value);
    IniFile.WriteFloat('Settings', 'VolumeV', TrackBVolV.Value);
    IniFile.WriteBool('Settings', 'SwitchVideo', SwitchVideo.IsChecked);
    IniFile.WriteBool('Settings', 'SwitchLoadMissed', SwitchLoadMissed.IsChecked);
  finally
    IniFile.Free;
  end;

  BaseName := EditBasePath.Text;
end;


procedure THeaderFooterForm.SpdBSettingsClick(Sender: TObject);
begin
  PanelMedia.Visible := False;
  if PanelSettings.Visible = False then
    PanelSettings.Visible := True
  else
    PanelSettings.Visible := False;
end;

procedure THeaderFooterForm.SpeedButton1Click(Sender: TObject);
var
  Files: TStrings;
  I: integer;
begin
  // Sets the Filter so only the supported files to be displayed
  OpenDialog1.Filter := TMediaCodecManager.GetFilterString;
  if (OpenDialog1.Execute) then
  begin
    Files := OpenDialog1.Files;
    for I := 0 to Files.Count - 1 do
    begin
      SGVideo.RowCount:=SGVideo.RowCount + 1;
      SGVideo.Cells[1,SGVideo.RowCount-1]:=extractFileName(Files[I]);
      SGVideo.Cells[2,SGVideo.RowCount-1]:='X';
    end;
  end;
end;


procedure THeaderFooterForm.Timer1Timer(Sender: TObject);
var
  TotalSeconds, I: Integer;
  CurrentSeconds: Integer;
  TargetTime, CurrentTimeTimer: TTime;
begin
//  Memo1.Lines.Add('Таймер запущен');

  for I := 30 downto 1 do
    RectangleArray[I].Fill.Color := TAlphaColors.White;

  {if NextTime = '' then begin
    Timer1.Enabled := False;
    Exit;
  end;  }

  if Time <= DeltaTime then
  begin
    CurrentTimeTimer := DeltaTime - Time;
    CurrentTimeTimer := StrToTime('23:00') - CurrentTimeTimer + StrToTime('01:00'); //если текущее время меньше дельты то высчитываем время назад от получночи
  end
  else
    CurrentTimeTimer := Time - DeltaTime;

  TargetTime := StrToTime(SGMain.Cells[2, TargetRow]);
  TotalSeconds := SecondsBetween(NextTimeTimer + DeltaTime, TargetTime); // Общее количество секунд таймера
  if TotalSeconds > 86300  then TotalSeconds := TotalSeconds - 86400;

  CurrentSeconds := TotalSeconds - SecondsBetween(NextTimeTimer, CurrentTimeTimer);
  RectCount := Int((CurrentSeconds / TotalSeconds) * 30);

 { Memo1.Lines.Add('///////Timer//////');
  Memo1.Lines.Add('NextTimeTimer: ' + TimeToStr(NextTimeTimer));
  Memo1.Lines.Add('CurrentTimeTimer: ' + TimeToStr(CurrentTimeTimer));
  Memo1.Lines.Add('TargetTime: ' + TimeToStr(TargetTime));
  Memo1.Lines.Add('NextTimeTimer + DeltaTime: ' + TimeToStr(NextTimeTimer + DeltaTime));
  Memo1.Lines.Add('TotalSeconds: ' + IntToStr(TotalSeconds));
  Memo1.Lines.Add('CurrentSeconds: ' + IntToStr(CurrentSeconds));
  Memo1.Lines.Add('RectCount: ' + FloatToStr(RectCount));
  Memo1.Lines.Add('..........................');  }
  // Оставшееся количество минут до NextTime

  // Memo1.Lines.Add('TotalMinutes: ' + IntToStr(TotalMinutes) + ' CurrentMinute: ' + IntToStr(CurrentMinute));

  //Плавное обнуление таймера
  {if (TimerCreateBool) and (LabelTimer.Text <> '00:00') then
  begin
    LabelTimer.Text := FormatDateTime('nn:ss', NextTimeTimer - CurrentTimeTimer);
    TimerRefresh := True;
    RectColored := 1;
  end
  else
    TimerRefresh := False;    }


  if TimerCreateBool then
  begin
    LabelTimer.Text := FormatDateTime('nn:ss', NextTimeTimer - CurrentTimeTimer);
    TimerCreate.Enabled := True;
    Timer1.Enabled := false;
    Exit;
  end;


  // Закрашиваем прямоугольники в соответствии с текущим временем
  for I := 30 downto 1 do
  begin
    if I <= RectCount then
      RectangleArray[I].Fill.Color := TAlphaColors.White
      // RectangleArray - массив прямоугольников
    else
      RectangleArray[I].Fill.Color := TAlphaColors.Dodgerblue;
    // Незакрашенные прямоугольники
  end;

  {if LabelTimer.Text = '00:00' then
    for I := 30 downto 1 do
      RectangleArray[I].Fill.Color := TAlphaColors.White;  }

  //Memo1.Lines.Add('NextTimeTimer: ' + TimeToStr(NextTimeTimer));
  //Memo1.Lines.Add('CurrentTimeTimer: ' + TimeToStr(CurrentTimeTimer));

  if NextTimeTimer > CurrentTimeTimer then
  begin
    LabelTimer.Text := FormatDateTime('nn:ss', NextTimeTimer - CurrentTimeTimer);
    Exit;
  end
  else
  begin
    Memo1.Text := '';
    Memo1.Lines.Add('Настало время идти дальше NextTime: ' + NextTime);
    FlagVideo := True;
    CheckPlan;
    RectColored := 30;
    Timer1.Enabled := False;
    //CheckAndPlayVideo;
  end;
end;

procedure THeaderFooterForm.TimerGenerator;
var
  I: Integer;
  Rect: TRectangle;
  Radius: Integer;
  CenterX, CenterY: Integer;
begin
  Radius := 120; // Радиус циферблата
  CenterX := Round(PanelClock.Width / 2);
  CenterY := Round(PanelClock.Height / 2);

  for I := 1 to 30 do
  begin
    Rect := TRectangle.Create(Self);
    Rect.Parent := PanelClock; // Установка родительской панели
    Rect.Width := 12;
    Rect.Height := 30;
    Rect.Position.X := CenterX + Radius * Sin(DegToRad(I * 12)) -
      Rect.Width / 2;
    Rect.Position.Y := CenterY - Radius * Cos(DegToRad(I * 12)) -
      Rect.Height / 2;
    Rect.RotationAngle := I * 12; // Угол поворота
    Rect.Fill.Color := TAlphaColors.White;
    Rect.Stroke.Color := TAlphaColors.Deepskyblue;
    Rect.Stroke.Thickness := 0.5;

    RectangleArray[I] := Rect; // Добавление прямоугольника в массив
  end;
end;

procedure THeaderFooterForm.TimerHistorySaveTimer(Sender: TObject);
begin
  if saveHeader = '' then
  begin
    saveHeader := HeaderLabel.Text;
    HeaderLabel.Text := 'Данные сохранены';
  end
  else
  begin
    HeaderLabel.Text := saveHeader;
    saveHeader := '';
    TimerHistorySave.Enabled := False;
  end;

end;

procedure THeaderFooterForm.TimerCreateTimer(Sender: TObject);
begin
 { if TimerRefresh then
  begin
    //Плавное обновление таймера
    if RectColored < 31 then
    begin
      // Продолжаем изменение цвета прямоугольников
      RectangleArray[RectColored].Fill.Color := TAlphaColors.White;
      RectColored := RectColored + 1;
    end
    else
    begin
      // Завершаем изменение цвета и возвращаемся к основному таймеру
      TimerRefresh := False;
      //LabelTimer.Text := '00:00';
      RectColored := 30;
    end;
    Exit;
  end;    }

 { Memo1.Lines.Add('/////CreateTimer///////');
  Memo1.Lines.Add('RectCount: ' + FloatToStr(RectCount));
  Memo1.Lines.Add('RectColored: ' + FloatToStr(RectColored));
  Memo1.Lines.Add('.........................');   }

 If RectColored > RectCount Then
  begin
    RectangleArray[RectColored].Fill.Color := TAlphaColors.Dodgerblue;
    RectColored := RectColored -1;
  end
  else begin
    TimerCreateBool := False;
    TimerCreate.Enabled := False;
    Timer1.Enabled := True;
    RectColored := 30;
    TimerRefresh := True;
  end;
end;

procedure THeaderFooterForm.TimerVideoTimer(Sender: TObject);
begin
  //Memo1.Text := '';
  {Memo1.Lines.Add('TimerVideoTimer');
  Memo1.Lines.Add(SoundFilePath);
  Memo1.Lines.Add(VideoFilePath);
  Memo1.Lines.Add(MusicCheck);}

  If (MusicCheck = 'Sound') and (MPMotivate.State = TMediaState.Stopped) then
    MusicCheck := 'SoundCheck';
  If (MusicCheck = 'Video') and (MPMotivate.State = TMediaState.Stopped) then
  begin
    MusicCheck := 'VideoCheck';
    PanelMedia.Visible := False;
    TimerVideo.Enabled := False;
  end;

  If MusicCheck = 'VideoCheck' then
  begin
    ButtonSkip.Visible := False;
    TimerVideo.Enabled := False;
    Exit;
  end;


  if MusicCheck = 'Video' then Exit;
  if MusicCheck = 'Sound' then Exit;

  If MusicCheck = 'NoSound' then begin
      // 6. Воспроизведение звука
    MPMotivate.Clear;
    try
      MPMotivate.FileName := SoundFilePath;
      MPMotivate.Volume := TrackBVolS.Value;
      MPMotivate.Play;
      MusicCheck := 'Sound';
      //Memo1.Lines.Add('Аудио запущено')
    except
      //Memo1.Lines.Add('Аудио не найдено');
      MusicCheck := 'SoundCheck';
    end;
  end;


  if MusicCheck = 'SoundCheck' then begin
    if SwitchVideo.IsChecked = False then begin
      MusicCheck := 'VideoCheck';
      ButtonSkip.Visible := False;
      TimerVideo.Enabled := False;
      Exit;
    end;                        /////////
    // 7. Загрузка и воспроизведение видео
    MPMotivate.Clear;
    try
      MPMotivate.FileName := VideoFilePath;
      MPMotivate.Volume := TrackBVolV.Value;
      MPMotivate.Play;
      PanelMedia.Visible := True;
      MusicCheck := 'Video';
      //Memo1.Lines.Add('Видео запущено');
      //TimerVideo.Enabled := False;
    except
      //Memo1.Lines.Add('Видео не найдено');
      MusicCheck := 'VideoCheck';
    end;
  end;

end;


end.
