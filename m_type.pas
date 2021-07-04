(* Индексы для больших таблиц.
*(c) the Santig, licensed under zLib license *)
unit m_type;

{$mode objfpc}{$H+}

interface

uses
  Classes, Grids,
  me_my;

type
  {n-порядковый номер столбца в таблице, х-атрибут в xml, e-№ столбца в excel}
  TiPoleTab = record
    //поле xml
    x:      ShortString;
    //№ столбца
    n, exl: shortint;
  end;

  //Описание полей таблицы данных
  TiPa = record
    //id: integer;
    //вкл строка или нет             (0)
    vkl_An:    TiPoleTab;
    //передатчик                     (1)
    prd_An:    TiPoleTab;
    // № БС
    BS_No:     TiPoleTab;
    // Владелец  comment             (36)
    vlad_An:   TiPoleTab;
    //  Технология - NetStandard_name(2)
    tech_Se:   TiPoleTab;
    //Антенна (марка, словесное описание) (4)
    Model_An:  TiPoleTab;
    // Цвет                          (34)
    color_Se:  TiPoleTab;
    // Частота                       (5)
    f_Tr:      TiPoleTab;
    // Мощность передатчика          (8)
    P_Tr:      TiPoleTab;
    // кол-во передатчиков, TRX_count(9)
    kolPrd_An_Co: TiPoleTab;
    // Высота, м                     (10)
    h_An:      TiPoleTab;
    // азимут                        (11)
    azim_An:   TiPoleTab;
    // К. усиления                   (14)
    K_An:      TiPoleTab;
    // Наклон мех.                   (12)
    TiltM_An:  TiPoleTab;
    // Модуляция                     (19)
    modu_An:   TiPoleTab;
    //Поляризация                    (-)
    pol_An:    TiPoleTab;
    // Потери в АФУ, дБ/м !!!        (17)
    b_An:      TiPoleTab;
    // Длинна фидера, м              (15)
    bLen_An:   TiPoleTab;
    // Погонные потери в фидере, дБ,м(16)
    bPogon_An: TiPoleTab;
    // Потери в пассивных элементах, дБ (18)
    bPassiv_Co: TiPoleTab;
    // КНД для РРС                   (27)
    KND_An:    TiPoleTab;
    // диаметр аппертуры             (28)
    KND_D_An:  TiPoleTab;
    // угол раскрыва                 (29)
    KND_a_An:  TiPoleTab;
    // ПДУ параметр                  (6)
    PDUparam_Se: TiPoleTab;
    // ПДУ значение                  (7)
    PDUznach_Se: TiPoleTab;
    // координаты по широте B        (35)
    koordN_An: TiPoleTab;
    // координаты по долготе L       (35)
    koordE_An: TiPoleTab;
    // Место установки               (е13)
    adr_Si:    TiPoleTab;
    // Вид РЭС                       (е3)
    vidRES_No: TiPoleTab;
    // св-во                         (е10)
    svid_No:   TiPoleTab;
    // Дата выдачи св-ва             (е11)
    svidDO_No: TiPoleTab;
    // порядковый № строки из экселя (е14)
    nnExcel_No: TiPoleTab;
    // Ширина сектора запрета        (21)
    secZaprShir_An: TiPoleTab;
    // азимут сектора запрета        (22)
    secZaprAzim_An: TiPoleTab;
    // Мощность импульса, кВт        (23)
    impP_An:   TiPoleTab;
    // Частота посылки импульсов, Гц (24)
    impF_An:   TiPoleTab;
    // Длительность импульса, мкс    (25)
    impT_An:   TiPoleTab;
    // Тип апертуры                  (26)
    typeApert_An: TiPoleTab;
    // Сторона апертуры (гор), м     (30)
    storApertGor_An: TiPoleTab;
    // Сторона апертуры (верт), м    (31)
    storApertVert_An: TiPoleTab;
    // Угол раскрыва (гор)           (32)
    ugRasGor_An: TiPoleTab;
    // Угол раскрыва (верт)          (33)
    ugRasVer_An: TiPoleTab;
    // ДНА верт
    DNAvert_DNA: TiPoleTab;
    //ДНА гор
    DNAhorz_DNA: TiPoleTab;

    //Скрытые столбцы
    CombinerType_Co: TiPoleTab;
    CombinerType_An: TiPoleTab;
    calcType_An:     TiPoleTab;
    Model_id_An:     TiPoleTab;
  end;

  //Описание полей таблицы данных
  TiZd = record
    vkl:    TiPoleTab;
    cap:    TiPoleTab;
    typ:    TiPoleTab;
    h:      TiPoleTab;
    otrajBool: TiPoleTab;
    otrajKoef: TiPoleTab;
    LossBool: TiPoleTab;
    LossZnach: TiPoleTab;
    LossP2346: TiPoleTab;
    koment: TiPoleTab;
    tochki: TiPoleTab;
    color:  TiPoleTab;
    Loss_m: TiPoleTab;
  end;

procedure iPaIni(var grid: TStringGrid);
//Индекс зданий, дорог
procedure iZdIni;
//добавить столбец в grid
procedure iPaGridAddCol(var znach: TiPoleTab; exl: shortint; xml: ShortString;
  var grid: TStringGrid; const ParZnach: array of const);
//добавить индексы
procedure iIndAdd(var znach: TiPoleTab; n: integer; xml: string);
//title:ShortString; Width: byte);

var
  iPa: TiPa;
  iZd: TiZd;

implementation

procedure iPaIni(var grid: TStringGrid);
var
  cGrey, cOran: string;
begin
  cGrey := '$00D4D4D4';
  cOran := '$00ACD0FD';
  //grid.RowCount:=3;
  //вкл строка или нет             (0)
  iPaGridAddCol(iPa.vkl_An, -1, 'checked', grid, ['T', 'Вкл', 'w', 25]);
  grid.Columns.Items[iPa.vkl_An.n].ButtonStyle  := cbsCheckboxColumn;
  grid.Columns.Items[iPa.vkl_An.n].ValueChecked := '-1';
  //передатчик                     (1)
  iPaGridAddCol(iPa.prd_An, 19, 'name', grid, ['T', 'Передатчик', 'w', 120]);
  // № БС                          (e1)
  iPagridAddCol(iPa.BS_No, 7, '', grid, ['T', 'БС', 'w', 80]);
  // Владелец                      (36)
  iPagridAddCol(iPa.vlad_An, 5, 'comment', grid, ['T', 'Владелец', 'w', 123]);
  //  Технология                   (2)
  iPagridAddCol(iPa.tech_Se, -1, 'NetStandard_name',
    grid, ['T', 'Технология', 'w', 75]);
  //Антенна (марка, словесное описание) (4)
  iPagridAddCol(iPa.Model_An, -1, 'Model', grid, ['T', 'Антенна', 'w', 125]);
  // Цвет                          (34)
  iPagridAddCol(iPa.color_Se, -1, 'color', grid, ['T', 'Цвет', 'w', 40]);
  grid.Columns.Items[iPa.color_Se.n].ButtonStyle := cbsButton;
  // Частота                       (5)
  iPagridAddCol(iPa.f_Tr, 2, 'freqwork', grid, ['T', 'f', 'w', 45]);
  // Мощность передатчика          (8)
  iPagridAddCol(iPa.P_Tr, 10, 'power', grid, ['T', 'P', 'w', 40]);
  // кол-во передатчиков           (9)
  iPagridAddCol(iPa.kolPrd_An_Co, -1, 'TRX_count', grid,
    ['T', 'N кол-во передатчиков', 'w', 51]);
  // Высота, м                     (10)
  iPagridAddCol(iPa.h_An, 9, 'Height', grid, ['T', 'h', 'w', 32]);
  // азимут                        (11)
  iPagridAddCol(iPa.azim_An, 18, 'Azimuth', grid, ['T', 'Азимут', 'w', 45]);
  // К. усиления                   (14)
  iPagridAddCol(iPa.K_An, -1, 'Gain', grid, ['T', 'К ус.', 'w', 44]);
  // Наклон мех. {Slope}           (12)
  iPagridAddCol(iPa.TiltM_An, -1, 'Slope', grid,
    ['T', 'Наклон, град', 'w', 55]);
  // Модуляция                     (19)
  iPagridAddCol(iPa.modu_An, -1, 'Modulation', grid,
    ['T', 'Модуляция', 'w', 63]);
  // Поляризация
  iPagridAddCol(iPa.pol_An, -1, 'Polarization', grid,
    ['T', 'Поляризация', 'w', 38]);
  // Потери в АФУ, дБ/м !!!        (17)
  iPagridAddCol(iPa.b_An, -1, 'FeederLoss', grid, ['T', 'АФТ общее', 'w', 55]);
  // Длинна фидера, м              (15)
  iPagridAddCol(iPa.bLen_An, -1, 'AFU_Length_m', grid,
    ['T', 'АФТ длинна', 'w', 47]);
  // Погонные потери в фидере, дБ,м(16)
  iPagridAddCol(iPa.bPogon_An, -1, 'AFU_Loss_per_m', grid,
    ['T', 'АФТ погонное затухание', 'w', 52]);
  // Потери в пассивных элементах, дБ (18)
  iPagridAddCol(iPa.bPassiv_Co, -1, 'losses', grid,
    ['T', 'АФТ в пас.элм.', 'w', 61]);
  // КНД для РРС                   (27)
  iPagridAddCol(iPa.KND_An, -1, 'apert_KND', grid, ['T', 'КНД РРС', 'w', 54]);
  // диаметр аппертуры             (28)
  iPagridAddCol(iPa.KND_D_An, -1, 'apert_diam', grid,
    ['T', 'КНД. Диаметр', 'w', 63]);
  // угол раскрыва                 (29)
  iPagridAddCol(iPa.KND_a_An, -1, 'apert_angle', grid,
    ['T', 'КНД. Раскрыв', 'w', 81]);
  // ПДУ параметр                  (6)
  iPagridAddCol(iPa.PDUparam_Se, -1, 'Sanpin_PDU_type', grid,
    ['T', 'ПДУ-параметр', 'w', 82]);
  grid.Columns.Items[iPa.PDUparam_Se.n].ButtonStyle := cbsPickList;
  grid.Columns.Items[iPa.PDUparam_Se.n].PickList.Clear;
  grid.Columns.Items[iPa.PDUparam_Se.n].PickList.AddStrings(
    ['Е, В/м', 'ППЭ, мкВт/см2']);
  //grid.Cells[iPa.PDUparam_Se.n,1]:=grid.Columns.Items[iPa.PDUparam_Se.n].PickList.Strings[0];
  // ПДУ значение                  (7)
  iPagridAddCol(iPa.PDUznach_Se, -1, 'Sanpin_PDU', grid, ['T', 'ПДУ', 'w', 32]);
  // координаты размещения широта B(35)
  iPagridAddCol(iPa.koordN_An, -1, 'B', grid, ['T', 'Широта', 'w', 80]);
  // координаты размещения долгота L(35)
  iPagridAddCol(iPa.koordE_An, -1, 'L', grid, ['T', 'Долгота', 'w', 80]);
  // Место установки               (е13)
  iPagridAddCol(iPa.adr_Si, 6, '', grid, ['T', 'Место установки',
    'w', 170, 'c', cOran]);
  // Вид РЭС                       (е3)
  iPagridAddCol(iPa.vidRES_No, 1, '', grid, ['T', 'Вид РЭС', 'w', 205]);
  // св-во                         (е10)
  iPagridAddCol(iPa.svid_No, 15, '', grid, ['T', 'св-во', 'w', 96, 'c', cOran]);
  // Дата выдачи св-ва             (е11)
  iPagridAddCol(iPa.svidDO_No, 16, '', grid, ['T', 'Выдача св-ва',
    'w', 74, 'c', cOran]);

  // порядковый № строки из экселя (е14)
  iPagridAddCol(iPa.nnExcel_No, 0, '', grid,
    ['T', '№ стр из экселя', 'w', 91, 'c', cOran]);
  // Ширина сектора запрета        (21)
  iPagridAddCol(iPa.secZaprShir_An, -1, 'noscan_sector_hor', grid,
    ['T', 'Ширина сектора запрета', 'w', 44, 'c', cGrey]);
  // азимут сектора запрета        (22)
  iPagridAddCol(iPa.secZaprAzim_An, -1, 'noscan_sector_azm', grid,
    ['T', 'Азимут сектора запрета', 'w', 36, 'c', cGrey]);
  // Мощность импульса, кВт        (23)
  iPagridAddCol(iPa.impP_An, -1, 'impulse_pow_kvt', grid,
    ['T', 'Мощность импульса, кВт', 'w', 37, 'c', cGrey]);
  // Частота посылки импульсов, Гц (24)
  iPagridAddCol(iPa.impF_An, -1, 'impulse_freq_hz', grid,
    ['T', 'Частота посылки импульсов, Гц',
    'w', 47, 'c', cGrey]);
  // Длительность импульса, мкс    (25)
  iPagridAddCol(iPa.impT_An, -1, 'impulse_dur_mks', grid,
    ['T', 'Длительность импульса, мкс', 'w', 37, 'c', cGrey]);
  // Тип апертуры                  (26)
  iPagridAddCol(iPa.typeApert_An, -1, 'apert_type', grid,
    ['T', 'Тип апертуры', 'w', 76, 'c', cGrey]);
  grid.Columns.Items[iPa.typeApert_An.n].ButtonStyle := cbsPickList;
  grid.Columns.Items[iPa.typeApert_An.n].PickList.Clear;
  grid.Columns.Items[iPa.typeApert_An.n].PickList.AddStrings(
    ['Круглая', 'Прямоугольная', 'РПА']);

  // Сторона апертуры (гор), м     (30)
  iPagridAddCol(iPa.storApertGor_An, -1, 'apert_a', grid,
    ['T', 'Сторона апертуры (гор), м', 'w', 50, 'c', cGrey]);
  // Сторона апертуры (верт), м    (31)
  iPagridAddCol(iPa.storApertVert_An, -1, 'apert_b', grid,
    ['T', 'Сторона апертуры (верт), м', 'w', 52, 'c', cGrey]);
  // Угол раскрыва (гор)           (32)
  iPagridAddCol(iPa.ugRasGor_An, -1, 'apert_angle_hor', grid,
    ['T', 'Угол раскрыва (гор)', 'w', 33, 'c', cGrey]);
  // Угол раскрыва (верт)          (33)
  iPagridAddCol(iPa.ugRasVer_An, -1, 'apert_angle_ver', grid,
    ['T', 'Угол раскрыва (верт)', 'w', 58, 'c', cGrey]);
  // ДНА верт
  iPagridAddCol(iPa.DNAvert_DNA, -1, 'value', grid, ['T', 'ДНА верт', 'w', 33]);
  //ДНА гор
  iPagridAddCol(iPa.DNAhorz_DNA, -1, 'value', grid, ['T', 'ДНА гор', 'w', 33]);

  //Скрытые столбцы

  iPagridAddCol(iPa.CombinerType_Co, -1, 'Combiner_type', grid,
    ['T', 'Combiner_type', 'w', 0]);
  iPagridAddCol(iPa.CombinerType_An, -1, 'Combiner_type', grid,
    ['T', 'Combiner_type', 'w', 0]);
  iPagridAddCol(iPa.calcType_An, -1, 'calc_type', grid, ['T', 'calc_type', 'w', 0]);
  iPagridAddCol(iPa.Model_id_An, -1, 'Model_id', grid, ['T', 'Model_id', 'w', 0]);
end;

//Индекс зданий, дорог
procedure iZdIni;
begin
  iIndAdd(iZd.vkl, 0 + 1, 'checked');
  iIndAdd(iZd.cap, 1 + 1, 'caption');
  iIndAdd(iZd.typ, 2 + 1, 'ClutterType');
  iIndAdd(iZd.h, 3 + 1, 'ClutterHeight');
  iIndAdd(iZd.otrajBool, 4 + 1, 'is_Reflection');
  iIndAdd(iZd.otrajKoef, 5 + 1, 'Coef_Reflection');
  iIndAdd(iZd.color, 6 + 1, 'Color');
  iIndAdd(iZd.LossBool, 7 + 1, 'is_Clutter_Loss');
  iIndAdd(iZd.LossZnach, 8 + 1, 'Clutter_Loss_dB');
  iIndAdd(iZd.Loss_m, 9 + 1, 'Loss_per_m');
  iIndAdd(iZd.LossP2346, 10 + 1, 'is_P2346');
  iIndAdd(iZd.tochki, 11 + 1, 'point');
  iIndAdd(iZd.koment, 12 + 1, 'comment');
end;

procedure iPaGridAddCol(var znach: TiPoleTab; exl: shortint; xml: ShortString;
  var grid: TStringGrid; const ParZnach: array of const);
begin
  iIndAdd(znach, grid.ColCount, xml);
  znach.exl := exl;
  //if Assigned(ParZnach) then
  GridAddCol(grid, ParZnach);
end;

//добавить индексы
procedure iIndAdd(var znach: TiPoleTab; n: integer; xml: string);
begin
  znach.n := n;
  znach.x := xml;
end;

end.
