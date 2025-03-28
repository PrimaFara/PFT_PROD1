unit LaporanDeptKemitraan;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  DateUtils, Mask, wwdbedit, Wwdbspin, DB, Wwdatsrc, OracleData, Dialogs,
  Oracle, Buttons, wwSpeedButton, wwDBNavigator, wwclearpanel, Grids,
  Wwdbigrd, Wwdbgrid, StdCtrls, wwdbdatetimepicker, ComCtrls, ExtCtrls,
  ppViewr, ppDB, ppDBPipe, ppComm, ppRelatv, ppProd, ppClass, ppReport,
  ppVar, ppCtrls, ppBands, ppPrnabl, ppCache, ppEndUsr, ppModule,
  daDataModule, DBCtrls, wwdblook, Wwdbdlg, Wwdotdot, Wwdbcomb, ppBarCod,
  wwcheckbox, ppStrtch, ppMemo, raCodMod, ppParameter;

type
  TLaporanDeptKemitraanFrm = class(TForm)
    QBrowse: TOracleDataSet;
    dsQBrowse: TwwDataSource;
    ppReportBrowse: TppReport;
    ppDBQBrowse: TppDBPipeline;
    QMaster: TOracleDataSet;
    dsQMaster: TwwDataSource;
    QTransaksi: TOracleDataSet;
    QTransaksiNAMA_TRANSAKSI: TStringField;
    QTransaksiKD_TRANSAKSI: TStringField;
    QTransaksiPREFIX: TStringField;
    QTransaksiPLINE: TStringField;
    QTransaksiPHEADER: TStringField;
    QTransaksiDISTRIBUSI: TStringField;
    QTransaksiTTD1: TStringField;
    QTransaksiTTD2: TStringField;
    QTransaksiTTD3: TStringField;
    QTransaksiTTD4: TStringField;
    QTransaksiDIV1: TStringField;
    QTransaksiDIV2: TStringField;
    QTransaksiDIV3: TStringField;
    QTransaksiDIV4: TStringField;
    QTransaksiJAB1: TStringField;
    QTransaksiJAB2: TStringField;
    QTransaksiJAB3: TStringField;
    QTransaksiJAB4: TStringField;
    dsQTransaksi: TwwDataSource;
    QProcUpdateMutasi: TOracleQuery;
    ppTitleBand1: TppTitleBand;
    ppNamaLaporan: TppLabel;
    ppLabel9: TppLabel;
    ppPeriode: TppLabel;
    ppDBText12: TppDBText;
    ppDBText13: TppDBText;
    ppDBText14: TppDBText;
    ppHeaderBand1: TppHeaderBand;
    ppLabel6: TppLabel;
    ppDetailBand1: TppDetailBand;
    ppNo: TppVariable;
    ppDBText1: TppDBText;
    ppFooterBand1: TppFooterBand;
    ppSystemVariable1: TppSystemVariable;
    ppSummaryBand1: TppSummaryBand;
    ppDBText19: TppDBText;
    ppDBText42: TppDBText;
    ppDBText43: TppDBText;
    ppDBText44: TppDBText;
    ppDBText45: TppDBText;
    ppDBText46: TppDBText;
    ppDBText47: TppDBText;
    ppDBText48: TppDBText;
    ppDBText49: TppDBText;
    ppDBText5: TppDBText;
    QTransaksiDOC_ISO: TStringField;
    QBrowseNAMA_MITRA: TStringField;
    QBrowseBENANG: TStringField;
    QBrowseLUSI_AWAL: TFloatField;
    QBrowsePAKAN_AWAL: TFloatField;
    QBrowseLUSI_KIRIM: TFloatField;
    QBrowsePAKAN_KIRIM: TFloatField;
    QBrowseLUSI_TERIMA_PRODUKSI: TFloatField;
    QBrowsePAKAN_TERIMA_PRODUKSI: TFloatField;
    QBrowseLUSI_TERIMA_RETUR: TFloatField;
    QBrowsePAKAN_TERIMA_RETUR: TFloatField;
    QBrowseLUSI_TERIMA_AFVAL: TFloatField;
    QBrowsePAKAN_TERIMA_AFVAL: TFloatField;
    QBrowseLUSI_KOREKSI: TFloatField;
    QBrowsePAKAN_KOREKSI: TFloatField;
    QBrowseLUSI_AKHIR: TFloatField;
    QBrowsePAKAN_AKHIR: TFloatField;
    dsQBrowse1: TwwDataSource;
    QBrowse1: TOracleDataSet;
    QBrowse1JENIS: TStringField;
    QBrowse1AWAL: TFloatField;
    QBrowse1PEMASUKAN: TFloatField;
    QBrowse1PENGELUARAN: TFloatField;
    QBrowse1AKHIR: TFloatField;
    ppLine1: TppLine;
    ppLine3: TppLine;
    ppUserCetak: TppLabel;
    ppLabel2: TppLabel;
    ppLabel3: TppLabel;
    ppLine4: TppLine;
    ppLine5: TppLine;
    ppLine6: TppLine;
    ppLine7: TppLine;
    ppLabel4: TppLabel;
    ppLabel5: TppLabel;
    ppLabel7: TppLabel;
    ppLabel8: TppLabel;
    ppLabel10: TppLabel;
    ppLine8: TppLine;
    ppLine9: TppLine;
    ppLabel11: TppLabel;
    ppLabel12: TppLabel;
    ppLine10: TppLine;
    ppLine11: TppLine;
    ppLabel14: TppLabel;
    ppLabel17: TppLabel;
    ppLine12: TppLine;
    ppLabel18: TppLabel;
    ppLabel19: TppLabel;
    ppLabel24: TppLabel;
    ppLabel27: TppLabel;
    ppLabel30: TppLabel;
    ppLabel31: TppLabel;
    ppLabel32: TppLabel;
    ppLabel33: TppLabel;
    ppLabel34: TppLabel;
    ppLabel37: TppLabel;
    ppLabel38: TppLabel;
    ppLabel39: TppLabel;
    ppLabel40: TppLabel;
    ppLine13: TppLine;
    ppLine14: TppLine;
    ppLine15: TppLine;
    ppLine16: TppLine;
    ppLine17: TppLine;
    ppLine18: TppLine;
    ppLine19: TppLine;
    ppLine20: TppLine;
    ppLine21: TppLine;
    ppLine22: TppLine;
    ppLine23: TppLine;
    ppLine24: TppLine;
    ppLine25: TppLine;
    ppDBText2: TppDBText;
    ppLine26: TppLine;
    ppDBText3: TppDBText;
    ppDBText4: TppDBText;
    ppDBText6: TppDBText;
    ppDBText7: TppDBText;
    ppDBText8: TppDBText;
    ppDBText9: TppDBText;
    ppDBText10: TppDBText;
    ppDBText17: TppDBText;
    ppDBText22: TppDBText;
    ppDBText23: TppDBText;
    ppDBText24: TppDBText;
    ppDBText27: TppDBText;
    ppDBText28: TppDBText;
    ppLine27: TppLine;
    ppLine28: TppLine;
    ppLine29: TppLine;
    ppLine30: TppLine;
    ppLine31: TppLine;
    ppLine32: TppLine;
    ppLine33: TppLine;
    ppLine34: TppLine;
    ppLine35: TppLine;
    ppLine36: TppLine;
    ppLine37: TppLine;
    ppLine38: TppLine;
    ppLine39: TppLine;
    ppLine40: TppLine;
    ppLine41: TppLine;
    ppLine42: TppLine;
    ppLine43: TppLine;
    ppLine44: TppLine;
    ppLine45: TppLine;
    ppDBCalc12: TppDBCalc;
    ppLine46: TppLine;
    ppDBCalc1: TppDBCalc;
    ppDBCalc3: TppDBCalc;
    ppDBCalc4: TppDBCalc;
    ppDBCalc5: TppDBCalc;
    ppDBCalc6: TppDBCalc;
    ppDBCalc7: TppDBCalc;
    ppDBCalc8: TppDBCalc;
    ppDBCalc9: TppDBCalc;
    ppDBCalc10: TppDBCalc;
    ppDBCalc11: TppDBCalc;
    ppDBCalc13: TppDBCalc;
    ppDBCalc14: TppDBCalc;
    ppDBCalc15: TppDBCalc;
    ppLine47: TppLine;
    ppLine48: TppLine;
    ppLine49: TppLine;
    ppLabel41: TppLabel;
    ppLine50: TppLine;
    ppLine51: TppLine;
    ppLine52: TppLine;
    ppLine53: TppLine;
    ppLine54: TppLine;
    ppLine55: TppLine;
    ppLine56: TppLine;
    ppLine57: TppLine;
    ppLine58: TppLine;
    ppLine59: TppLine;
    ppLine60: TppLine;
    ppLine61: TppLine;
    ppLine62: TppLine;
    ppLine63: TppLine;
    ppDBText29: TppDBText;
    ppDBText30: TppDBText;
    ppDBText31: TppDBText;
    ppLabel42: TppLabel;
    PanelMain: TPanel;
    Label1: TLabel;
    PanelHeader: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Label10: TLabel;
    PanelFooter1: TPanel;
    BtnClose1: TBitBtn;
    Panel3: TPanel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    vtglAwal0: TwwDBDateTimePicker;
    vtglAkhir0: TwwDBDateTimePicker;
    BitBtn2: TBitBtn;
    wwDBSpinEdit2: TwwDBSpinEdit;
    Panel4: TPanel;
    wwDBGrid1: TwwDBGrid;
    TabSheet2: TTabSheet;
    LabelBanner: TLabel;
    PanelFilter: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    BtnFind: TSpeedButton;
    BtnOk2: TSpeedButton;
    vTglAwal: TwwDBDateTimePicker;
    vTglAkhir: TwwDBDateTimePicker;
    BtnOk: TBitBtn;
    wwDBSpinLine2: TwwDBSpinEdit;
    cbAdaTransaksi: TCheckBox;
    PanelBrowse: TPanel;
    wwDBGrid2: TwwDBGrid;
    PanelFooter2: TPanel;
    wwDBNavigator1: TwwDBNavigator;
    wwDBNavigator1PriorPage: TwwNavButton;
    wwDBNavigator1NextPage: TwwNavButton;
    wwDBNavigator1SaveBookmark: TwwNavButton;
    wwDBNavigator1RestoreBookmark: TwwNavButton;
    BtnClose2: TBitBtn;
    BtnExport: TBitBtn;
    BtnPrintBrowse: TBitBtn;
    BtnDesign2: TBitBtn;
    TabSheet3: TTabSheet;
    Panel1: TPanel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    vTglAwal2: TwwDBDateTimePicker;
    vTglAkhir2: TwwDBDateTimePicker;
    BitBtn1: TBitBtn;
    wwDBSpinEdit1: TwwDBSpinEdit;
    CheckBox1: TCheckBox;
    Panel2: TPanel;
    wwDBGrid3: TwwDBGrid;
    QMasterMITRA: TStringField;
    QMasterJENIS: TStringField;
    QMasterKP: TStringField;
    QMasterARAH: TStringField;
    QMasterKONSTRUKSI: TStringField;
    QMasterQTY_KG: TFloatField;
    QMasterQTY_PTG: TFloatField;
    QMasterQTY_SPRING: TFloatField;
    QMasterQTY_BEAM: TFloatField;
    QMasterQTY_BS: TFloatField;
    QProcRekapBulanan: TOracleQuery;
    ppDBQMaster: TppDBPipeline;
    ppDBQMasterppField1: TppField;
    ppDBQMasterppField2: TppField;
    ppDBQMasterppField3: TppField;
    ppDBQMasterppField4: TppField;
    ppDBQMasterppField5: TppField;
    ppDBQMasterppField6: TppField;
    ppDBQMasterppField7: TppField;
    ppDBQMasterppField8: TppField;
    ppDBQMasterppField9: TppField;
    ppDBQMasterppField10: TppField;
    ppDBQMasterppField11: TppField;
    ppDBQMasterppField12: TppField;
    ppDBQMasterppField13: TppField;
    ppDBQMasterppField14: TppField;
    ppDBQMasterppField15: TppField;
    ppDBQMasterppField16: TppField;
    ppDBQMasterppField17: TppField;
    ppDBQMasterppField18: TppField;
    ppDBQMasterppField19: TppField;
    ppReportInput: TppReport;
    ppTitleBand2: TppTitleBand;
    ppDBText15: TppDBText;
    ppDBText50: TppDBText;
    ppVariable2: TppVariable;
    ppDBText26: TppDBText;
    ppLabel25: TppLabel;
    ppDBText11: TppDBText;
    ppDBText16: TppDBText;
    ppDBText18: TppDBText;
    ppLabel1: TppLabel;
    ppLabel13: TppLabel;
    ppLabel15: TppLabel;
    ppDBText20: TppDBText;
    ppLabel16: TppLabel;
    ppDBText25: TppDBText;
    ppLabel20: TppLabel;
    ppLabel35: TppLabel;
    ppLabel21: TppLabel;
    ppLabel22: TppLabel;
    ppLabel23: TppLabel;
    ppDBText51: TppDBText;
    ppHeaderBand2: TppHeaderBand;
    ppLabel26: TppLabel;
    ppLabel28: TppLabel;
    ppLabel29: TppLabel;
    ppLabel36: TppLabel;
    ppLabel43: TppLabel;
    ppLabel44: TppLabel;
    ppLabel45: TppLabel;
    ppLabel46: TppLabel;
    ppLabel47: TppLabel;
    ppLabel48: TppLabel;
    ppLabel49: TppLabel;
    ppDetailBand2: TppDetailBand;
    ppDBText21: TppDBText;
    ppDBText32: TppDBText;
    ppDBMemo1: TppDBMemo;
    ppDBText33: TppDBText;
    ppDBText34: TppDBText;
    ppDBText35: TppDBText;
    ppDBText36: TppDBText;
    ppDBText37: TppDBText;
    ppDBText38: TppDBText;
    ppDBText39: TppDBText;
    ppVariable1: TppVariable;
    ppLabel50: TppLabel;
    ppFooterBand2: TppFooterBand;
    ppUserCetak2: TppLabel;
    ppDBText41: TppDBText;
    ppSummaryBand2: TppSummaryBand;
    ppLabel51: TppLabel;
    ppDBMemo2: TppDBMemo;
    ppDBText40: TppDBText;
    ppDBText52: TppDBText;
    ppDBText53: TppDBText;
    ppDBText54: TppDBText;
    ppDBText55: TppDBText;
    ppDBText56: TppDBText;
    ppDBText57: TppDBText;
    ppDBText58: TppDBText;
    ppDBText59: TppDBText;
    ppLine2: TppLine;
    ppLine64: TppLine;
    ppDBCalc2: TppDBCalc;
    ppDBCalc16: TppDBCalc;
    ppPageStyle1: TppPageStyle;
    raCodeModule1: TraCodeModule;
    ppParameterList2: TppParameterList;
    ppDBQTransaksi: TppDBPipeline;
    ppDBQDetail: TppDBPipeline;
    ppDBQDetailppMasterFieldLink1: TppMasterFieldLink;
    ppReport1: TppReport;
    ppTitleBand3: TppTitleBand;
    ppLabel52: TppLabel;
    ppLabel53: TppLabel;
    ppLabel54: TppLabel;
    ppDBText60: TppDBText;
    ppDBText61: TppDBText;
    ppDBText62: TppDBText;
    ppLabel55: TppLabel;
    ppHeaderBand3: TppHeaderBand;
    ppLabel56: TppLabel;
    ppLabel57: TppLabel;
    ppLabel58: TppLabel;
    ppLabel59: TppLabel;
    ppLabel60: TppLabel;
    ppLabel61: TppLabel;
    ppLabel62: TppLabel;
    ppLabel63: TppLabel;
    ppLabel64: TppLabel;
    ppLabel65: TppLabel;
    ppLabel66: TppLabel;
    ppLabel67: TppLabel;
    ppLabel68: TppLabel;
    ppDetailBand3: TppDetailBand;
    ppDBText63: TppDBText;
    ppDBText64: TppDBText;
    ppDBText65: TppDBText;
    ppDBMemo3: TppDBMemo;
    ppVariable3: TppVariable;
    ppDBText66: TppDBText;
    ppDBText67: TppDBText;
    ppDBText68: TppDBText;
    ppDBText69: TppDBText;
    ppDBText70: TppDBText;
    ppDBText71: TppDBText;
    ppDBText72: TppDBText;
    ppDBText73: TppDBText;
    ppFooterBand3: TppFooterBand;
    ppSystemVariable2: TppSystemVariable;
    ppSummaryBand3: TppSummaryBand;
    ppDBText74: TppDBText;
    ppDBText75: TppDBText;
    ppDBText76: TppDBText;
    ppDBText77: TppDBText;
    ppDBText78: TppDBText;
    ppDBText79: TppDBText;
    ppDBText80: TppDBText;
    ppDBText81: TppDBText;
    ppDBText82: TppDBText;
    ppParameterList3: TppParameterList;
    Panel5: TPanel;
    QMasterQTY_PERSEN: TFloatField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BtnExportClick(Sender: TObject);
    procedure BtnOkClick(Sender: TObject);
    procedure wwDBGrid2TitleButtonClick(Sender: TObject;
      AFieldName: String);
    procedure QBrowseAfterScroll(DataSet: TDataSet);
    procedure BtnClose1Click(Sender: TObject);
    procedure BtnClose2Click(Sender: TObject);
    procedure vTglAwalChange(Sender: TObject);
    procedure wwDBSpinLine2Change(Sender: TObject);
    procedure BtnFindClick(Sender: TObject);
    procedure BtnOk2Click(Sender: TObject);
    procedure ppHeaderBand1BeforePrint(Sender: TObject);
    procedure BtnPrintBrowseClick(Sender: TObject);
    procedure ppTitleBand1BeforePrint(Sender: TObject);
    procedure TabSheet1Show(Sender: TObject);
    procedure ppDetailBand2BeforePrint(Sender: TObject);
    procedure wwDBGrid1Enter(Sender: TObject);
    procedure wwDBGrid2DblClick(Sender: TObject);
    procedure Label5Click(Sender: TObject);
    procedure LookItemEnter(Sender: TObject);
    procedure QTransaksiBeforeOpen(DataSet: TDataSet);
    procedure QMasterAfterPost(DataSet: TDataSet);
    procedure QMasterBeforeInsert(DataSet: TDataSet);
    procedure QProcUpdateMutasiBeforeQuery(Sender: TOracleQuery);
    procedure FormShow(Sender: TObject);
    procedure ppDetailBand1BeforePrint(Sender: TObject);
    procedure QBrowseCalcFields(DataSet: TDataSet);
    procedure wwDBGrid2UpdateFooter(Sender: TObject);
    procedure wwDBGrid1CalcCellColors(Sender: TObject; Field: TField;
      State: TGridDrawState; Highlight: Boolean; AFont: TFont;
      ABrush: TBrush);
    procedure QBrowseFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    procedure cbAdaTransaksiClick(Sender: TObject);
    procedure vTglAwal2Change(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure ppSummaryBand1BeforePrint(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure vtglAwal0Change(Sender: TObject);
    procedure Label9Click(Sender: TObject);
    procedure wwDBSpinEdit2Change(Sender: TObject);
    procedure QMasterCalcFields(DataSet: TDataSet);
  //  procedure ppNo2Print(Sender: TObject);
  private
    { Private declarations }
    vshift, vgrup, vorder, vfilter, SelectedFont, vopr, vkode, vjns_brg, vjns_lokasi, vkd_benang : String;
  public
    { Public declarations }

  end;

var
  LaporanDeptKemitraanFrm: TLaporanDeptKemitraanFrm;

procedure ShowForm(pNamaMenu:String; pkode : String; pjudul : string; pjns_brg : String; pjns_lokasi : String);

implementation

uses DM, Pembelian, HasilSoftCones, ComObj;

{$R *.dfm}

procedure ShowForm(pNamaMenu:String; pkode : String; pjudul : string; pjns_brg : String; pjns_lokasi : String);
var
  mychar : string[125];
Begin
// Hak Menu
  DMFrm.cHakInput:=DMFrm.QMenuUser.Lookup('NAMA_COMPONENT',pNamaMenu,'HAK_INPUT')='1';
  DMFrm.cBtnDesign:=DMFrm.QMenuUser.Lookup('NAMA_COMPONENT',pNamaMenu,'HAK_DESIGN')='1';
  DMFrm.cBtnExport:=DMFrm.QMenuUser.Lookup('NAMA_COMPONENT',pNamaMenu,'HAK_EXPORT')='1';

//  if BPBFrm=Nil then
  begin
    LaporanDeptKemitraanFrm:=TLaporanDeptKemitraanFrm.Create(Application);
    LaporanDeptKemitraanFrm.PageControl1.ActivePageIndex:=0;

    mychar:=pjudul;
    Delete(mychar,Pos('&',mychar),1);
    pjudul:=mychar;
    LaporanDeptKemitraanFrm.Caption:=UpperCase(pjudul);
    LaporanDeptKemitraanFrm.vkode:=pkode;
    LaporanDeptKemitraanFrm.vjns_lokasi:=pjns_lokasi;
    LaporanDeptKemitraanFrm.vjns_brg:=pjns_brg;
    LaporanDeptKemitraanFrm.QTransaksi.Open;

    LaporanDeptKemitraanFrm.PanelHeader.Caption:=LaporanDeptKemitraanFrm.QTransaksiNAMA_TRANSAKSI.AsString;

    LaporanDeptKemitraanFrm.wwDBGrid2.IniAttributes.FileName:=DMFrm.sAppPath+Application.Title+'2.ini';
    LaporanDeptKemitraanFrm.wwDBGrid2.IniAttributes.SectionName:=LaporanDeptKemitraanFrm.Caption+'2';
    LaporanDeptKemitraanFrm.wwDBGrid2.IniAttributes.Enabled:=True;
    LaporanDeptKemitraanFrm.wwDBGrid2.LoadFromIniFile;
    DMFrm.ProcReadIni(Application.Title,LaporanDeptKemitraanFrm.Caption+'2',LaporanDeptKemitraanFrm.wwDBGrid2);
    LaporanDeptKemitraanFrm.wwDBSpinLine2.Value:=LaporanDeptKemitraanFrm.wwDBGrid2.RowHeightPercent;

  end;

  LaporanDeptKemitraanFrm.Show;
end;

procedure TLaporanDeptKemitraanFrm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   DMFrm.ProcWtiteIni(Application.Title,Caption+'2',wwDBGrid2);
   Action:=caFree;
   LaporanDeptKemitraanFrm:=Nil;
end;

procedure TLaporanDeptKemitraanFrm.FormCreate(Sender: TObject);
begin
//barcode
     DMFrm.FontToUse := TFont.Create;
     DMFrm.selected := 'UCC 128';
     SelectedFont := 'CIA Code 128 Medium';
     StrPCopy(DMFrm.TempSelected, DMFrm.Selected);
     DMFrm.BType := 'C128';
     DMFrm.FontToUse.Size := 12;
     DMFrm.FontToUse.Name := SelectedFont;
//     QRBarcode11.Font := LoginFrm.FontToUse;
     DMFrm.BType := DMFrm.BType + '-';
     DMFrm.BType := DMFrm.BType + DMFrm.Format;
     DMFrm.BType := DMFrm.BType + '.BH';
// end barcode

//  BtnAmbilData.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\COPY.Bmp');

  BtnClose1.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\CLOSE.Bmp');
  BtnOk.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\CHECK.Bmp');
  BtnOk2.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\CHECK.Bmp');
  BtnFind.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\FIND.Bmp');
  BtnDesign2.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\DESIGN.Bmp');
  BtnPrintBrowse.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\PRINT.Bmp');
  BtnExport.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\EXPORT.Bmp');
  BtnClose2.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\CLOSE.Bmp');
  vTglAwal.Date:=Trunc(Date);
  vTglAwal0.Date:=Trunc(Date);
  vTglAwal2.Date:=Trunc(Date);
//Otoritas Button
  BtnExport.Visible:=DMFrm.cBtnExport;
  BtnDesign2.Visible:=DMFrm.cBtnDesign;
end;

procedure TLaporanDeptKemitraanFrm.BtnExportClick(Sender: TObject);
var
  ExcelApp: Variant;
  Workbook: Variant;
  Worksheet: Variant;
  i, j: Integer;
  Field: TField;
  FormatSettings: TFormatSettings;
begin
  ShowMessage('Tunggu Hingga Proses Selesai. Klik Ok untuk memulai');
  
  // Inisialisasi TFormatSettings untuk mengontrol pemisah desimal
  GetLocaleFormatSettings(GetThreadLocale, FormatSettings);
  FormatSettings.DecimalSeparator := '.';

  // Menampilkan dialog Save As
  DMFrm.SaveDialog1.Filter := 'Excel Files|*.xlsx';
  DMFrm.SaveDialog1.DefaultExt := 'xlsx';

  if DMFrm.SaveDialog1.Execute then
  begin
    // Membuat instance dari Excel
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Visible := False; // Atur ke True jika Anda ingin melihat proses di Excel

    // Menambahkan workbook baru
    Workbook := ExcelApp.Workbooks.Add;
    Worksheet := Workbook.Worksheets[1];

    // Menulis header ke worksheet
    for i := 0 to wwDBGrid2.DataSource.DataSet.FieldCount - 1 do
    begin
      Worksheet.Cells[1, i + 1].Value := wwDBGrid2.DataSource.DataSet.Fields[i].DisplayName;
    end;

    // Menulis data ke worksheet
    wwDBGrid2.DataSource.DataSet.First;
    j := 2; // Mulai dari baris kedua untuk data, karena baris pertama untuk header
    while not wwDBGrid2.DataSource.DataSet.Eof do
    begin
      for i := 0 to wwDBGrid2.DataSource.DataSet.FieldCount - 1 do
      begin
        Field := wwDBGrid2.DataSource.DataSet.Fields[i];

        // Memeriksa tipe data field
        if Field.DataType in [ftDate, ftDateTime] then
        begin
          Worksheet.Cells[j, i + 1].Value := Field.AsDateTime;
          Worksheet.Cells[j, i + 1].NumberFormat := 'dd/mm/yyyy';
        end
        else if Field.DataType in [ftFloat, ftCurrency, ftBCD] then
        begin
          // Menyimpan nilai float dalam string dengan FormatSettings yang benar
          Worksheet.Cells[j, i + 1].Value := FloatToStr(Field.AsFloat, FormatSettings);
          Worksheet.Cells[j, i + 1].NumberFormat := '0.00'; // Atur format angka dengan dua desimal
        end
        else
        begin
          Worksheet.Cells[j, i + 1].Value := Field.AsString;
        end;

        // Mengatur alignment sel untuk angka menjadi rata kanan
        if Field.DataType in [ftInteger, ftFloat, ftCurrency, ftBCD] then
        begin
          Worksheet.Cells[j, i + 1].HorizontalAlignment := 3; // xlRight
        end;
      end;
      Inc(j);
      wwDBGrid2.DataSource.DataSet.Next;
    end;

    // Menyimpan workbook ke file yang dipilih pengguna
    Workbook.SaveAs(DMFrm.SaveDialog1.FileName);

    // Menutup workbook dan mengeluarkan Excel dari memori
    Workbook.Close(False);
    ExcelApp.Quit;

    // Melepaskan objek COM
    Worksheet := Unassigned;
    Workbook := Unassigned;
    ExcelApp := Unassigned;

    ShowMessage('Data berhasil diekspor ke ' + DMFrm.SaveDialog1.FileName);
  end
  else
  begin
    ShowMessage('Proses penyimpanan dibatalkan.');
  end;
end;

procedure TLaporanDeptKemitraanFrm.BtnOkClick(Sender: TObject);
var vawal_l, vawal_p, vkirim_l, vkirim_p, vterima_prod_l, vterima_prod_p, vterima_retur_l, vterima_retur_p, vterima_afval_l, vterima_afval_p, vkoreksi_l, vkoreksi_p, vakhir_l, vakhir_p : Real;
begin
  if vTglAwal.Date>vTglAkhir.DateTime then
    ShowMessage('Tgl. Akhir harus lebih besar dari Tgl. Awal !')
    else
    begin

      QProcUpdateMutasi.Close;
      QProcUpdateMutasi.SetVariable('pawal', vTglAwal.Date);
      QProcUpdateMutasi.SetVariable('pakhir', vTglAkhir.Date);
      QProcUpdateMutasi.Execute;

      if QBrowse.QBEMode then QBrowse.QBEMode:=False;
      QBrowse.DisableControls;
      QBrowse.Close;
      QBrowse.Open;
      vawal_l:=0;
      vkirim_l:=0;
      vterima_prod_l:=0;
      vterima_retur_l:=0;
      vterima_afval_l:=0;
      vkoreksi_l:=0;
      vakhir_l:=0;
      vawal_p:=0;
      vkirim_p:=0;
      vterima_prod_p:=0;
      vterima_retur_p:=0;
      vterima_afval_p:=0;
      vkoreksi_p:=0;
      vakhir_p:=0;
      while not QBrowse.Eof do
      begin
        vawal_l:=vawal_l+QBrowseLUSI_AWAL.AsFloat;
        vkirim_l:=vkirim_l+QBrowseLUSI_KIRIM.AsFloat;
        vterima_prod_l:=vterima_prod_l+QBrowseLUSI_TERIMA_PRODUKSI.AsFloat;
        vterima_retur_l:=vterima_retur_l+QBrowseLUSI_TERIMA_RETUR.AsFloat;
        vterima_afval_l:=vterima_afval_l+QBrowseLUSI_TERIMA_AFVAL.AsFloat;
        vkoreksi_l:=vkoreksi_l+QBrowseLUSI_KOREKSI.AsFloat;
        vakhir_l:=vakhir_l+QBrowseLUSI_AKHIR.AsFloat;

        vawal_p:=vawal_p+QBrowsePAKAN_AWAL.AsFloat;
        vkirim_p:=vkirim_p+QBrowsePAKAN_KIRIM.AsFloat;
        vterima_prod_p:=vterima_prod_p+QBrowsePAKAN_TERIMA_PRODUKSI.AsFloat;
        vterima_retur_p:=vterima_retur_p+QBrowsePAKAN_TERIMA_RETUR.AsFloat;
        vterima_afval_p:=vterima_afval_p+QBrowsePAKAN_TERIMA_AFVAL.AsFloat;
        vkoreksi_p:=vkoreksi_p+QBrowsePAKAN_KOREKSI.AsFloat;
        vakhir_p:=vakhir_p+QBrowsePAKAN_AKHIR.AsFloat;

        QBrowse.Next;
      end;
      QBrowse.EnableControls;
      wwDBGrid2.ColumnByName('LUSI_AWAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vawal_l);
      wwDBGrid2.ColumnByName('LUSI_KIRIM').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkirim_l);
      wwDBGrid2.ColumnByName('LUSI_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_prod_l);
      wwDBGrid2.ColumnByName('LUSI_TERIMA_RETUR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_retur_l);
      wwDBGrid2.ColumnByName('LUSI_TERIMA_AFVAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_afval_l);
      wwDBGrid2.ColumnByName('LUSI_KOREKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkoreksi_l);
      wwDBGrid2.ColumnByName('LUSI_AKHIR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vakhir_l);

      wwDBGrid2.ColumnByName('PAKAN_AWAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vawal_p);
      wwDBGrid2.ColumnByName('PAKAN_KIRIM').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkirim_p);
      wwDBGrid2.ColumnByName('PAKAN_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_prod_p);
      wwDBGrid2.ColumnByName('PAKAN_TERIMA_RETUR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_retur_p);
      wwDBGrid2.ColumnByName('PAKAN_TERIMA_AFVAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_afval_p);
      wwDBGrid2.ColumnByName('PAKAN_KOREKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkoreksi_p);
      wwDBGrid2.ColumnByName('PAKAN_AKHIR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vakhir_p);

      LabelBanner.Caption:='Data : '+FormatFloat('#,#',QBrowse.RecordCount)+' Records';
    end;

end;

procedure TLaporanDeptKemitraanFrm.wwDBGrid2TitleButtonClick(Sender: TObject;
  AFieldName: String);
begin
  if QBrowse.FieldByName(AFieldName).FieldKind=fkData then
  begin
    vorder:='order by '+AFieldName;
    BtnOkClick(Nil);
  end
  else
    ShowMessage('Maaf, tidak bisa Urut menurut kolom '+AFieldName+' !');
end;

procedure TLaporanDeptKemitraanFrm.QBrowseAfterScroll(DataSet: TDataSet);
begin
  LabelBanner.Caption:='Record ke '+IntToStr(QBrowse.RecNo)+' dari '+FormatFloat('#,#',QBrowse.RecordCount)+' Records';
end;

procedure TLaporanDeptKemitraanFrm.BtnClose1Click(Sender: TObject);
begin
  Close;
end;

procedure TLaporanDeptKemitraanFrm.BtnClose2Click(Sender: TObject);
begin
  Close;
end;

procedure TLaporanDeptKemitraanFrm.vTglAwalChange(Sender: TObject);
begin
  vTglAkhir.DateTime:=EndOfTheMonth(vTglAwal.Date);
end;

procedure TLaporanDeptKemitraanFrm.wwDBSpinLine2Change(Sender: TObject);
begin
    wwDBGrid2.RowHeightPercent:=Round(wwDBSpinLine2.Value);
end;

procedure TLaporanDeptKemitraanFrm.BtnFindClick(Sender: TObject);
begin
  if not QBrowse.QBEMode then
  begin
    wwDBGrid2.Options:=wwDBGrid2.Options-[dgRowSelect,dgAlwaysShowSelection];
    QBrowse.QBEMode:=TRUE;
  end
  else
    QBrowse.ClearQBE;
end;

procedure TLaporanDeptKemitraanFrm.BtnOk2Click(Sender: TObject);
var vawal_l, vawal_p, vkirim_l, vkirim_p, vterima_prod_l, vterima_prod_p, vterima_retur_l, vterima_retur_p, vterima_afval_l, vterima_afval_p, vkoreksi_l, vkoreksi_p, vakhir_l, vakhir_p : Real;
begin
  if QBrowse.QBEMode then
  begin
    QBrowse.ExecuteQBE;
    wwDBGrid2.Options:=wwDBGrid2.Options+[dgRowSelect,dgAlwaysShowSelection];
      {azmi}
      QBrowse.Open;
      vawal_l:=0;
      vkirim_l:=0;
      vterima_prod_l:=0;
      vterima_retur_l:=0;
      vterima_afval_l:=0;
      vkoreksi_l:=0;
      vakhir_l:=0;
      vawal_p:=0;
      vkirim_p:=0;
      vterima_prod_p:=0;
      vterima_retur_p:=0;
      vterima_afval_p:=0;
      vkoreksi_p:=0;
      vakhir_p:=0;
      while not QBrowse.Eof do
      begin
        vawal_l:=vawal_l+QBrowseLUSI_AWAL.AsFloat;
        vkirim_l:=vkirim_l+QBrowseLUSI_KIRIM.AsFloat;
        vterima_prod_l:=vterima_prod_l+QBrowseLUSI_TERIMA_PRODUKSI.AsFloat;
        vterima_retur_l:=vterima_retur_l+QBrowseLUSI_TERIMA_RETUR.AsFloat;
        vterima_afval_l:=vterima_afval_l+QBrowseLUSI_TERIMA_AFVAL.AsFloat;
        vkoreksi_l:=vkoreksi_l+QBrowseLUSI_KOREKSI.AsFloat;
        vakhir_l:=vakhir_l+QBrowseLUSI_AKHIR.AsFloat;

        vawal_p:=vawal_p+QBrowsePAKAN_AWAL.AsFloat;
        vkirim_p:=vkirim_p+QBrowsePAKAN_KIRIM.AsFloat;
        vterima_prod_p:=vterima_prod_p+QBrowsePAKAN_TERIMA_PRODUKSI.AsFloat;
        vterima_retur_p:=vterima_retur_p+QBrowsePAKAN_TERIMA_RETUR.AsFloat;
        vterima_afval_p:=vterima_afval_p+QBrowsePAKAN_TERIMA_AFVAL.AsFloat;
        vkoreksi_p:=vkoreksi_p+QBrowsePAKAN_KOREKSI.AsFloat;
        vakhir_p:=vakhir_p+QBrowsePAKAN_AKHIR.AsFloat;

        QBrowse.Next;
      end;
      QBrowse.EnableControls;
      wwDBGrid2.ColumnByName('LUSI_AWAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vawal_l);
      wwDBGrid2.ColumnByName('LUSI_KIRIM').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkirim_l);
      wwDBGrid2.ColumnByName('LUSI_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_prod_l);
      wwDBGrid2.ColumnByName('LUSI_TERIMA_RETUR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_retur_l);
      wwDBGrid2.ColumnByName('LUSI_TERIMA_AFVAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_afval_l);
      wwDBGrid2.ColumnByName('LUSI_KOREKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkoreksi_l);
      wwDBGrid2.ColumnByName('LUSI_AKHIR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vakhir_l);

      wwDBGrid2.ColumnByName('PAKAN_AWAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vawal_p);
      wwDBGrid2.ColumnByName('PAKAN_KIRIM').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkirim_p);
      wwDBGrid2.ColumnByName('PAKAN_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_prod_p);
      wwDBGrid2.ColumnByName('PAKAN_TERIMA_RETUR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_retur_p);
      wwDBGrid2.ColumnByName('PAKAN_TERIMA_AFVAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_afval_p);
      wwDBGrid2.ColumnByName('PAKAN_KOREKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkoreksi_p);
      wwDBGrid2.ColumnByName('PAKAN_AKHIR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vakhir_p);
      
      LabelBanner.Caption:='Data : '+FormatFloat('#,#',QBrowse.RecordCount)+' Records';
    {azmi}
  end;

end;

procedure TLaporanDeptKemitraanFrm.ppHeaderBand1BeforePrint(Sender: TObject);
begin
 ppNo.AsInteger:=0;

  end;

procedure TLaporanDeptKemitraanFrm.BtnPrintBrowseClick(Sender: TObject);
begin
  if vTglAwal.Date>vTglAkhir.DateTime then
    ShowMessage('Tgl. Akhir harus lebih besar dari Tgl. Awal !')
    else
    begin
      ppReportBrowse.Print;
    end;
end;

procedure TLaporanDeptKemitraanFrm.ppTitleBand1BeforePrint(Sender: TObject);
begin
//ppNo2.AsInteger:=0;
  ppNo.AsInteger:=0;
  ppNamaLaporan.Caption:='LAPORAN STOK BARANG DALAM PROSES KEMITRAAN';
  ppPeriode.Caption:=vTglAwal.Text+' s/d '+vTglAkhir.Text;
  DMFrm.QTime.Close;
  DMFrm.QTime.Open;
  ppUserCetak.Caption:=DMFrm.QTimeVUSER_CETAK.AsString;
end;

procedure TLaporanDeptKemitraanFrm.TabSheet1Show(Sender: TObject);
begin
  {QMaster.Close;
  QMaster.SetVariable('myparam1',QBrowseIBUKTI.AsInteger);
  QMaster.SetVariable('myparam2',QBrowseNO_NOTA.AsString);
  QMaster.Open;
  QDetail.Close;
  QDetail.Open;
  EditCari.Text:=QBrowseIBUKTI.AsString;
  if (QBrowseNO_NOTA.AsString<>'') then EditCari.Text:=QBrowseNO_NOTA.AsString;
  wwDBGrid1UpdateFooter(nil);}
end;

procedure TLaporanDeptKemitraanFrm.wwDBGrid1Enter(Sender: TObject);
begin
  if QMaster.State<>dsBrowse then
  try
    QMaster.Post;
  except
    ShowMessage('Maaf, ada masalah di pengisian MASTER !');
  end;
end;

procedure TLaporanDeptKemitraanFrm.wwDBGrid2DblClick(Sender: TObject);
begin
  TabSheet1.Show;
end;

procedure TLaporanDeptKemitraanFrm.Label5Click(Sender: TObject);
begin
  if DMFrm.FontDialog1.Execute then
  begin
    wwDBGrid2.Font.Name:=DMFrm.FontDialog1.Font.Name;
    wwDBGrid2.Font.Size:=DMFrm.FontDialog1.Font.Size;
    wwDBGrid2.Font.Color:=DMFrm.FontDialog1.Font.Color;
    wwDBGrid2.Font.Style:=DMFrm.FontDialog1.Font.Style;
  end;
end;

procedure TLaporanDeptKemitraanFrm.LookItemEnter(Sender: TObject);
begin
  (sender as TwwDBLookupComboDlg).LookupTable.Close;
  (sender as TwwDBLookupComboDlg).LookupTable.Open;

end;

procedure TLaporanDeptKemitraanFrm.QTransaksiBeforeOpen(DataSet: TDataSet);
begin
  QTransaksi.DeclareVariable('kd_transaksi', otString);
  QTransaksi.SQL.Text:='select a.* from '+cUserTabel+'vtransaksi a where a.kd_transaksi=:kd_transaksi';
  QTransaksi.SetVariable('kd_transaksi',vkode);
end;

{procedure THasilSoftConesFrm.ppPageStyle1BeforePrint(Sender: TObject);
begin
  ppNo2.AsInteger:=0;
end;   }

procedure TLaporanDeptKemitraanFrm.QMasterAfterPost(DataSet: TDataSet);
begin
  {
  PageControl1.Pages[1].TabVisible:=QMaster.IsEmpty or (QMasterISPOST.AsString='1');
  PageControl1.Pages[2].TabVisible:=QMaster.IsEmpty or (QMasterISPOST.AsString='1');
  if QMasterISPOST.AsString='1' then
  begin
      QProc_Update_PO.Close;
      QProc_Update_PO.Execute;
  end;
  }
end;

procedure TLaporanDeptKemitraanFrm.QMasterBeforeInsert(DataSet: TDataSet);
begin
  if DataSet['ISPOST']='0' then
  begin
    if MessageDlg('Data belum di-POSTING, batalkan ?', mtWarning, [mbYes, mbNo],0)=mrYes then
      DataSet.Delete
      else
        Abort;
  end;

end;

procedure TLaporanDeptKemitraanFrm.QProcUpdateMutasiBeforeQuery(Sender: TOracleQuery);
begin
  {QProc_Update_PO.SetVariable('NO_PO',QMasterNO_BUKTI.AsString);}
end;

procedure TLaporanDeptKemitraanFrm.FormShow(Sender: TObject);
begin
//  PanelHeader.Caption:=QTransaksiKD_TRANSAKSI.AsString+'. '+UpperCase(Caption);
end;

procedure TLaporanDeptKemitraanFrm.ppDetailBand2BeforePrint(Sender: TObject);
begin
 // ppNo2.AsInteger:=ppNo2.AsInteger+1;
end;

procedure TLaporanDeptKemitraanFrm.ppDetailBand1BeforePrint(Sender: TObject);
begin
  ppNo.AsInteger:=ppNo.AsInteger+1;
end;

procedure TLaporanDeptKemitraanFrm.QBrowseCalcFields(DataSet: TDataSet);
begin
  {if copy(QBrowseKETERANGAN.AsString,1,4) = '30/2' then
    begin
      QBrowseeffisiensi.AsFloat := (QBrowseQTY2.AsFloat/QBrowseSPEED_PER_MNT2.AsFloat)/425*100;
    end
  else
    begin
      QBrowseeffisiensi.AsFloat := (QBrowseQTY2.AsFloat/QBrowseSPEED_PER_MNT.AsFloat)/425*100;
    end;

     QBrowseRASIO3.AsFloat:=QBrowseQTY1.AsFloat/QBrowseQTY2.AsFloat; }

end;

procedure TLaporanDeptKemitraanFrm.wwDBGrid2UpdateFooter(Sender: TObject);
var vawal_l, vawal_p, vkirim_l, vkirim_p, vterima_prod_l, vterima_prod_p, vterima_retur_l, vterima_retur_p, vterima_afval_l, vterima_afval_p, vkoreksi_l, vkoreksi_p, vakhir_l, vakhir_p : Real;
begin
      QBrowse.Open;
      vawal_l:=0;
      vkirim_l:=0;
      vterima_prod_l:=0;
      vterima_retur_l:=0;
      vterima_afval_l:=0;
      vkoreksi_l:=0;
      vakhir_l:=0;
      vawal_p:=0;
      vkirim_p:=0;
      vterima_prod_p:=0;
      vterima_retur_p:=0;
      vterima_afval_p:=0;
      vkoreksi_p:=0;
      vakhir_p:=0;
      while not QBrowse.Eof do
      begin
        vawal_l:=vawal_l+QBrowseLUSI_AWAL.AsFloat;
        vkirim_l:=vkirim_l+QBrowseLUSI_KIRIM.AsFloat;
        vterima_prod_l:=vterima_prod_l+QBrowseLUSI_TERIMA_PRODUKSI.AsFloat;
        vterima_retur_l:=vterima_retur_l+QBrowseLUSI_TERIMA_RETUR.AsFloat;
        vterima_afval_l:=vterima_afval_l+QBrowseLUSI_TERIMA_AFVAL.AsFloat;
        vkoreksi_l:=vkoreksi_l+QBrowseLUSI_KOREKSI.AsFloat;
        vakhir_l:=vakhir_l+QBrowseLUSI_AKHIR.AsFloat;

        vawal_p:=vawal_p+QBrowsePAKAN_AWAL.AsFloat;
        vkirim_p:=vkirim_p+QBrowsePAKAN_KIRIM.AsFloat;
        vterima_prod_p:=vterima_prod_p+QBrowsePAKAN_TERIMA_PRODUKSI.AsFloat;
        vterima_retur_p:=vterima_retur_p+QBrowsePAKAN_TERIMA_RETUR.AsFloat;
        vterima_afval_p:=vterima_afval_p+QBrowsePAKAN_TERIMA_AFVAL.AsFloat;
        vkoreksi_p:=vkoreksi_p+QBrowsePAKAN_KOREKSI.AsFloat;
        vakhir_p:=vakhir_p+QBrowsePAKAN_AKHIR.AsFloat;
        QBrowse.Next;
      end;
      QBrowse.EnableControls;
      wwDBGrid2.ColumnByName('LUSI_AWAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vawal_l);
      wwDBGrid2.ColumnByName('LUSI_KIRIM').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkirim_l);
      wwDBGrid2.ColumnByName('LUSI_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_prod_l);
      wwDBGrid2.ColumnByName('LUSI_TERIMA_RETUR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_retur_l);
      wwDBGrid2.ColumnByName('LUSI_TERIMA_AFVAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_afval_l);
      wwDBGrid2.ColumnByName('LUSI_KOREKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkoreksi_l);
      wwDBGrid2.ColumnByName('LUSI_AKHIR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vakhir_l);

      wwDBGrid2.ColumnByName('PAKAN_AWAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vawal_p);
      wwDBGrid2.ColumnByName('PAKAN_KIRIM').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkirim_p);
      wwDBGrid2.ColumnByName('PAKAN_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_prod_p);
      wwDBGrid2.ColumnByName('PAKAN_TERIMA_RETUR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_retur_p);
      wwDBGrid2.ColumnByName('PAKAN_TERIMA_AFVAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vterima_afval_p);
      wwDBGrid2.ColumnByName('PAKAN_KOREKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vkoreksi_p);
      wwDBGrid2.ColumnByName('PAKAN_AKHIR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vakhir_p);
end;


procedure TLaporanDeptKemitraanFrm.wwDBGrid1CalcCellColors(Sender: TObject;
  Field: TField; State: TGridDrawState; Highlight: Boolean; AFont: TFont;
  ABrush: TBrush);
begin
  if not Highlight then
    if (Sender as TwwDBGrid).ColumnByName(Field.FieldName).ReadOnly then
    begin
      ABrush.Color:=DMFrm.vclGridRead;
      AFont.Color:=DMFrm.vclGridReadFont;
    end
    else
    begin
      ABrush.Color:=DMFrm.vclGridEdit;
      AFont.Color:=DMFrm.vclGridEditFont;
    end;
end;

procedure TLaporanDeptKemitraanFrm.QBrowseFilterRecord(DataSet: TDataSet;
  var Accept: Boolean);
begin
Accept:=
    ((QBrowseLUSI_AWAL.AsFloat)<>0) or
    ((QBrowseLUSI_KIRIM.AsFloat)<>0) or
    ((QBrowseLUSI_TERIMA_PRODUKSI.AsFloat)<>0) or
    ((QBrowseLUSI_TERIMA_RETUR.AsFloat)<>0) or
    ((QBrowseLUSI_TERIMA_AFVAL.AsFloat)<>0) or
    ((QBrowseLUSI_KOREKSI.AsFloat)<>0) or
    ((QBrowseLUSI_AKHIR.AsFloat)<>0) or
    ((QBrowsePAKAN_AWAL.AsFloat)<>0) or
    ((QBrowsePAKAN_KIRIM.AsFloat)<>0) or
    ((QBrowsePAKAN_TERIMA_PRODUKSI.AsFloat)<>0) or
    ((QBrowsePAKAN_TERIMA_RETUR.AsFloat)<>0) or
    ((QBrowsePAKAN_TERIMA_AFVAL.AsFloat)<>0) or
    ((QBrowsePAKAN_KOREKSI.AsFloat)<>0) or
    ((QBrowsePAKAN_AKHIR.AsFloat)<>0);
end;

procedure TLaporanDeptKemitraanFrm.cbAdaTransaksiClick(Sender: TObject);
begin
   QBrowse.Filtered:=cbAdaTransaksi.Checked;
end;

procedure TLaporanDeptKemitraanFrm.vTglAwal2Change(Sender: TObject);
begin
  vTglAkhir2.DateTime:=EndOfTheMonth(vTglAwal2.Date);
end;

procedure TLaporanDeptKemitraanFrm.BitBtn1Click(Sender: TObject);
var
  vqty1, vqty2 : real;
begin
     QBrowse1.DisableControls;
     QBrowse1.SetVariable('pawal', vTglAwal2.Date);
     QBrowse1.SetVariable('pakhir', vTglAkhir2.Date);
     QBrowse1.Close;
     QBrowse1.Open;
     QBrowse1.EnableControls;
     vqty1:=0;
     vqty2:=0;
     while not QBrowse1.Eof do
     begin
       vqty1:=vqty1+QBrowse1PEMASUKAN.AsFloat;
       vqty2:=vqty2+QBrowse1PENGELUARAN.AsFloat;
       QBrowse1.Next;
     end;
     QBrowse1.EnableControls;
     wwDBGrid3.ColumnByName('AWAL').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',0);
     wwDBGrid3.ColumnByName('PEMASUKAN').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty1);
     wwDBGrid3.ColumnByName('PENGELUARAN').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty1);
     wwDBGrid3.ColumnByName('AKHIR').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',0);
end;

procedure TLaporanDeptKemitraanFrm.ppSummaryBand1BeforePrint(
  Sender: TObject);
begin
  ppLabel42.Caption:='Pekalongan, '+FormatDateTime('mmmm yyyy',vTglAwal.Date);
end;

procedure TLaporanDeptKemitraanFrm.BitBtn2Click(Sender: TObject);
var vt1, vt2, vt3, vt4, vt5 : Real;
begin
  if vTglAwal0.Date>vTglAkhir0.DateTime then
    ShowMessage('Tgl. Akhir harus lebih besar dari Tgl. Awal !')
    else
    begin

      QProcRekapBulanan.Close;
      QProcRekapBulanan.SetVariable('pawal', vTglAwal0.Date);
      QProcRekapBulanan.SetVariable('pakhir', vTglAkhir0.Date);
      QProcRekapBulanan.Execute;

      if QMaster.QBEMode then QMaster.QBEMode:=False;
      QMaster.DisableControls;
      QMaster.Close;
      QMaster.Open;
      vt1:=0;
      vt2:=0;
      vt3:=0;
      vt4:=0;
      vt5:=0;
      while not QMaster.Eof do
      begin
        vt1:=vt1+QMasterQTY_KG.AsFloat;
        vt2:=vt2+QMasterQTY_PTG.AsFloat;
        vt3:=vt3+QMasterQTY_SPRING.AsFloat;
        vt4:=vt4+QMasterQTY_BEAM.AsFloat;
        vt5:=vt5+QMasterQTY_BS.AsFloat;
        QMaster.Next;
      end;
      QMaster.EnableControls;
      wwDBGrid1.ColumnByName('QTY_KG').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vt1);
      wwDBGrid1.ColumnByName('QTY_PTG').FooterValue:=FormatFloat('#,0;-#,0;-',vt2);
      wwDBGrid1.ColumnByName('QTY_SPRING').FooterValue:=FormatFloat('#,0;-#,0;-',vt3);
      wwDBGrid1.ColumnByName('QTY_BEAM').FooterValue:=FormatFloat('#,0;-#,0;-',vt4);
      wwDBGrid1.ColumnByName('QTY_BS').FooterValue:=FormatFloat('#,0;-#,0;-',vt5);
      wwDBGrid1.ColumnByName('QTY_PERSEN').FooterValue:=FormatFloat('#,0.00;-#,0.00;-',vt5/vt2*100);
    end;
end;

procedure TLaporanDeptKemitraanFrm.vtglAwal0Change(Sender: TObject);
begin
  vTglAkhir0.DateTime:=EndOfTheMonth(vTglAwal0.Date);
end;

procedure TLaporanDeptKemitraanFrm.Label9Click(Sender: TObject);
begin
  if DMFrm.FontDialog1.Execute then
  begin
    wwDBGrid1.Font.Name:=DMFrm.FontDialog1.Font.Name;
    wwDBGrid1.Font.Size:=DMFrm.FontDialog1.Font.Size;
    wwDBGrid1.Font.Color:=DMFrm.FontDialog1.Font.Color;
    wwDBGrid1.Font.Style:=DMFrm.FontDialog1.Font.Style;
  end;
end;

procedure TLaporanDeptKemitraanFrm.wwDBSpinEdit2Change(Sender: TObject);
begin
  wwDBGrid1.RowHeightPercent:=Round(wwDBSpinLine2.Value);
end;

procedure TLaporanDeptKemitraanFrm.QMasterCalcFields(DataSet: TDataSet);
var vqty_ptg: Real;
begin
  if QMasterQTY_PTG.AsFloat >= 1 then vqty_ptg:=QMasterQTY_PTG.AsFloat else vqty_ptg:=1;
  QMasterQTY_PERSEN.AsFloat:=(QMasterQTY_BS.AsFloat/vqty_ptg)*100;
end;

end.
