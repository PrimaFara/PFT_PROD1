unit RekapMitra;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  DateUtils, Mask, wwdbedit, Wwdbspin, DB, Wwdatsrc, OracleData, Dialogs,
  Oracle, Buttons, wwSpeedButton, wwDBNavigator, wwclearpanel, Grids,
  Wwdbigrd, Wwdbgrid, StdCtrls, wwdbdatetimepicker, ComCtrls, ExtCtrls,
  ppViewr, ppDB, ppDBPipe, ppComm, ppRelatv, ppProd, ppClass, ppReport,
  ppVar, ppCtrls, ppBands, ppPrnabl, ppCache, ppEndUsr, ppModule,
  daDataModule, DBCtrls, wwdblook, Wwdbdlg, Wwdotdot, Wwdbcomb, ppBarCod,
  wwcheckbox, ppStrtch, ppMemo, raCodMod, wwriched, QRCtrls, QuickRpt,
  ppParameter;

type
  TRekapMitraFrm = class(TForm)
    PanelMain: TPanel;
    PanelHeader: TPanel;
    PanelFilter: TPanel;
    LabelBanner: TLabel;
    PanelBrowse: TPanel;
    PanelFooter2: TPanel;
    BtnOk: TBitBtn;
    wwDBNavigator1: TwwDBNavigator;
    wwDBNavigator1PriorPage: TwwNavButton;
    wwDBNavigator1NextPage: TwwNavButton;
    wwDBNavigator1SaveBookmark: TwwNavButton;
    wwDBNavigator1RestoreBookmark: TwwNavButton;
    QBrowse: TOracleDataSet;
    dsQBrowse: TwwDataSource;
    BtnExport: TBitBtn;
    BtnPrintBrowse: TBitBtn;
    Label1: TLabel;
    ppReportBrowse: TppReport;
    ppDBQBrowseDetail: TppDBPipeline;
    ppDesigner1: TppDesigner;
    BtnDesign2: TBitBtn;
    ppDBPerusahaan: TppDBPipeline;
    DBText3: TDBText;
    ppTitleBand1: TppTitleBand;
    ppNamaLaporan: TppLabel;
    ppLabel9: TppLabel;
    ppPeriode: TppLabel;
    ppDBText12: TppDBText;
    ppDBText13: TppDBText;
    ppDBText14: TppDBText;
    ppUserCetak: TppLabel;
    ppHeaderBand1: TppHeaderBand;
    ppLabel6: TppLabel;
    ppLabel7: TppLabel;
    ppDetailBand1: TppDetailBand;
    ppDBText8: TppDBText;
    ppFooterBand1: TppFooterBand;
    ppSystemVariable1: TppSystemVariable;
    ppSummaryBand1: TppSummaryBand;
    ppDBText6: TppDBText;
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
    ppDBQTransaksi: TppDBPipeline;
    dsQTransaksi: TwwDataSource;
    ppDBText19: TppDBText;
    ppDBText42: TppDBText;
    ppDBText43: TppDBText;
    ppDBText44: TppDBText;
    ppDBText45: TppDBText;
    ppDBText46: TppDBText;
    ppDBText47: TppDBText;
    ppDBText48: TppDBText;
    ppDBText49: TppDBText;
    ppLabel1: TppLabel;
    ppLabel5: TppLabel;
    ppDBMemo1: TppDBMemo;
    ppDBText7: TppDBText;
    ppLabel12: TppLabel;
    BtnClose1: TBitBtn;
    ppLblNomer: TppLabel;
    ppDBText3: TppDBText;
    ppLabel8: TppLabel;
    ppLabel10: TppLabel;
    ppLabel11: TppLabel;
    ppDBText5: TppDBText;
    ppDBText10: TppDBText;
    vTglAwal: TwwDBDateTimePicker;
    Label3: TLabel;
    ppLabel2: TppLabel;
    ppDBText1: TppDBText;
    ppLabel3: TppLabel;
    ppDBText2: TppDBText;
    ppDBMemo2: TppDBMemo;
    ppLabel4: TppLabel;
    ppLabel23: TppLabel;
    ppLabel28: TppLabel;
    ppLabel30: TppLabel;
    ppDBText4: TppDBText;
    ppDBText11: TppDBText;
    ppDBText27: TppDBText;
    QDump: TOracleQuery;
    vTglAkhir: TwwDBDateTimePicker;
    Label4: TLabel;
    ppDBText28: TppDBText;
    ppLine1: TppLine;
    ppLine2: TppLine;
    ppLine3: TppLine;
    ppLine4: TppLine;
    ppLine5: TppLine;
    ppLine6: TppLine;
    ppLine7: TppLine;
    ppLine8: TppLine;
    ppLine9: TppLine;
    ppLine10: TppLine;
    ppLine11: TppLine;
    ppLine12: TppLine;
    ppDBText50: TppDBText;
    ppDBText51: TppDBText;
    ppLine13: TppLine;
    ppLabel31: TppLabel;
    ppLabel32: TppLabel;
    ppLine14: TppLine;
    ppLine15: TppLine;
    ppLine16: TppLine;
    ppLine17: TppLine;
    ppLine18: TppLine;
    ppShape1: TppShape;
    ppShape2: TppShape;
    ppShape3: TppShape;
    ppShape4: TppShape;
    ppLabel33: TppLabel;
    ppShape5: TppShape;
    ppShape8: TppShape;
    ppLabel34: TppLabel;
    ppShape9: TppShape;
    ppShape10: TppShape;
    ppLabel35: TppLabel;
    ppLine19: TppLine;
    ppLine20: TppLine;
    ppShape11: TppShape;
    ppLine21: TppLine;
    ppLine22: TppLine;
    ppLine23: TppLine;
    ppLine24: TppLine;
    ppLabel36: TppLabel;
    ppShape6: TppShape;
    ppShape7: TppShape;
    ppShape12: TppShape;
    ppLine25: TppLine;
    ppLine26: TppLine;
    ppLabel37: TppLabel;
    ppLine27: TppLine;
    QuickRep1: TQuickRep;
    DetailBand1: TQRBand;
    QRDBText3: TQRDBText;
    QRExpr9: TQRExpr;
    SummaryBand1: TQRBand;
    QRLabel2: TQRLabel;
    QRExpr1: TQRExpr;
    QRDBText4: TQRDBText;
    QRExpr2: TQRExpr;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRExpr3: TQRExpr;
    QRDBText5: TQRDBText;
    QRLabel1: TQRLabel;
    QRExpr4: TQRExpr;
    QRShape4: TQRShape;
    QRLabel3: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    TitleBand1: TQRBand;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    PageFooterBand1: TQRBand;
    QRDBText6: TQRDBText;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRDBText7: TQRDBText;
    QRLabel16: TQRLabel;
    QRExpr5: TQRExpr;
    QRShape14: TQRShape;
    QRDBText1: TQRDBText;
    QRLabel4: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QuickRep2: TQuickRep;
    QRBand1: TQRBand;
    QRDBText2: TQRDBText;
    QRBand2: TQRBand;
    QRLabel21: TQRLabel;
    QRBand3: TQRBand;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRShape3: TQRShape;
    QRLabel38: TQRLabel;
    QRShape7: TQRShape;
    QRLabel49: TQRLabel;
    QRShape6: TQRShape;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel32: TQRLabel;
    QRBand4: TQRBand;
    QRSysData2: TQRSysData;
    QRShape9: TQRShape;
    QRDBText8: TQRDBText;
    QRDBText13: TQRDBText;
    QRDBText14: TQRDBText;
    QRDBText15: TQRDBText;
    QRDBText16: TQRDBText;
    QRBand5: TQRBand;
    QRLabel39: TQRLabel;
    QRLabel40: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel42: TQRLabel;
    QRLabel44: TQRLabel;
    QRLabel45: TQRLabel;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRExpr6: TQRExpr;
    QRExpr7: TQRExpr;
    QRExpr8: TQRExpr;
    QRExpr10: TQRExpr;
    QRDBText17: TQRDBText;
    QRBand6: TQRBand;
    QRDBText18: TQRDBText;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRShape5: TQRShape;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
    QRShape8: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape12: TQRShape;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRDBText9: TQRDBText;
    QRDBText10: TQRDBText;
    QRDBText11: TQRDBText;
    QRDBText12: TQRDBText;
    QRShape15: TQRShape;
    QRShape30: TQRShape;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRExpr11: TQRExpr;
    QRExpr12: TQRExpr;
    QRExpr13: TQRExpr;
    QRExpr14: TQRExpr;
    BitBtnPrint2: TBitBtn;
    QRShape13: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape26: TQRShape;
    QRBand7: TQRBand;
    QRLabel35: TQRLabel;
    QRLabel36: TQRLabel;
    QRLabel37: TQRLabel;
    QRLabel43: TQRLabel;
    QRLabel46: TQRLabel;
    QRLabel47: TQRLabel;
    QRShape27: TQRShape;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRShape33: TQRShape;
    QRShape34: TQRShape;
    QRShape35: TQRShape;
    QRExpr15: TQRExpr;
    QRExpr16: TQRExpr;
    QRExpr17: TQRExpr;
    QRExpr18: TQRExpr;
    QRDBText19: TQRDBText;
    QRShape36: TQRShape;
    QRShape37: TQRShape;
    QRShape38: TQRShape;
    QRShape39: TQRShape;
    QRExpr19: TQRExpr;
    QRExpr20: TQRExpr;
    QRExpr21: TQRExpr;
    QRExpr22: TQRExpr;
    QRShape40: TQRShape;
    QRShape41: TQRShape;
    QRShape42: TQRShape;
    QRLabel50: TQRLabel;
    QRLabel51: TQRLabel;
    QBrowse2: TOracleDataSet;
    dsQBrowse2: TwwDataSource;
    PageControl2: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    LBanner2: TLabel;
    Panel1: TPanel;
    Label7: TLabel;
    Label8: TLabel;
    BitBtn1: TBitBtn;
    vtglAwal2: TwwDBDateTimePicker;
    vtglAkhir2: TwwDBDateTimePicker;
    wwDBGrid2: TwwDBGrid;
    wwDBGrid1: TwwDBGrid;
    QProcPerTgl: TOracleQuery;
    QBrowseTGL: TDateTimeField;
    QBrowseNO_NOTA: TStringField;
    QBrowseMITRA: TStringField;
    QBrowseNO_SERI_BEAM: TStringField;
    QBrowseKP: TStringField;
    QBrowseLUSI: TFloatField;
    QBrowsePAKAN: TFloatField;
    QBrowseNAMA_ITEM: TStringField;
    QBrowseISPOST: TStringField;
    Panel2: TPanel;
    Label5: TLabel;
    wwDBSpinLine2: TwwDBSpinEdit;
    BtnFind: TSpeedButton;
    BtnOk2: TSpeedButton;
    Panel3: TPanel;
    wwDBSpinEdit1: TwwDBSpinEdit;
    Label6: TLabel;
    QBrowse2TGL: TDateTimeField;
    QBrowse2NO_NOTA: TStringField;
    QBrowse2MITRA: TStringField;
    QBrowse2KD_ITEM: TStringField;
    QBrowse2KD_WARNA: TStringField;
    QBrowse2LUSI: TFloatField;
    QBrowse2PAKAN: TFloatField;
    QBrowse2NAMA_ITEM: TStringField;
    QBrowse2ISPOST: TStringField;
    QBrowse2WARNA: TStringField;
    BtnFind2: TSpeedButton;
    SpeedButton2: TSpeedButton;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    wwDBGrid3: TwwDBGrid;
    Panel4: TPanel;
    Label2: TLabel;
    Label9: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton3: TSpeedButton;
    BitBtn2: TBitBtn;
    vTglAwal3: TwwDBDateTimePicker;
    vTglAkhir3: TwwDBDateTimePicker;
    Panel5: TPanel;
    Label10: TLabel;
    wwDBSpinEdit2: TwwDBSpinEdit;
    LBanner3: TLabel;
    QBrowse3: TOracleDataSet;
    dsQBrowse3: TwwDataSource;
    QBrowse3TGL: TDateTimeField;
    QBrowse3NO_NOTA: TStringField;
    QBrowse3NAMA_MITRA: TStringField;
    QBrowse3KD_PRODUKSI: TStringField;
    QBrowse3KONSTRUKSI: TStringField;
    QBrowse3QTY_PTG: TFloatField;
    QBrowse3LS_TERIMA_PRODUKSI: TFloatField;
    QBrowse3PK_TERIMA_PRODUKSI: TFloatField;
    QBrowse3PK_TERIMA_PRODUKSI2: TFloatField;
    QBrowse3OPR_INSERT: TStringField;
    QBrowseOPR_INSERT: TStringField;
    QBrowse2OPR_INSERT: TStringField;
    Panel6: TPanel;
    Label11: TLabel;
    Label12: TLabel;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    BitBtn3: TBitBtn;
    vTglAwal4: TwwDBDateTimePicker;
    vTglAkhir4: TwwDBDateTimePicker;
    Panel7: TPanel;
    Label13: TLabel;
    wwDBSpinEdit3: TwwDBSpinEdit;
    wwDBGrid4: TwwDBGrid;
    LBanner4: TLabel;
    QBrowse4: TOracleDataSet;
    dsQBrowse4: TwwDataSource;
    QBrowse4TGL: TDateTimeField;
    QBrowse4NO_NOTA: TStringField;
    QBrowse4NAMA_MITRA: TStringField;
    QBrowse4NO_SERI_BEAM: TStringField;
    QBrowse4QTY_LUSI1: TFloatField;
    QBrowse4OPR_INSERT: TStringField;
    QBrowse4KP: TStringField;
    QBrowse4KONSTRUKSI: TStringField;
    QBrowse4NAMA_ITEM: TStringField;
    QBrowse4ISPOST: TStringField;
    wwDBGrid5: TwwDBGrid;
    Panel8: TPanel;
    Label14: TLabel;
    Label15: TLabel;
    SpeedButton6: TSpeedButton;
    SpeedButton7: TSpeedButton;
    BitBtn4: TBitBtn;
    vTglAwal5: TwwDBDateTimePicker;
    vTglAkhir5: TwwDBDateTimePicker;
    Panel9: TPanel;
    Label16: TLabel;
    wwDBSpinEdit4: TwwDBSpinEdit;
    LBanner5: TLabel;
    QBrowse5: TOracleDataSet;
    dsQBrowse5: TwwDataSource;
    QBrowse5TGL: TDateTimeField;
    QBrowse5NO_NOTA: TStringField;
    QBrowse5MITRA: TStringField;
    QBrowse5KD_ITEM: TStringField;
    QBrowse5KD_WARNA: TStringField;
    QBrowse5LUSI: TFloatField;
    QBrowse5PAKAN: TFloatField;
    QBrowse5NAMA_ITEM: TStringField;
    QBrowse5OPR_INSERT: TStringField;
    QBrowse5ISPOST: TStringField;
    QBrowse5WARNA: TStringField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BtnOkClick(Sender: TObject);
    procedure wwDBGrid2TitleButtonClick(Sender: TObject;
      AFieldName: String);
    procedure BtnClose1Click(Sender: TObject);
    procedure wwDBSpinLine2Change(Sender: TObject);
    procedure BtnFindClick(Sender: TObject);
    procedure BtnOk2Click(Sender: TObject);
    procedure BtnPrintBrowseClick(Sender: TObject);
    procedure BtnPrintBrowse1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure BtnDesign2Click(Sender: TObject);
    procedure ppTitleBand1BeforePrint(Sender: TObject);
    procedure Label5Click(Sender: TObject);
    procedure LookItemEnter(Sender: TObject);
    procedure ppDetailBand1BeforePrint(Sender: TObject);
    procedure QTransaksiBeforeOpen(DataSet: TDataSet);
    procedure vTglAwalChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure QBrowseAfterScroll(DataSet: TDataSet);
    procedure BitBtnPrint2Click(Sender: TObject);
    procedure cbAdaTransaksiClick(Sender: TObject);
    procedure TitleBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand5BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand6BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand8BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure PageFooterBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QuickRep1AfterPreview(Sender: TObject);
    procedure QuickRep1AfterPrint(Sender: TObject);
    procedure wwDBGrid1UpdateFooter(Sender: TObject);
    procedure wwDBGrid2UpdateFooter(Sender: TObject);
    procedure SummaryBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QuickRep1BeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure TabSheet3Show(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure vtglAwal2Change(Sender: TObject);
    procedure QBrowse2AfterScroll(DataSet: TDataSet);
    procedure TabSheet2Show(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure Label6Click(Sender: TObject);
    procedure wwDBSpinEdit1Change(Sender: TObject);
    procedure BtnExportClick(Sender: TObject);
    procedure BtnFind2Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure TabSheet1Show(Sender: TObject);
    procedure QBrowse3AfterScroll(DataSet: TDataSet);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure vTglAwal3Change(Sender: TObject);
    procedure vTglAwal4Change(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure QBrowse4AfterScroll(DataSet: TDataSet);
    procedure vTglAwal5Change(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SpeedButton7Click(Sender: TObject);
  private
    { Private declarations }
    vorder, SelectedFont, vkode, vitem : String;
    t1, t2, t3, t4, t5, t6 : real;
    t7, t8, t9, t10, t11, t12 : real;
  public
    { Public declarations }

  end;

var
  RekapMitraFrm: TRekapMitraFrm;

Procedure ShowForm(pNamaMenu:String; pkode : String; pjudul : string; pbrg : string);

implementation

uses DM, Pembelian, Kriteria_Tanggal1, KartuStokBB, InfoWIPPengeringan,
  Math, LapProduksi, ComObj;

{$R *.dfm}

procedure ShowForm(pNamaMenu:String; pkode : String; pjudul : string; pbrg : String);
var
  mychar : string[125];
Begin
// Hak Menu
  DMFrm.cHakInput:=DMFrm.QMenuUser.Lookup('NAMA_COMPONENT',pNamaMenu,'HAK_INPUT')='1';
  DMFrm.cBtnDesign:=DMFrm.QMenuUser.Lookup('NAMA_COMPONENT',pNamaMenu,'HAK_DESIGN')='1';
  DMFrm.cBtnExport:=DMFrm.QMenuUser.Lookup('NAMA_COMPONENT',pNamaMenu,'HAK_EXPORT')='1';

//  if InfoWIPPengeringanFrm=Nil then
  begin
    RekapMitraFrm:=TRekapMitraFrm.Create(Application);
    mychar:=pjudul;
    Delete(mychar,Pos('&',mychar),1);
    pjudul:=mychar;
    RekapMitraFrm.vkode:=pbrg;
    RekapMitraFrm.QTransaksi.Open;


    RekapMitraFrm.PanelHeader.Caption:=pjudul;
    RekapMitraFrm.Caption:=UpperCase(RekapMitraFrm.PanelHeader.Caption);
    {LapProduksiFrm.wwDBGrid2.IniAttributes.FileName:=DMFrm.sAppPath+Application.Title+'2.ini';
    LapProduksiFrm.wwDBGrid2.IniAttributes.SectionName:=LapProduksiFrm.Caption+'2';
    LapProduksiFrm.wwDBGrid2.IniAttributes.Enabled:=True;
    LapProduksiFrm.wwDBGrid2.LoadFromIniFile;
    DMFrm.ProcReadIni(Application.Title,LapProduksiFrm.Caption+'2',LapProduksiFrm.wwDBGrid2);
    LapProduksiFrm.wwDBSpinLine2.Value:=LapProduksiFrm.wwDBGrid2.RowHeightPercent;}

  end;

  RekapMitraFrm.Show;
end;

procedure TRekapMitraFrm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   DMFrm.ProcWtiteIni(Application.Title,Caption+'2',wwDBGrid1);
   Action:=caFree;
   LapProduksiFrm:=Nil;
   QuickRep1:=Nil;
end;

procedure TRekapMitraFrm.FormCreate(Sender: TObject);
begin
//barcode
     DMFrm.FontToUse := TFont.Create;
     DMFrm.selected := 'UCC 128';
     SelectedFont := 'CIA Code 128 Medium';
     StrPCopy(DMFrm.TempSelected, DMFrm.Selected);
     DMFrm.BType := 'C128';
     DMFrm.FontToUse.Size := 12;
     DMFrm.FontToUse.Name := SelectedFont;
//     LBarcode.Font := DMFrm.FontToUse;
//     QRBarcode11.Font := LoginFrm.FontToUse;
     DMFrm.BType := DMFrm.BType + '-';
     DMFrm.BType := DMFrm.BType + DMFrm.Format;
     DMFrm.BType := DMFrm.BType + '.BH';
// end barcode

  PanelMain.Color:=cWarnaPanelUtama;
  BtnClose1.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\CLOSE.Bmp');
  BtnOk.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\CHECK.Bmp');
//  BtnOk2.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\CHECK.Bmp');
//  BtnFind.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\FIND.Bmp');
  BtnDesign2.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\DESIGN.Bmp');
  BtnPrintBrowse.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\PRINT.Bmp');
//  BtnPrintBrowse1.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\PRINT.Bmp');
  BtnExport.Glyph.LoadFromFile(DMFrm.sAppPath+'Images\EXPORT.Bmp');
//Otoritas Button
//  BtnExport.Visible:=DMFrm.cBtnExport;
  BtnDesign2.Visible:=DMFrm.cBtnDesign;
end;

procedure TRekapMitraFrm.BtnOkClick(Sender: TObject);
var
  vqty1 : real;
begin
     QBrowse.DisableControls;
     QBrowse.SetVariable('ptgl', vTglAwal.Date);
     QBrowse.SetVariable('ptgl2', vTglAkhir.Date);
     QBrowse.Close;
     QBrowse.Open;
     QBrowse.EnableControls;
     vqty1:=0;
     while not QBrowse.Eof do
     begin
       vqty1:=vqty1+QBrowseLUSI.AsFloat;
       QBrowse.Next;
     end;
     QBrowse.EnableControls;
     LabelBanner.Caption:='Data : '+FormatFloat('#,#',QBrowse.RecordCount)+' Records';
     wwDBGrid1.ColumnByName('LUSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty1);
end;

procedure TRekapMitraFrm.wwDBGrid2TitleButtonClick(Sender: TObject;
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

procedure TRekapMitraFrm.BtnClose1Click(Sender: TObject);
begin
  Close;
end;

procedure TRekapMitraFrm.wwDBSpinLine2Change(Sender: TObject);
begin
    wwDBGrid1.RowHeightPercent:=Round(wwDBSpinLine2.Value);
end;

procedure TRekapMitraFrm.BtnFindClick(Sender: TObject);
begin
  if not QBrowse.QBEMode then
  begin
    wwDBGrid1.Options:=wwDBGrid1.Options-[dgRowSelect,dgAlwaysShowSelection];
    QBrowse.QBEMode:=TRUE;
  end
  else
    QBrowse.ClearQBE;
end;

procedure TRekapMitraFrm.BtnOk2Click(Sender: TObject);
var t1 : real;
begin
  if QBrowse.QBEMode then
  begin
    QBrowse.ExecuteQBE;
    wwDBGrid1.Options:=wwDBGrid1.Options+[dgRowSelect,dgAlwaysShowSelection];
    QBrowse.Open;
    t1:=0;
    while not QBrowse.Eof do
    begin
      t1:=t1+QBrowseLUSI.AsFloat;
      QBrowse.Next;
    end;
    QBrowse.EnableControls;
    LabelBanner.Caption:='Data : '+FormatFloat('#,#',QBrowse.RecordCount)+' Records';
    wwDBGrid1.ColumnByName('LUSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',t1);
  end;
end;

procedure TRekapMitraFrm.BtnPrintBrowseClick(Sender: TObject);
begin
  QuickRep1.Preview;
end;

procedure TRekapMitraFrm.BtnPrintBrowse1Click(Sender: TObject);
var checbox : byte;
begin

end;

procedure TRekapMitraFrm.Button1Click(Sender: TObject);
begin
  ppDesigner1.ShowModal;
end;

procedure TRekapMitraFrm.BtnDesign2Click(Sender: TObject);
begin
  ppDesigner1.ShowModal;
end;

procedure TRekapMitraFrm.ppTitleBand1BeforePrint(Sender: TObject);
begin
  ppNamaLaporan.Caption:='MUTASI STOK '+DMFrm.QJnsItem.FieldByName('JNS_BRG').AsString;
  ppPeriode.Caption:=vTglAwal.Text+' SD '+vTglAkhir.Text;
  DMFrm.QTime.Close;
  DMFrm.QTime.Open;
  ppUserCetak.Caption:=DMFrm.QTimeVUSER_CETAK.AsString;
end;

procedure TRekapMitraFrm.Label5Click(Sender: TObject);
begin
  if DMFrm.FontDialog1.Execute then
  begin
    wwDBGrid1.Font.Name:=DMFrm.FontDialog1.Font.Name;
    wwDBGrid1.Font.Size:=DMFrm.FontDialog1.Font.Size;
    wwDBGrid1.Font.Color:=DMFrm.FontDialog1.Font.Color;
    wwDBGrid1.Font.Style:=DMFrm.FontDialog1.Font.Style;
  end;
end;

procedure TRekapMitraFrm.LookItemEnter(Sender: TObject);
begin
  (sender as TwwDBLookupComboDlg).LookupTable.Open;
end;

procedure TRekapMitraFrm.ppDetailBand1BeforePrint(Sender: TObject);
begin
 ppLblNomer.Caption:=IntToStr(ppDBQBrowseDetail.RecordNo+1)
end;

procedure TRekapMitraFrm.QTransaksiBeforeOpen(DataSet: TDataSet);
begin
  QTransaksi.DeclareVariable('kd_transaksi', otString);
  QTransaksi.SQL.Text:='select a.* from '+cUserTabel+'vtransaksi a where a.kd_transaksi=:kd_transaksi';
  QTransaksi.SetVariable('kd_transaksi',vkode);
end;

procedure TRekapMitraFrm.vTglAwalChange(Sender: TObject);
begin
  vTglAkhir.DateTime:=EndOfTheMonth(vTglAwal.Date);
end;

procedure TRekapMitraFrm.FormShow(Sender: TObject);
begin
    vTglAwal.Date:=Trunc(DMFrm.QTimeJAM.AsDateTime);
    vTglAkhir.Date:=Trunc(DMFrm.QTimeJAM.AsDateTime);
    LapProduksiFrm:=Nil;
end;

procedure TRekapMitraFrm.QBrowseAfterScroll(DataSet: TDataSet);
begin
    LabelBanner.Caption:='Record ke '+IntToStr(QBrowse.RecNo)+' dari '+FormatFloat('#,#',QBrowse.RecordCount)+' Records';
end;

procedure TRekapMitraFrm.BitBtnPrint2Click(Sender: TObject);
begin
  QuickRep2.Preview;
end;

procedure TRekapMitraFrm.cbAdaTransaksiClick(Sender: TObject);
begin
//  QBrowse.Filtered:=cbAdaTransaksi.Checked;
end;

procedure TRekapMitraFrm.TitleBand1BeforePrint(
  Sender: TQRCustomBand; var PrintBand: Boolean);
begin
  QRLabel13.Caption:='Bulan  '+FormatDateTime('MMMM YYYY',VTglAkhir.Date);
  QRLabel14.Caption:='(Per : '+vTglAwal.Text+' s/d '+vTglAkhir.Text+')';
end;

procedure TRekapMitraFrm.QRBand5BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  DMFrm.QDateTime.Open;
//  QRLabel84.Caption:=FormatDateTime('dd mmmm yyyy',vTglAkhir.Date);
end;

procedure TRekapMitraFrm.QRBand6BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  DMFrm.QDateTime.Open;
end;

procedure TRekapMitraFrm.QRBand8BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
//  QRLabel55.Caption:=vTglAwal.Text+' s/d '+vTglAkhir.Text;
end;

procedure TRekapMitraFrm.QRBand2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRLabel50.Caption:='Bulan  '+FormatDateTime('MMMM YYYY',VTglAkhir.Date);
  QRLabel51.Caption:='(Per : '+vTglAwal.Text+' s/d '+vTglAkhir.Text+')';
end;

procedure TRekapMitraFrm.PageFooterBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  DMFrm.QDateTime.Open;
end;

procedure TRekapMitraFrm.QuickRep1AfterPreview(Sender: TObject);
begin
  LapProduksiFrm :=Nil;
end;

procedure TRekapMitraFrm.QuickRep1AfterPrint(Sender: TObject);
begin
  LapProduksiFrm :=Nil;
end;

procedure TRekapMitraFrm.wwDBGrid1UpdateFooter(Sender: TObject);
begin
{  QTotal.Close;
  QTotal.SetVariable('item',QBrowseNAMA_ITEM.AsString);
}
{  QTotal.Close;
  QTotal.Open;
  wwDBGrid1.ColumnByName('QTY1').FooterValue:=FormatFloat('0.00,0;(0.00,0);-',QTotalQTY1.AsFloat);
  wwDBGrid1.ColumnByName('QTY2').FooterValue:=FormatFloat('0,0;(0,0);-',QTotalQTY2.AsFloat);
  wwDBGrid1.ColumnByName('QTY3').FooterValue:=FormatFloat('0.00,0;(0.00,0);-',QTotalQTY3.AsFloat);
  wwDBGrid1.ColumnByName('QTY4').FooterValue:=FormatFloat('0,0;(0,0);-',QTotalQTY4.AsFloat);
  wwDBGrid1.ColumnByName('QTY5').FooterValue:=FormatFloat('0.00,0;(0.00,0);-',QTotalQTY5.AsFloat);
  wwDBGrid1.ColumnByName('QTY6').FooterValue:=FormatFloat('0,0;(0,0);-',QTotalQTY6.AsFloat);
  wwDBGrid1.ColumnByName('QTY7').FooterValue:=FormatFloat('0.00,0;(0.00,0);-',QTotalQTY7.AsFloat);
  wwDBGrid1.ColumnByName('QTY8').FooterValue:=FormatFloat('0,0;(0,0);-',QTotalQTY8.AsFloat);
}
end;

procedure TRekapMitraFrm.wwDBGrid2UpdateFooter(Sender: TObject);
var  vqty1 : real;
begin
  vqty1:=0;
  while not QBrowse.Eof do
  begin
    vqty1:=vqty1+QBrowseLUSI.AsFloat;
    QBrowse.Next;
  end;
  QBrowse.EnableControls;
end;

procedure TRekapMitraFrm.SummaryBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  DMFrm.QDateTime.Open;
end;

procedure TRekapMitraFrm.QuickRep1BeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
  LapProduksiFrm :=Nil;
end;

procedure TRekapMitraFrm.TabSheet3Show(Sender: TObject);
begin
  vTglAwal3.Date:=Trunc(DMFrm.QTimeJAM.AsDateTime);
  vTglAkhir3.Date:=Trunc(DMFrm.QTimeJAM.AsDateTime);
  QBrowse3.DisableControls;
  QBrowse3.Close;
  QBrowse3.SetVariable('ptgl', vTglAwal3.Date);
  QBrowse3.SetVariable('ptgl2', vTglAkhir3.Date);
  QBrowse3.SetVariable('myparam', 'order by no_nota');
  QBrowse3.Open;
  QBrowse3.EnableControls;
end;

procedure TRekapMitraFrm.BitBtn1Click(Sender: TObject);
var
  vqty1, vqty2 : real;
begin
  QBrowse2.DisableControls;
  QBrowse2.Close;
  QBrowse2.SetVariable('ptgl', vtglAwal2.Date);
  QBrowse2.SetVariable('ptgl2', vtglAkhir2.Date);
  QBrowse2.Open;
  QBrowse2.EnableControls;

  vqty1:=0;
  vqty2:=0;
  while not QBrowse2.Eof do
  begin
    vqty1:=vqty1+QBrowse2PAKAN.AsFloat;
    vqty2:=vqty2+QBrowse2LUSI.AsFloat;
    QBrowse2.Next;
  end;
  QBrowse2.EnableControls;
  LBanner2.Caption:='Data : '+FormatFloat('#,#',QBrowse2.RecordCount)+' Records';
  wwDBGrid2.ColumnByName('PAKAN').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty1);
  wwDBGrid2.ColumnByName('LUSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty2);
end;

procedure TRekapMitraFrm.vtglAwal2Change(Sender: TObject);
begin
  vTglAkhir2.DateTime:=EndOfTheMonth(vTglAwal2.Date);
end;

procedure TRekapMitraFrm.QBrowse2AfterScroll(DataSet: TDataSet);
begin
    LBanner2.Caption:='Record ke '+IntToStr(QBrowse2.RecNo)+' dari '+FormatFloat('#,#',QBrowse2.RecordCount)+' Records';
end;

procedure TRekapMitraFrm.TabSheet2Show(Sender: TObject);
begin
    vTglAwal2.Date:=Trunc(DMFrm.QTimeJAM.AsDateTime);
    vTglAkhir2.Date:=Trunc(DMFrm.QTimeJAM.AsDateTime);
    QBrowse2.DisableControls;
    QBrowse2.Close;
    QBrowse2.SetVariable('ptgl', vTglAwal2.Date);
    QBrowse2.SetVariable('ptgl2', vTglAkhir2.Date);
    QBrowse2.SetVariable('myparam', 'order by no_nota');
    QBrowse2.Open;
    QBrowse2.EnableControls;
end;

procedure TRekapMitraFrm.BitBtn2Click(Sender: TObject);
var
  vqty1, vqty2, vqty3, vqty4 : real;
begin
  QBrowse3.DisableControls;
  QBrowse3.Close;
  QBrowse3.SetVariable('ptgl', vtglAwal3.Date);
  QBrowse3.SetVariable('ptgl2', vtglAkhir3.Date);
  QBrowse3.Open;
  QBrowse3.EnableControls;

  vqty1:=0;
  vqty2:=0;
  vqty3:=0;
  vqty4:=0;
  while not QBrowse3.Eof do
  begin
    vqty1:=vqty1+QBrowse3QTY_PTG.AsFloat;
    vqty2:=vqty2+QBrowse3LS_TERIMA_PRODUKSI.AsFloat;
    vqty3:=vqty3+QBrowse3PK_TERIMA_PRODUKSI.AsFloat;
    vqty4:=vqty4+QBrowse3PK_TERIMA_PRODUKSI2.AsFloat;
    QBrowse3.Next;
  end;
  QBrowse3.EnableControls;
  LBanner3.Caption:='Data : '+FormatFloat('#,#',QBrowse3.RecordCount)+' Records';
  wwDBGrid3.ColumnByName('QTY_PTG').FooterValue:=FormatFloat('#,0.0;-#,0.0;-',vqty1);
  wwDBGrid3.ColumnByName('LS_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty2);
  wwDBGrid3.ColumnByName('PK_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty3);
  wwDBGrid3.ColumnByName('PK_TERIMA_PRODUKSI2').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty4);
end;

procedure TRekapMitraFrm.Label6Click(Sender: TObject);
begin
  if DMFrm.FontDialog1.Execute then
  begin
    wwDBGrid2.Font.Name:=DMFrm.FontDialog1.Font.Name;
    wwDBGrid2.Font.Size:=DMFrm.FontDialog1.Font.Size;
    wwDBGrid2.Font.Color:=DMFrm.FontDialog1.Font.Color;
    wwDBGrid2.Font.Style:=DMFrm.FontDialog1.Font.Style;
  end;
end;

procedure TRekapMitraFrm.wwDBSpinEdit1Change(Sender: TObject);
begin
  wwDBGrid2.RowHeightPercent:=Round(wwDBSpinLine2.Value);
end;

procedure TRekapMitraFrm.BtnExportClick(Sender: TObject);
begin
  DMFrm.SaveDialog1.DefaultExt:='XLK';
  DMFrm.SaveDialog1.Filter:='Excel files (*.XLK)|*.XLK';

  if PageControl2.TabIndex=0 then
  begin
    DMFrm.SaveDialog1.FileName:=TabSheet1.Caption+' Per  '+ vTglAwal.Text+' sd '+vTglAkhir.Text+'.xlk';
    wwDBGrid1.ExportOptions.TitleName:=TabSheet1.Caption+' Per '+vTglAwal.Text+' sd '+vTglAkhir.Text;
  end;

  if PageControl2.TabIndex=1 then
  begin
    DMFrm.SaveDialog1.FileName:=TabSheet2.Caption+' Per  '+ vTglAwal2.Text+' sd '+vTglAkhir2.Text+'.xlk';
    wwDBGrid2.ExportOptions.TitleName:=TabSheet2.Caption+' Per '+vTglAwal2.Text+' sd '+vTglAkhir2.Text;
  end;

  if PageControl2.TabIndex=2 then
  begin
    DMFrm.SaveDialog1.FileName:=TabSheet3.Caption+' Per  '+ vTglAwal3.Text+' sd '+vTglAkhir3.Text+'.xlk';
    wwDBGrid3.ExportOptions.TitleName:=TabSheet3.Caption+' Per '+vTglAwal3.Text+' sd '+vTglAkhir3.Text;
  end;

  if PageControl2.TabIndex=3 then
  begin
    DMFrm.SaveDialog1.FileName:=TabSheet4.Caption+' Per  '+ vTglAwal4.Text+' sd '+vTglAkhir4.Text+'.xlk';
    wwDBGrid4.ExportOptions.TitleName:=TabSheet4.Caption+' Per '+vTglAwal4.Text+' sd '+vTglAkhir4.Text;
  end;

  if PageControl2.TabIndex=4 then
  begin
    DMFrm.SaveDialog1.FileName:=TabSheet5.Caption+' Per  '+ vTglAwal5.Text+' sd '+vTglAkhir5.Text+'.xlk';
    wwDBGrid5.ExportOptions.TitleName:=TabSheet5.Caption+' Per '+vTglAwal5.Text+' sd '+vTglAkhir5.Text;
  end;

  if DMFrm.SaveDialog1.Execute then
  begin
    try
        if PageControl2.TabIndex=0 then
        begin
          wwDBGrid1.ExportOptions.FileName:=DMFrm.SaveDialog1.FileName;
          wwDBGrid1.ExportOptions.Save;
          ShowMessage('Simpan Sukses !');
        end;

        if PageControl2.TabIndex=1 then
        begin
          wwDBGrid2.ExportOptions.FileName:=DMFrm.SaveDialog1.FileName;
          wwDBGrid2.ExportOptions.Save;
          ShowMessage('Simpan Sukses !');
        end;

        if PageControl2.TabIndex=2 then
        begin
          wwDBGrid3.ExportOptions.FileName:=DMFrm.SaveDialog1.FileName;
          wwDBGrid3.ExportOptions.Save;
          ShowMessage('Simpan Sukses !');
        end;

        if PageControl2.TabIndex=3 then
        begin
          wwDBGrid4.ExportOptions.FileName:=DMFrm.SaveDialog1.FileName;
          wwDBGrid4.ExportOptions.Save;
          ShowMessage('Simpan Sukses !');
        end;

        if PageControl2.TabIndex=4 then
        begin
          wwDBGrid5.ExportOptions.FileName:=DMFrm.SaveDialog1.FileName;
          wwDBGrid5.ExportOptions.Save;
          ShowMessage('Simpan Sukses !');
        end;
    except
      ShowMessage('Simpan Gagal !');
    end;
  end;
end;

procedure TRekapMitraFrm.BtnFind2Click(Sender: TObject);
begin
  if not QBrowse2.QBEMode then
  begin
    wwDBGrid2.Options:=wwDBGrid2.Options-[dgRowSelect,dgAlwaysShowSelection];
    QBrowse2.QBEMode:=TRUE;
  end
  else
    QBrowse2.ClearQBE;
end;

procedure TRekapMitraFrm.SpeedButton2Click(Sender: TObject);
var t1, t2 : real;
begin
  if QBrowse2.QBEMode then
  begin
    QBrowse2.ExecuteQBE;
    wwDBGrid2.Options:=wwDBGrid2.Options+[dgRowSelect,dgAlwaysShowSelection];
    QBrowse2.Open;
    t1:=0;
    t2:=0;
    while not QBrowse2.Eof do
    begin
      t1:=t1+QBrowse2PAKAN.AsFloat;
      t2:=t2+QBrowse2LUSI.AsFloat;
      QBrowse2.Next;
    end;
    QBrowse2.EnableControls;
    LBanner2.Caption:='Data : '+FormatFloat('#,#',QBrowse2.RecordCount)+' Records';
    wwDBGrid2.ColumnByName('PAKAN').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',t1);
    wwDBGrid2.ColumnByName('LUSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',t2);
  end;
end;

procedure TRekapMitraFrm.TabSheet1Show(Sender: TObject);
begin
  QBrowse.DisableControls;
  QBrowse.Close;
  QBrowse.SetVariable('ptgl', vTglAwal.Date);
  QBrowse.SetVariable('ptgl2', vTglAkhir.Date);
  QBrowse.SetVariable('myparam', 'order by no_nota');
  QBrowse.Open;
  QBrowse.EnableControls;
end;

procedure TRekapMitraFrm.QBrowse3AfterScroll(DataSet: TDataSet);
begin
  LBanner3.Caption:='Record ke '+IntToStr(QBrowse3.RecNo)+' dari '+FormatFloat('#,#',QBrowse3.RecordCount)+' Records';
end;

procedure TRekapMitraFrm.SpeedButton1Click(Sender: TObject);
begin
  if not QBrowse3.QBEMode then
  begin
    wwDBGrid3.Options:=wwDBGrid3.Options-[dgRowSelect,dgAlwaysShowSelection];
    QBrowse3.QBEMode:=TRUE;
  end
  else
    QBrowse3.ClearQBE;
end;

procedure TRekapMitraFrm.SpeedButton3Click(Sender: TObject);
var t1, t2, t3, t4 : real;
begin
  if QBrowse3.QBEMode then
  begin
    QBrowse3.ExecuteQBE;
    wwDBGrid3.Options:=wwDBGrid3.Options+[dgRowSelect,dgAlwaysShowSelection];
    QBrowse3.Open;
    t1:=0;
    t2:=0;
    t3:=0;
    t4:=0;
    while not QBrowse3.Eof do
    begin
      t1:=t1+QBrowse3QTY_PTG.AsFloat;
      t2:=t2+QBrowse3LS_TERIMA_PRODUKSI.AsFloat;
      t3:=t3+QBrowse3PK_TERIMA_PRODUKSI.AsFloat;
      t4:=t4+QBrowse3PK_TERIMA_PRODUKSI2.AsFloat;
      QBrowse3.Next;
    end;
    QBrowse3.EnableControls;
    LBanner3.Caption:='Data : '+FormatFloat('#,#',QBrowse3.RecordCount)+' Records';
    wwDBGrid3.ColumnByName('QTY_PTG').FooterValue:=FormatFloat('#,0.00;-#,0.00;-',t1);
    wwDBGrid3.ColumnByName('LS_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',t2);
    wwDBGrid3.ColumnByName('PK_TERIMA_PRODUKSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',t3);
    wwDBGrid3.ColumnByName('PK_TERIMA_PRODUKSI2').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',t4);
  end;
end;

procedure TRekapMitraFrm.vTglAwal3Change(Sender: TObject);
begin
  vTglAkhir3.DateTime:=EndOfTheMonth(vTglAwal3.Date);
end;

procedure TRekapMitraFrm.vTglAwal4Change(Sender: TObject);
begin
    vTglAkhir4.DateTime:=EndOfTheMonth(vTglAwal4.Date);
end;

procedure TRekapMitraFrm.BitBtn3Click(Sender: TObject);
var
  vqty1 : real;
begin
     QBrowse4.DisableControls;
     QBrowse4.SetVariable('ptgl', vTglAwal4.Date);
     QBrowse4.SetVariable('ptgl2', vTglAkhir4.Date);
     QBrowse4.Close;
     QBrowse4.Open;
     QBrowse4.EnableControls;
     vqty1:=0;
     while not QBrowse4.Eof do
     begin
       vqty1:=vqty1+QBrowse4QTY_LUSI1.AsFloat;
       QBrowse4.Next;
     end;
     QBrowse4.EnableControls;
     LBanner4.Caption:='Data : '+FormatFloat('#,#',QBrowse4.RecordCount)+' Records';
     wwDBGrid4.ColumnByName('QTY_LUSI1').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty1);
end;

procedure TRekapMitraFrm.SpeedButton4Click(Sender: TObject);
begin
  if not QBrowse4.QBEMode then
  begin
    wwDBGrid4.Options:=wwDBGrid4.Options-[dgRowSelect,dgAlwaysShowSelection];
    QBrowse4.QBEMode:=TRUE;
  end
  else
    QBrowse4.ClearQBE;
end;

procedure TRekapMitraFrm.SpeedButton5Click(Sender: TObject);
var t1 : real;
begin
  if QBrowse4.QBEMode then
  begin
    QBrowse4.ExecuteQBE;
    wwDBGrid4.Options:=wwDBGrid4.Options+[dgRowSelect,dgAlwaysShowSelection];
    QBrowse4.Open;
    t1:=0;
    while not QBrowse4.Eof do
    begin
      t1:=t1+QBrowse4QTY_LUSI1.AsFloat;
      QBrowse4.Next;
    end;
    QBrowse4.EnableControls;
    LBanner4.Caption:='Data : '+FormatFloat('#,#',QBrowse4.RecordCount)+' Records';
    wwDBGrid4.ColumnByName('QTY_LUSI1').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',t1);
  end;

end;

procedure TRekapMitraFrm.QBrowse4AfterScroll(DataSet: TDataSet);
begin
  LBanner4.Caption:='Record ke '+IntToStr(QBrowse4.RecNo)+' dari '+FormatFloat('#,#',QBrowse4.RecordCount)+' Records';
end;

procedure TRekapMitraFrm.vTglAwal5Change(Sender: TObject);
begin
  vTglAkhir5.DateTime:=EndOfTheMonth(vTglAwal5.Date);
end;

procedure TRekapMitraFrm.BitBtn4Click(Sender: TObject);
var
  vqty1, vqty2 : real;
begin
     QBrowse5.DisableControls;
     QBrowse5.SetVariable('ptgl', vTglAwal5.Date);
     QBrowse5.SetVariable('ptgl2', vTglAkhir5.Date);
     QBrowse5.Close;
     QBrowse5.Open;
     QBrowse5.EnableControls;
     vqty1:=0;
     vqty2:=0;
     while not QBrowse5.Eof do
     begin
       vqty1:=vqty1+QBrowse5LUSI.AsFloat;
       vqty2:=vqty2+QBrowse5PAKAN.AsFloat;
       QBrowse5.Next;
     end;
     QBrowse5.EnableControls;
     LBanner5.Caption:='Data : '+FormatFloat('#,#',QBrowse5.RecordCount)+' Records';
     wwDBGrid5.ColumnByName('LUSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty1);
     wwDBGrid5.ColumnByName('PAKAN').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',vqty2);
end;

procedure TRekapMitraFrm.SpeedButton6Click(Sender: TObject);
begin
  if not QBrowse5.QBEMode then
  begin
    wwDBGrid5.Options:=wwDBGrid5.Options-[dgRowSelect,dgAlwaysShowSelection];
    QBrowse5.QBEMode:=TRUE;
  end
  else
    QBrowse5.ClearQBE;
end;

procedure TRekapMitraFrm.SpeedButton7Click(Sender: TObject);
var t1, t2 : real;
begin
  if QBrowse5.QBEMode then
  begin
    QBrowse5.ExecuteQBE;
    wwDBGrid5.Options:=wwDBGrid5.Options+[dgRowSelect,dgAlwaysShowSelection];
    QBrowse5.Open;
    t1:=0;
    t2:=0;
    while not QBrowse5.Eof do
    begin
      t1:=t1+QBrowse5LUSI.AsFloat;
      t2:=t2+QBrowse5PAKAN.AsFloat;
      QBrowse5.Next;
    end;
    QBrowse5.EnableControls;
    LBanner5.Caption:='Data : '+FormatFloat('#,#',QBrowse5.RecordCount)+' Records';
    wwDBGrid5.ColumnByName('LUSI').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',t1);
    wwDBGrid5.ColumnByName('PAKAN').FooterValue:=FormatFloat('#,0.000;-#,0.000;-',t2);
  end;
end;

end.
