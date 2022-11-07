{ DU-E autenticar servidor com certificado digital e enviar xml

DRI Soluções 17/07/2018

Vamos precisar dos seguintes componentes: TIdHTTP (Indy Clients) e TIdSSLIOHandlerSocketOpenSSL (Indy I/O Handlers), TMemo(mmResponse) e TButton (Standard)
O certificado digital deve estar convertido para .pem ( https://www.sslshopper.com/ssl-converter.html )

-> Configurando TIdHTTP
AllowCookies := True
HandleRedirects := True
HTTPOptions
hoKeepOrigProtocol :=True
IOHandler := IdSSLIOHandlerSocketOpenSSL1
ProxyParams
BasicAuthentication :=True
ProxyPassword <senha>
ProxyPort <porta>
ProxyServer <servidor>
ProxyUsername <usuário>
Request
Accept := 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
AcceptChasSet := 'pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3'
AcceptLanguage :='pt-BR'
CharSet :='UTF-8'
ContentType := 'application/json'
UserAgent := 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Acoo Browser; GTB5; Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1) ; Maxthon; InfoPath.1; .NET CLR 3.5.30729; .NET CLR 3.0.30618);'

-> Configurando idSSLIOHandlerSocketOpenSSL1
SSLOptions
CertFile := 'cert.pem'
KeyFile := 'cert.pem'
Method := 'sslvTLSv1'
Mode := 'sslmClient'
RootCertFile := 'cert.pem'

Variáveis globais
Token, XCSRF, URL : string;  }

unit Commex.Controller.DUE;

interface

uses
  Classes, System.SysUtils, System.DateUtils, System.StrUtils, IniFiles, IdIOHandler, IdIOHandlerSocket,
  IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdHTTP, Vcl.Dialogs, Xml.xmldom, Xml.XMLIntf, Xml.XMLDoc;
  
const
  DUE_VERSAO = '0.0.1a';
  DUE_NAMESPACE = 'https://portalunico.siscomex.gov.br/due/';

type
   TDUE = class;
   TTipoDoctoFiscal = (tdNFe, tdNFFormulario, tdSemNF);
   TResponse = (rAutorizada, rError);
   TUtils = class;
   TIdeDespacho = class;
   TIdeEmbarque = class;
   TIdeRetorno = class;
   TFormaExportacao = class;
   TSituacaoEspecial = class;
   TViaTransporte = class;
   TObservacao = class;
   TDeclarante = class;
   TCambio = class;
   TReferenciaCarga = class;
   TProd = class;
   TDetCollection = class;        // Collection
   TDetCollectionItem = class;    // CollectionItem
   TEnquad = class;
   TNFRefCollection = class;                // Collection
   TNFRefCollectionItem = class;       // CollectionItem
   TNotaCollection = class;       // Collection
   TNotaCollectionItem = class;   // CollectionItem
   TErroCollection = class;
   TErroCollectionItem = class;
   TErroInfo = class;
   TDetalhe = class;

   { Classe Base }
   TDUE = class(TPersistent)
   private
     FIdeDespacho : TIdeDespacho;
     FIdeEmbarque : TIdeEmbarque;
     FIdeRetorno  : TIdeRetorno;
     FFormaExportacao : TFormaExportacao;
     FSituacaoEspecial: TSituacaoEspecial;
     FViaTransporte : TViaTransporte;
     FObservacao : TObservacao;
     FDeclarante : TDeclarante;
     FCambio : TCambio;
     FReferenciaCarga: TReferenciaCarga;
     FDet : TDetCollection;
     FIdeErro : TErroCollection;
     FUtils : TUtils;
     FTipoDoctoFiscal: TTipoDoctoFiscal;
   public
     constructor Create;
     destructor Destroy; override;
     procedure Assign(Source: TPersistent); override;
   published
     property Utils: TUtils read FUtils write FUtils;
     property TipoDoctoFiscal: TTipoDoctoFiscal read FTipoDoctoFiscal write FTipoDoctoFiscal;
     property IdeDespacho: TIdeDespacho read FIdeDespacho write FIdeDespacho;
     property IdeEmbarque: TIdeEmbarque read FIdeEmbarque write FIdeEmbarque;
     property IdeRetorno: TIdeRetorno read FIdeRetorno write FIdeRetorno;
     property FormaExportacao: TFormaExportacao read FFormaExportacao write FFormaExportacao;
     property SituacaoEspecial: TSituacaoEspecial read FSituacaoEspecial write FSituacaoEspecial;
     property ViaTransporte: TViaTransporte read FViaTransporte write FViaTransporte;
     property Observacao: TObservacao read FObservacao write FObservacao;
     property Declarante: TDeclarante read FDeclarante write FDeclarante;
     property Cambio: TCambio read FCambio write FCambio;
     property ReferenciaCarga: TReferenciaCarga read FReferenciaCarga write FReferenciaCarga;
     property Det: TDetCollection read FDet write FDet;
     property IdeErro: TErroCollection read FIdeErro write FIdeErro;
   end;

   TIdeDespacho = class(TPersistent)
   private
     FnRFB : Integer;        // Obr (7)
     FnRecinto : Integer;    // Obr (7) ou (14) para despacho fora de Recinto
     FnTipo  : Integer;        // Obr (3)
   public
     procedure Assign(Source: TPersistent); override;
   published
     property nRFB : Integer read FnRFB write FnRFB;
     property nRecinto : Integer read FnRecinto write FnRecinto;
     property nTipo : Integer read FnTipo write FnTipo;
   end;

   TIdeEmbarque = class(TPersistent)
   private
     FnRFB : Integer;        // Obr (7)
     FnRecinto : Integer;    // Obr (7)
     FnTipo : Integer;         // Obr (3)
   public
     procedure Assign(Source: TPersistent); override;
   published
     property nRFB: Integer read FnRFB write FnRFB;
     property nRecinto: Integer read FnRecinto write FnRecinto;
     property nTipo: Integer read FnTipo write FnTipo;
   end;

   TIdeRetorno = class(TPersistent)
   private
     FMensagem : String;
     FNroDUE : String;
     FNroRUC : String;
     FChaveAcesso : String;
     FData : String;
     FCPF : String;
   public
     procedure Assign(Source: TPersistent); override;
   published
     property Mensagem: String read FMensagem write FMensagem;
     property NroDUE: String read FNroDUE write FNroDUE;
     property NroRUC: String read FNroRUC write FNroRUC;
     property ChaveAcesso: String read FChaveAcesso write FChaveAcesso;
     property Data: String read FData write FData;
     property CPF: String read FCPF write FCPF;
   end;

   TFormaExportacao = class(TPersistent)
   private
     FTipo : String;     // Obr (3)
     FnCodigo : Integer; // Opt (4)
   public
     procedure Assign(Source: TPersistent); override;
   published
     property Tipo: String read FTipo write FTipo;
     property nCodigo: Integer read FnCodigo write FnCodigo;
   end;

   TSituacaoEspecial = class(TPersistent)
   private
     FTipo : String;      // Obr (3)
     FnCodigo: Integer;   // Opt (4)
   public
     procedure Assign(Source: TPersistent); override;
   published
     property Tipo: String read FTipo write FTipo;
     property nCodigo: Integer read FnCodigo write FnCodigo;
   end;

   TViaTransporte = class(TPersistent)
   private
     FTipo : String;     // Obr (3)
     FnCodigo : Integer; // Opt (4)
   public
     procedure Assign(Source: TPersistent); override;
   published
     property Tipo: String read FTipo write FTipo;
     property nCodigo: Integer read FnCodigo write FnCodigo;
   end;

   TObservacao = class(TPersistent)
   private
     FTipoDesc: String;  // Opt (3) AAI, DEF, ABC
     FDesc : String;     // Opt (1000) AAI, (600) DEF e Numeric(17) ABC
   public
     procedure Assign(Source: TPersistent); override;
   published
     property TipoDescricao: String read FTipoDesc write FTipoDesc;
     property Desc: String read FDesc write FDesc;
   end;

   TDeclarante = class(TPersistent)
   private
     FDocto : String;         // Obr (14) CNPJ e (11) CPF
     FNome : String;          // Opt (100)
     FContato : String;       // Opt (14) telefone e (100) Email
     FTipoContato: String;    // Opt (2)
   public
     procedure Assign(Source: TPersistent); override;
   published
     property Docto: String read FDocto write FDocto;
     property Nome: String read FNome write FNome;
     property Contato: String read FContato write FContato;
     property TipoContato: String read FTipoContato write FTipoContato;
   end;

   TCambio = class(TPersistent)
   private
     FMoeda : String;         // Obr (3)
   public
     procedure Assign(Source: TPersistent); override;
   published
     property Moeda: String read FMoeda write FMoeda;
   end;

   TReferenciaCarga = class(TPersistent)
   private
     FRUC : String;         // Opt (35)
   public
     procedure Assign(Source: TPersistent); override;
   published
     property RUC: String read FRUC write FRUC;
   end;

   { Notas - Collection }
   TNotaCollection = class(TCollection)
   private
     FnChave : String;       // Obr (44)
     FnTipo : Integer;       // Obr (3)
    function GetItem(Index: Integer): TNotaCollectionItem;
    procedure SetItem(Index: Integer; const Value: TNotaCollectionItem);
   public
     constructor Create(AOwner: TProd);
     procedure Assign(Source: TPersistent); override;
     function Add: TNotaCollectionItem;
     property Items[Index: Integer]: TNotaCollectionItem read GetItem write SetItem;
   published
     property nChave: String read FnChave write FnChave;
     property nTipo: Integer read FnTipo write FnTipo;
   end;

   TNotaCollectionItem = class(TCollectionItem)
   private
     FnChave : String;         // Obr (44)
     FnTipo  : String;         // Obr (3)
     FnDocto : String;         // Obr (14) CNPJ (11) CPF
   public
     constructor Create(Collection: TCollection); override;
     destructor Destroy; override;
     procedure Assign(Source: TPersistent); override;
   published
     property nChave: String read FnChave write FnChave;
     property nTipo: String read FnTipo write FnTipo;
     property nDocto: String read FnDocto write FnDocto;
   end;

   { Detalhes - Collection }
   TDetCollection = class(TCollection)
   private
    function GetItem(Index: Integer): TDetCollectionItem;
    procedure SetItem(Index: Integer; const Value: TDetCollectionItem);

   public
     constructor Create(AOwner: TDUE);
     function Add: TDetCollectionItem;
     property Items[Index: Integer]: TDetCollectionItem read GetItem write SetItem; default;
   end;

   TDetCollectionItem = class(TCollectionItem)
   private
     FProd : TProd;                         // OBR
   public
     constructor Create(Collection: TCollection); override;
     destructor Destroy; override;
//     procedure Assig(Source: TPersistent); override;
   published
     property Prod: TProd read FProd write FProd;
   end;

   { Erros - Collection }
   TErroCollection = class(TCollection)
   private
     FMensagem : String;
     FCodigo : String;
     FCampo : String;
     FTag: String;
     FData: String;
     FStatus: String;
     FGravidade: String;
     FInfo: TErroInfo;
    function GetItem(Index: Integer): TErroCollectionItem;
    procedure SetItem(Index: Integer; const Value: TErroCollectionItem);
   public
     constructor Create(AOwner: TDUE);
     procedure Assign(Source: TPersistent); override;
     function Add: TErroCollectionItem;
     property Items[Index: Integer]: TErroCollectionItem read GetItem write SetItem; default;
   published
     property Mensagem: String read FMensagem write FMensagem;
     property Codigo: String read FCodigo write FCodigo;
     property Campo: String read FCampo write FCampo;
     property Tag: String read FTag write FTag;
     property Data: String read FData write FData;
     property Status: String read FStatus write FStatus;
     property Gravidade: String read FGravidade write FGravidade;
     property Info: TErroInfo read FInfo write FInfo;
   end;

   TErroCollectionItem = class(TCollectionItem)
   private
     FMensagem : String;
     FCodigo : String;
     FTag : String;
     FData : String;
     FStatus : String;
     FGravidade : String;
	 FDetalhe: TDetalhe;
   public
     constructor Create(Collection: TCollection); override;
     destructor Destroy; override;
     procedure Assign(Source: TPersistent); override;
   published
     property Mensagem: String read FMensagem write FMensagem;
     property Codigo: String read FCodigo write FCodigo;
     property Tag: String read FTag write FTag;
     property Data: String read FData write FData;
     property Status: String read FStatus write FStatus;
     property Gravidade: String read FGravidade write FGravidade;
	 property Detalhe: TDetalhe read FDetalhe write FDetalhe;
   end;

   TErroInfo = class(TPersistent)
   private
     FAmbiente : String;
     FMNemonico : String;
     FSistema : String;
     FUrl : String;
     FUsuario : String;
     FVisao : String;
   public
     procedure Assign(Source: TPersistent); override;
   published
     property Ambiente: String read FAmbiente write FAmbiente;
     property MNemonico: String read FMNemonico write FMNemonico;
     property Sistema: String read FSistema write FSistema;
     property Url: String read FUrl write FUrl;
     property Usuario: String read FUsuario write FUsuario;
     property Visao: String read FVisao write FVisao;
   end;
   
   TDetalhe = class(TPersistent)
   private
     FMensagem : String;
     FCodigo : String;
     FTag : String;
     FData : String;
     FStatus : String;
     FGravidade : String;
   public
     constructor Create(AOwner: TErroCollectionItem);
     destructor Destroy; override;
     procedure Assign(Source: TPersistent); override;
   published
     property Mensagem: String read FMensagem write FMensagem;
     property Codigo: String read FCodigo write FCodigo;
     property Tag: String read FTag write FTag;
     property Data: String read FData write FData;
     property Status: String read FStatus write FStatus;
     property Gravidade: String read FGravidade write FGravidade;
   end;
   
   TProd = class(TPersistent)
   private
     FnItem  : Integer;          // Obr (9) o máximo é 999  Nro. do item da DUE
     FvVMLE : Currency;          // Obr (17,2)
     FvVMCV : Currency;          // Obr (17,2)
     FDescCompl : String;        // Opt (600)
     FvTotalMerc: Currency;      // Opt (17,2)
     FsNCM: String;
     FsTpAtributo: String;
     FsAtributo: String;
     FsTpNCM: String;
     FvFinanc:  Currency;        // Opt (17,2)
     FsPais  : String;           // Obr (2)
     FqEst: Double;              // Obr (19,5)
     FvPesoLiqTotal: Double;     // Obr (19,5)
     FnCatAjust: Integer;        // Opt (3) 149 - Comissão do Agente
     FpComissaoAgente: Currency; // Opt (5,2)
     FuEst: String;              // Opt (3)
     FuCom: String;              // Opt (3)
     FqUnidEst: Double;          // Opt (19,5)
     FqUnidCom: Double;          // Opt (19,5)
     FsCondVenda: String;        // Obr (3)
     FEnquad : TEnquad;          // OBR somente um
     FNFRef : TNFRefCollection;  // Invoiceline OBR
     FNotas : TNotaCollection;  // Invoice Obr
     procedure SetNFRef(const Value: TNFRefCollection);
    procedure SetNotas(const Value: TNotaCollection);
   public
     constructor Create(AOwner: TDetCollectionItem);
     destructor Destroy; override;
     procedure Assign(Source: TPersistent); override;
   published
     property nItem: Integer read FnItem write FnItem;
     property vVMLE: Currency read FvVMLE write FvVMLE;
     property vVMCV: Currency read FvVMCV write FvVMCV;
     property DescCompl: String read FDescCompl write FDescCompl;
     property vTotalMerc: Currency read FvTotalMerc write FvTotalMerc;
     property sNCM: String read FsNCM write FsNCM;
     property sTpNCM: String read FsTpNCM write FsTpNCM;
     property sTpAtributo: String read FsTpAtributo write FsTpAtributo;
     property sAtributo: String read FsAtributo write FsAtributo;
     property vFinanc:  Currency read FvFinanc write FvFinanc;
     property sPais: String read FsPais write FsPais;
     property qEst: Double read FqEst write FqEst;
     property vPesoLiqTotal: Double read FvPesoLiqTotal write FvPesoLiqTotal;
     property nCatAjust: Integer read FnCatAjust write FnCatAjust;
     property pComissaoAgente: Currency read FpComissaoAgente write FpComissaoAgente;
     property uEst: String read FuEst write FuEst;
     property uCom: String read FuCom write FuCom;
     property qUnidEst: Double read FqUnidEst write FqUnidEst;
     property qUnidCom: Double read FqUnidCom write FqUnidCom;
     property sCondVenda: String read FsCondVenda write FsCondVenda;
     property Enquad: TEnquad read FEnquad write FEnquad;
     property NFRef: TNFRefCollection read FNFRef write SetNFRef;
     property Notas: TNotaCollection read FNotas write SetNotas;
   end;

   TEnquad = class(TPersistent)
   private
     FnEnquad01: Integer;             // Obr (5)
     FnEnquad02: Integer;             // Opt (5)
     FnEnquad03: Integer;             // Opt (5)
     FnEnquad04: Integer;             // Opt (5)
   public
     procedure Assign(Source: TPersistent); override;
   published
     property nEnquad01: Integer read FnEnquad01 write FnEnquad01;
     property nEnquad02: Integer read FnEnquad02 write FnEnquad02;
     property nEnquad03: Integer read FnEnquad03 write FnEnquad03;
     property nEnquad04: Integer read FnEnquad04 write FnEnquad04;
   end;

   { Notas Referenciadas dos itens }
   TNFRefCollection = class(TCollection)
   private
     FnItemNF  : Integer;
    function GetItem(Index: Integer): TNFRefCollectionItem;
    procedure SetItem(Index: Integer; const Value: TNFRefCollectionItem);                  // Obr (3) Relaciona o Item da DUE com o item da NF de exportação
   public
     constructor Create(AOwner: TProd);
     destructor Destroy; override;
     procedure Assign(Source: TPersistent); override;
     function Add: TNFRefCollectionItem;
     property Items[Index: Integer]: TNFRefCollectionItem read GetItem write SetItem;
   published
     property nItemNF: Integer read FnItemNF write FnItemNF;
   end;

   TNFRefCollectionItem = class(TCollectionItem)
   private
     FnItem : Integer;   // Obr (3) Nro. do item na NF de Remessa ou Complementar
     FnChave : String;   // Obr (44) Chave de acesso da NF de Remessa ou Complementar
     FqUnidEst: Double;  // Obr (19,5) Quantidade Estatistica na unid. de medida da NF Remessa
   public
     procedure Assign(Source: TPersistent); override;
   published
     property nItem: Integer read FnItem write FnItem;
     property nChave: String read FnChave write FnChave;
     property qUnidEst: Double read FqUnidEst write FqUnidEst;
   end;

   { Utilitários da DU-E }
   TUtils = class
     IdHTTPAutentica: TIdHTTP;
     IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;
   private
     FToken : String;
     FXCSRF : String;
   public
     FErroRegrasNegocios: String;
     function GeraXML(var aXML: String; aDUE: TDUE): Boolean;
     function CriarArquivoXML(aNomeArquivo, aXML, aPath: String; aConteudoEhUTF8: Boolean) :Boolean;
     function Valida(aDUE : TDUE): Boolean;
     function EnviarArquivoXML(aPathFile: String): String;
     function EnviarXML: String;
     function Autenticar(var Response: String): Boolean;
     function LeResposta(aDUE: TDUE ;const Response: String): Boolean;
     function PathWithDelim( const APath: String): String;
     function XmlEhUTF8(const AXML: String): Boolean;
     function EstaVazio(const AValue: String): Boolean;
     function NaoEstaVazio(const AValue: String): Boolean;
     procedure WriteToTXT( const ArqTXT : String; const ABinaryString : AnsiString;
       const AppendIfExists : Boolean = True; const AddLineBreak : Boolean = True;
       const ForceDirectory : Boolean = False);
     procedure GerarException(const Msg: String; E: Exception = nil);
     property Token: String read FToken write FToken;
     property XCSRF: String read FXCSRF write FXCSRF;
     property ErroRegrasNegocios: String read FErroRegrasNegocios;
   end;

   EDUEException = class(Exception)
   public
     constructor Create(const Msg: String);
   end;

   EDUEExceptionNoPrivateKey = class(EDUEException);

implementation

{ TIdeDespacho }

procedure TIdeDespacho.Assign(Source: TPersistent);
begin
  if Source is TIdeDespacho then
  begin
    nRFB     := TIdeDespacho(Source).nRFB;
    nRecinto := TIdeDespacho(Source).nRecinto;
    nTipo      := TIdeDespacho(Source).nTipo;
  end
  else
    inherited;
end;

{ TDUE }

procedure TDUE.Assign(Source: TPersistent);
begin
  if Source is TDUE then
  begin
    IdeDespacho.Assign(TDUE(Source).IdeDespacho);
    IdeEmbarque.Assign(TDUE(Source).IdeEmbarque);
    IdeRetorno.Assign(TDUE(Source).IdeRetorno);
    FormaExportacao.Assign(TDUE(Source).FormaExportacao);
    SituacaoEspecial.Assign(TDUE(Source).SituacaoEspecial);
    ViaTransporte.Assign(TDUE(Source).ViaTransporte);
    Declarante.Assign(TDUE(Source).Declarante);
    Cambio.Assign(TDUE(Source).Cambio);
    ReferenciaCarga.Assign(TDUE(Source).ReferenciaCarga);
    Det.Assign(TDUE(Source).Det);
    IdeErro.Assign(TDUE(Source).IdeErro);
  end
  else
    inherited;
end;

constructor TDUE.Create;
begin
  FUtils						:= TUtils.Create;
  FIdeDespacho      := TIdeDespacho.Create;
  FIdeEmbarque      := TIdeEmbarque.Create;
  FIdeRetorno       := TIdeRetorno.Create;
  FFormaExportacao  := TFormaExportacao.Create;
  FSituacaoEspecial := TSituacaoEspecial.Create;
  FViaTransporte    := TViaTransporte.Create;
  FObservacao				:= TObservacao.Create;
  FDeclarante       := TDeclarante.Create;
  FCambio           := TCambio.Create;
  FReferenciaCarga  := TReferenciaCarga.Create;
  FDet              := TDetCollection.Create(Self);
  FIdeErro          := TErroCollection.Create(Self);

  FIdeDespacho.nTipo              := 281;
  FIdeEmbarque.nTipo              := 281;
  FFormaExportacao.nCodigo        := 1001;
  FFormaExportacao.Tipo           := 'CUS';
  FObservacao.TipoDescricao  			:= 'AAI';
  FSituacaoEspecial.nCodigo       := 0;
  FSituacaoEspecial.Tipo          := '';
  FViaTransporte.Tipo             := 'TRA';
  FViaTransporte.nCodigo          := 4001;
end;

destructor TDUE.Destroy;
begin
  FUtils.Free;
  FIdeDespacho.Free;
  FIdeEmbarque.Free;
  FIdeRetorno.Free;
  FFormaExportacao.Free;
  FSituacaoEspecial.Free;
  FViaTransporte.Free;
  FObservacao.Free;
  FDeclarante.Free;
  FCambio.Free;
  FReferenciaCarga.Free;
  FDet.Free;
  FIdeErro.Free;
  inherited Destroy;
end;

{ TNotas }

function TNotaCollection.Add: TNotaCollectionItem;
begin
  Result := TNotaCollectionItem(inherited Add);
  //// Result.create;
end;

procedure TNotaCollection.Assign(Source: TPersistent);
begin
  if Source is TNotaCollection then
  begin
    nChave := TNotaCollection(Source).nChave;
    nTipo  := TNotaCollection(Source).nTipo;
  end
  else
    inherited;
end;

constructor TNotaCollection.Create(AOwner: TProd);
begin
  inherited Create(TNotaCollectionItem);
  nTipo := 388;
end;

function TNotaCollection.GetItem(Index: Integer): TNotaCollectionItem;
begin
  Result := TNotaCollectionItem(inherited GetItem(Index));
//    Result := Items[Index] as TFilme;
end;

procedure TNotaCollection.SetItem(Index: Integer; const Value: TNotaCollectionItem);
begin
  inherited SetItem(Index, Value);
end;

{ TDetCollection }

function TDetCollection.Add: TDetCollectionItem;
begin
  Result := TDetCollectionItem(inherited Add);
end;

constructor TDetCollection.Create(AOwner: TDUE);
begin
  inherited Create(TDetCollectionItem);
end;

function TDetCollection.GetItem(Index: Integer): TDetCollectionItem;
begin
  Result  := TDetCollectionItem(inherited GetItem(Index));
end;

procedure TDetCollection.SetItem(Index: Integer;
  const Value: TDetCollectionItem);
begin
  inherited SetItem(Index, Value);
end;

{ TIdeEmbarque }

procedure TIdeEmbarque.Assign(Source: TPersistent);
begin
  if Source is TIdeEmbarque then
  begin
    nRFB     := TIdeEmbarque(Source).nRFB;
    nRecinto := TIdeEmbarque(Source).nRecinto;
    nTipo      := TIdeEmbarque(Source).nTipo;
  end
  else
    inherited;
end;

{ TFormaExportacao }

procedure TFormaExportacao.Assign(Source: TPersistent);
begin
  if Source is TFormaExportacao then
  begin
    Tipo          := TFormaExportacao(Source).Tipo;
    nCodigo       := TFormaExportacao(Source).nCodigo;
  end
  else
    inherited;
end;

{ TViaTransporte }

procedure TViaTransporte.Assign(Source: TPersistent);
begin
  if Source is TViaTransporte then
  begin
    Tipo    := TViaTransporte(Source).Tipo;
    nCodigo := TViaTransporte(Source).nCodigo;
  end
  else
    inherited;
end;

{ TObservacao }

procedure TObservacao.Assign(Source: TPersistent);
begin
  if Source is TFormaExportacao then
  begin
    TipoDescricao := TObservacao(Source).TipoDescricao;
    Desc          := TObservacao(Source).Desc;
  end
  else
    inherited;
end;

{ TDeclarante }

procedure TDeclarante.Assign(Source: TPersistent);
begin
  if Source is TDeclarante then
  begin
    Docto       := TDeclarante(Source).Docto;
    Nome        := TDeclarante(Source).Nome;
    Contato     := TDeclarante(Source).Contato;
    TipoContato := TDeclarante(Source).TipoContato;
  end
  else
    inherited;
end;

{ TCambio }

procedure TCambio.Assign(Source: TPersistent);
begin
  if Source is TCambio then
  begin
    Moeda := TCambio(Source).Moeda;
  end
  else
    inherited;
end;

{ TReferenciaCarga }

procedure TReferenciaCarga.Assign(Source: TPersistent);
begin
  if Source is TReferenciaCarga then
  begin
    RUC := TReferenciaCarga(Source).RUC;
  end
  else
    inherited;
end;

{ TNotaCollectionItem }

procedure TNotaCollectionItem.Assign(Source: TPersistent);
begin
  if Source is TNotaCollectionItem then
  begin
    nChave   := TNotaCollectionItem(Source).nChave;
    nTipo    := TNotaCollectionItem(Source).nTipo;
    nDocto   := TNotaCollectionItem(Source).nDocto;
  end
  else
    inherited;
end;

constructor TNotaCollectionItem.Create(Collection: TCollection);
begin
  inherited;
//  FProd := TProd.Create(Self);
end;

destructor TNotaCollectionItem.Destroy;
begin
//
  inherited;
end;

{ TDetCollectionItem }

constructor TDetCollectionItem.Create(Collection: TCollection);
begin
  inherited;
  FProd := TProd.Create(Self);
end;

destructor TDetCollectionItem.Destroy;
begin
  FProd.Free;
  inherited;
end;

{ TProd }

procedure TProd.Assign(Source: TPersistent);
begin
  if Source is TProd then
  begin
     nItem        	:= TProd(Source).nItem;
     vVMLE        	:= TProd(Source).vVMLE;
     vVMCV        	:= TProd(Source).vVMCV;
     DescCompl    	:= TProd(Source).DescCompl;
     vTotalMerc   	:= TProd(Source).vTotalMerc;
     sNCM         	:= TProd(Source).sNCM;
     sTpNCM       	:= TProd(Source).sTpNCM;
     sTpAtributo  	:= TProd(Source).sTpAtributo;
     sAtributo    	:= TProd(Source).sAtributo;
     vFinanc     	:= TProd(Source).vFinanc;
     sPais       	:= TProd(Source).sPais;
     qEst         	:= TProd(Source).qEst;
     vPesoLiqTotal	:= TProd(Source).vPesoLiqTotal;
     nCatAjust    	:= TProd(Source).nCatAjust;
     pComissaoAgente  := TProd(Source).pComissaoAgente;
     uEst         	:= TProd(Source).uEst;
     uCom         	:= TProd(Source).uCom;
     qUnidEst     	:= TProd(Source).qUnidEst;
     qUnidCom     	:= TProd(Source).qUnidCom;
     sCondVenda   	:= TProd(Source).sCondVenda;
  end
  else
    inherited;
end;

constructor TProd.Create(AOwner: TDetCollectionItem);
begin
  inherited Create;
  FnCatAjust  := 149; // 149 - Comissão do Agente
  FsTpNCM     := 'HS';
  FNFRef      := TNFRefCollection.Create(Self);
  FNotas      := TNotaCollection.Create(Self);
  FEnquad     := TEnquad.Create;
end;

destructor TProd.Destroy;
begin
  FEnquad.Free;
  FNFRef.Free;
  FNotas.Free;
  inherited;
end;


procedure TProd.SetNFRef(const Value: TNFRefCollection);
begin
  FNFRef.Assign(Value);
end;

procedure TProd.SetNotas(const Value: TNotaCollection);
begin
  FNotas.Assign(Value);
end;

{ TNFRefCollectionItem }

procedure TNFRefCollectionItem.Assign(Source: TPersistent);
begin
  if Source is TNFRefCollectionItem then
  begin
    nItem := TNFRefCollectionItem(Source).nItem;
    nChave := TNFRefCollectionItem(Source).nChave;
    qUnidEst := TNFRefCollectionItem(Source).qUnidEst;
  end
  else
    inherited;
end;

{ TNFRefCollection }

function TNFRefCollection.Add: TNFRefCollectionItem;
begin
  Result := TNFRefCollectionItem(inherited Add);
end;

procedure TNFRefCollection.Assign(Source: TPersistent);
begin
  if Source is TNFRefCollection then
    nItemNF := TNFRefCollection(Source).nItemNF;
  inherited;
end;

constructor TNFRefCollection.Create(AOwner: TProd);
begin
 inherited Create(TNFRefCollectionItem);
end;

destructor TNFRefCollection.Destroy;
begin

  inherited;
end;

function TNFRefCollection.GetItem(Index: Integer): TNFRefCollectionItem;
begin
  Result := TNFRefCollectionItem(inherited GetItem(Index));
end;

procedure TNFRefCollection.SetItem(Index: Integer;
  const Value: TNFRefCollectionItem);
begin
  inherited SetItem(Index, Value);
end;

{ TEnquad }

procedure TEnquad.Assign(Source: TPersistent);
begin
  if Source is TEnquad then
  begin
    nEnquad01 := TEnquad(Source).nEnquad01;
    nEnquad02 := TEnquad(Source).nEnquad02;
    nEnquad03 := TEnquad(Source).nEnquad03;
    nEnquad04 := TEnquad(Source).nEnquad04;
  end
  else
    inherited;
end;

{ TSituacaoEspecial }

procedure TSituacaoEspecial.Assign(Source: TPersistent);
begin
  if Source is TSituacaoEspecial then
  begin
    Tipo  := TSituacaoEspecial(Source).Tipo;
    nCodigo := TSituacaoEspecial(Source).nCodigo;
  end
  else
    inherited;
end;

{ TUtils }

function TUtils.Autenticar(var Response: String): Boolean;
var
  URL: String;
  Request : TStringList;
  idSSLIOHandlerSocketOpenSSL1 : TIdSSLIOHandlerSocketOpenSSL;
begin
	{ O certificado digital deve estar convertido para .pem ( https://www.sslshopper.com/ssl-converter.html )
  	Variáveis globais Token, XCSRF, URL : string; }

  Result := True;
  URL := 'https://val.portalunico.siscomex.gov.br/portal/api/autenticar'; { url para ambiente de validação }
  // URL := 'https://portalunico.siscomex.gov.br/portal/api/autenticar'; { url para ambiente de produção }

  Request := TStringList.Create;

  { Configuração do Soquete Indy }
  idSSLIOHandlerSocketOpenSSL1 := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
  idSSLIOHandlerSocketOpenSSL1.SSLOptions.CertFile	:= 'C:\Apcohd\Certificado\14832129_out.pem';
  idSSLIOHandlerSocketOpenSSL1.SSLOptions.KeyFile		:= 'C:\Apcohd\Certificado\14832129_out.pem';
  idSSLIOHandlerSocketOpenSSL1.SSLOptions.Method		:= sslvTLSv1;
  idSSLIOHandlerSocketOpenSSL1.SSLOptions.Mode      := sslmClient;
	idSSLIOHandlerSocketOpenSSL1.SSLOptions.RootCertFile:= 'C:\Apcohd\Certificado\14832129_out.pem';

  { Configuração do HTTP Indy }
  IdHTTPAutentica := TIdHTTP.Create(nil);
  IdHTTPAutentica.AllowCookies		:= True;
  IdHTTPAutentica.HandleRedirects	:= True;
  IdHTTPAutentica.HTTPOptions := [hoKeepOrigProtocol];
  IdHTTPAutentica.IOHandler		:= idSSLIOHandlerSocketOpenSSL1;
  IdHTTPAutentica.ProxyParams.BasicAuthentication	:= False;
	IdHTTPAutentica.ProxyParams.ProxyPassword := '';
  IdHTTPAutentica.ProxyParams.ProxyPort			:= 0;
  IdHTTPAutentica.ProxyParams.ProxyServer		:= '';
  IdHTTPAutentica.ProxyParams.ProxyUsername	:= '';
  IdHTTPAutentica.Request.Accept					:= 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8';
  IdHTTPAutentica.Request.AcceptCharSet 	:= 'pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3';
  IdHTTPAutentica.Request.AcceptLanguage 	:= 'pt-BR';
  IdHTTPAutentica.Request.CharSet					:= 'UTF-8';
  IdHTTPAutentica.Request.ContentType			:= 'application/json';
  IdHTTPAutentica.Request.UserAgent 			:= 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Acoo Browser; GTB5; Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1) ; Maxthon; InfoPath.1; .NET CLR 3.5.30729; .NET CLR 3.0.30618);';
  IdHTTPAutentica.Request.CustomHeaders.Clear;
  IdHTTPAutentica.Request.CustomHeaders.AddValue('Role-Type', 'IMPEXP');
  IdHTTPAutentica.Response.Clear;
  IdHTTPAutentica.Response.ContentType 		:= 'application/json';
  IdHTTPAutentica.Response.CharSet 				:= 'UTF-8';
  IdHTTPAutentica.Response.KeepAlive 			:= True;

  try
    try
      Response 	:= IdHTTPAutentica.Post(URL, Request);
      Token 		:= IdHTTPAutentica.Response.RawHeaders.Values['Set-Token']; {recebendo o valor de header Set-Token}
      XCSRF 		:= IdHTTPAutentica.Response.RawHeaders.Values['X-CSRF-Token']; {recebendo o valor de header X-CSRF-Token}
    except
      on E: EIdHTTPProtocolException do
      begin
        Result := False;
        Response  := (E.ErrorMessage);
      end;
    end;
  finally
    Request.Free;
    //IdHTTPAutentica.Free;
  end;
end;

function TUtils.CriarArquivoXML(aNomeArquivo, aXML, aPath: String;
  aConteudoEhUTF8: Boolean): Boolean;
var
  SoNome, SoPath: String;
begin
  Result := False;
  try
    SoNome	:= ExtractFileName(aNomeArquivo);
    if SoNome = EmptyStr then
      raise EDUEException.Create('Nome do arquivo não informado');

    SoPath := ExtractFilePath(aPath);
    if SoPath = EmptyStr then
      raise EDUEException.Create('Pasta destino do arquivo não informado');

    SoPath	:= PathWithDelim(SoPath);

    if not DirectoryExists(SoPath) then
      ForceDirectories(SoPath);

    aNomeArquivo := SoPath + SoNome;

    WriteToTXT(aNomeArquivo, aXML, False, False);
    Result := True;
  except
    on E: Exception do
      GerarException('Erro ao Gerar o Arquivo!', E);
  end;
end;

function TUtils.EnviarArquivoXML(aPathFile: String): String;
var
  Response, URL : String;
  Arquivo : TMemoryStream;
  IResponse: TResponse;
begin
  Arquivo := TMemoryStream.Create;

  try
    if not FileExists(aPathFile) then
    begin
      Response := 'Erro: Arquivo não existente na pasta ' + ExtractFilePath(aPathFile);
      Exit;
    end;

    Arquivo.LoadFromFile(aPathFile);

    URL := 'https://val.portalunico.siscomex.gov.br/due/api/ext/due'; { url para ambiente de validação }
    //URL := 'https://portalunico.siscomex.gov.br/due/api/ext/due'; { url para ambiente de produção }

    IdHTTPAutentica.Request.ContentType := 'Application/xml';
    IdHTTPAutentica.Request.CustomHeaders.Clear;
    IdHTTPAutentica.Request.CustomHeaders.AddValue('Authorization', Token); { valor pego na autenticação }
    IdHTTPAutentica.Request.CustomHeaders.AddValue('X-CSRF-Token', XCSRF); { valor pego na autenticação }

    try
      Response 	:= '';
      Response	:= IdHTTPAutentica.Post(URL, Arquivo); { enviando arquivo }
    except
      on E:EIdHTTPProtocolException do
      begin
        Response  := '';
        Response  := e.ErrorMessage;
      end;
    end;
  finally
    Result := Response;
    Arquivo.Free;
  end;
end;

function TUtils.EnviarXML: String;
var
  Response, URL, PathFile: String;
  Arquivo : TMemoryStream;
Begin
  Arquivo := TMemoryStream.Create;

  try
    PathFile	:= Trim(ExtractFilePath(ParamStr(0)) + '\1234.xml');
    if not FileExists(PathFile) then
    begin
      Response := 'Erro: Arquivo não existente na pasta ' + ExtractFilePath(PathFile);
      Exit
    end;

    Arquivo.LoadFromFile(PathFile);

    URL := 'https://val.portalunico.siscomex.gov.br/due/api/ext/due'; { url para ambiente de validação }
    //URL := 'https://portalunico.siscomex.gov.br/due/api/ext/due'; { url para ambiente de produção }

    IdHTTPAutentica.Request.ContentType := 'Application/xml';
    IdHTTPAutentica.Request.CustomHeaders.Clear;
    IdHTTPAutentica.Request.CustomHeaders.AddValue('Authorization', Token); { valor pego na autenticação }
    IdHTTPAutentica.Request.CustomHeaders.AddValue('X-CSRF-Token', XCSRF); { valor pego na autenticação }

    try
      Response 	:= '';
      Response	:= IdHTTPAutentica.Post(URL, Arquivo); { enviando arquivo }
    except
      on E:EIdHTTPProtocolException do
      begin
        Response := '';
        Response  := e.ErrorMessage;
      end;
    end;
  finally
    Result := Response;
    Arquivo.Free();
  end;
end;

function TUtils.EstaVazio(const AValue: String): Boolean;
begin
  Result	:= (AValue = '');
end;

procedure TUtils.GerarException(const Msg: String; E: Exception);
var
  Tratado: Boolean;
  MsgErro: String;
begin
  MsgErro := Msg;
  if Assigned(E) then
    MsgErro := MsgErro + sLineBreak + E.Message;

  Tratado := False;
//  FazerLog('ERRO: ' + MsgErro, Tratado);

  // MsgErro já está na String Nativa da IDE... por isso deve usar "CreateDef"
  if not Tratado then
    raise EDUEException.Create(MsgErro);
end;

function TUtils.GeraXML(var aXML: String; aDUE: TDUE): Boolean;
var
  I, J, K : Integer;
begin
  FormatSettings.DecimalSeparator := '.';

  try
    aXML  := EmptyStr;
    { Cabeçalho do XML }
    aXML  := '<?xml version="1.0" encoding="UTF-8"?>' + sLineBreak;

    { Declaration }
    aXML  := aXML + '<Declaration xsi:schemaLocation='+QuotedStr('urn:wco:datamodel:WCO:GoodsDeclaration:1 GoodsDeclaration_1p0_DUE.xsd ') + ' xmlns= ' + QuotedStr('urn:wco:datamodel:WCO:GoodsDeclaration:1') + ' xmlns:xsi=' + QuotedStr('http://www.w3.org/2001/XMLSchema-instance')+ ' xmlns:ds=' + QuotedStr('urn:wco:datamodel:WCO:GoodsDeclaration_DS:1') + '>'  + sLineBreak;

    { Com NF-e }
    try
      { DeclarationNFE }
      aXML  := aXML + '  <DeclarationNFe>' + sLineBreak;

        { DeclarationOffice Local de Despacho }
        aXML  := aXML + '    <DeclarationOffice>' + sLineBreak;
          { Código da Unidade da Receita Federal de Despacho RFB }
          aXML  := aXML + '      <ID listID=' + QuotedStr('token')+ '>' + IntToStr(aDUE.IdeDespacho.nRFB) + '</ID>'  + sLineBreak;
          { Dados do Local do despacho sendo fora ou dentro de recinto alfandegado }
          aXML  := aXML + '      <Warehouse>' + sLineBreak;
            { Código do Recinto Alfandegado de Despacho ou , quando despacho fora de recinto, o CNPJ do responsável pelo local }
            aXML  := aXML + '        <ID>' + IntToStr(aDUE.IdeDespacho.nRecinto) + '</ID>'  + sLineBreak;
            { Flag que identifica se o despacho será realizado dentro ou fora de recinto alfandegado }
            { 19 - Fora de Recinto Alfandegado - Domiciliar,
              22 - Fora de Recinto Alfandegado - Não Domiciliar,
              281 - Recinto alfandegado }
            aXML  := aXML + '        <TypeCode>' + IntToStr(aDUE.IdeDespacho.nTipo) + '</TypeCode>'  + sLineBreak;
          aXML  := aXML + '      </Warehouse>' + sLineBreak;
        aXML  := aXML + '    </DeclarationOffice>' + sLineBreak;

        { Forma de Exportação }
        aXML  := aXML + '    <AdditionalInformation>' + sLineBreak;
          { StatementCode Código da Forma de Exportação }
          { 1001 - Forma Exportação/Por conta própria,
            1002 - Forma Exportação/Por conta e ordem de terceiros,
            1003 - Forma Exportação/Por operador de remessa postal ou expressa, }
          aXML  := aXML + '      <StatementCode>' + IntToStr(aDUE.FormaExportacao.nCodigo) + '</StatementCode>'  + sLineBreak;
          { Código que indica qual dado será informado }
          aXML  := aXML + '      <StatementTypeCode>' + aDUE.FormaExportacao.Tipo + '</StatementTypeCode>'  + sLineBreak;
        aXML  := aXML + '    </AdditionalInformation>' + sLineBreak;

        { Situação Especial }
        if aDUE.SituacaoEspecial.nCodigo > 0 then
        begin
          aXML  := aXML + '    <AdditionalInformation>' + sLineBreak;
            { StatementCode Código da Situação Especial }
            { 2001 - Situação Especial/Despacho a posteriori,
              2002 - Situação Especial/Embarque antecipado,
              2003 - Situação Especial/Sem saída da mercadoria do país,}
            aXML  := aXML + '      <StatementCode>' + IntToStr(aDUE.SituacaoEspecial.nCodigo) + '</StatementCode>'  + sLineBreak;

            { Código que indica qual dado será informado }
            aXML  := aXML + '      <StatementTypeCode>' + aDUE.SituacaoEspecial.Tipo + '</StatementTypeCode>'  + sLineBreak;

          aXML  := aXML + '    </AdditionalInformation>' + sLineBreak;
        end;

        { AdditionalInformation Caso Especial de Transporte }
        aXML  := aXML + '    <AdditionalInformation>' + sLineBreak;
          { StatementCode Código do Caso Especial de Transporte }
          { 4001 - Caso Especial Transporte/Meios próprios,
            4002 - Caso Especial Transporte/Dutos,
            4003 - Caso Especial Transporte/Linhas de transmissão,
            4004 - Caso Especial Transporte/Em mãos,
            4005 - Caso Especial Transporte/Por reboque }
          aXML  := aXML + '      <StatementCode>' + IntToStr(aDUE.ViaTransporte.nCodigo) + '</StatementCode>'  + sLineBreak;

          { Código que indica qual dado será informado }
          aXML  := aXML + '      <StatementTypeCode>' + aDUE.ViaTransporte.Tipo + '</StatementTypeCode>' + sLineBreak;

        aXML  := aXML + '    </AdditionalInformation>' + sLineBreak;

        { AdditionalInformation Observações Gerais }
        aXML  := aXML + '    <AdditionalInformation>' + sLineBreak;
          { StatementTypeCode Código que indica qual dado será informado }
          aXML  := aXML + '      <StatementTypeCode>' + aDUE.Observacao.TipoDescricao + '</StatementTypeCode>'  + sLineBreak;
          { Observações Gerais }
          aXML  := aXML + '      <StatementDescription>' + aDUE.Observacao.Desc + '</StatementDescription>'  + sLineBreak;
        aXML  := aXML + '    </AdditionalInformation>' + sLineBreak;

        { Moeda de Negociação }
        aXML  := aXML + '    <CurrencyExchange>' + sLineBreak;
          aXML  := aXML + '      <CurrencyTypeCode>' + aDUE.Cambio.Moeda + '</CurrencyTypeCode>'  + sLineBreak;
        aXML  := aXML + '    </CurrencyExchange>' + sLineBreak;

        { Declarant CPF/CNPJ do Declarante da DU-E (Declarante é diferente de Exportador) }
        aXML  := aXML + '    <Declarant>' + sLineBreak;
          aXML  := aXML + '      <ID schemeID=' + QuotedStr('token') + '>' + aDUE.Declarante.Docto + '</ID>'  + sLineBreak;
        aXML  := aXML + '    </Declarant>' + sLineBreak;

        { ExitOffice Dados do local de embarque }
        aXML  := aXML + '    <ExitOffice>' + sLineBreak;
          { Código da Unidade da Receita Federal de Embarque RFB }
          aXML  := aXML + '      <ID>' + IntToStr(aDUE.IdeEmbarque.nRFB) + '</ID>'  + sLineBreak;

          { Dados do Local do embarque sendo fora ou dentro de recinto alfandegado }
          aXML  := aXML + '      <Warehouse>' + sLineBreak;
              { Código do Recinto Alfandegado de Embarque ou , quando embarque fora de recinto, o CNPJ do responsável pelo local }
              aXML  := aXML + '        <ID>' + IntToStr(aDUE.IdeEmbarque.nRecinto) + '</ID>'  + sLineBreak;

              { Flag que indica que o embarque será realizado dentro de recinto alfandegado }
              { 19 - Fora de Recinto Alfandegado - Domiciliar,
                22 - Fora de Recinto Alfandegado - Não Domiciliar,
                281 - Recinto alfandegado }
              aXML  := aXML + '        <TypeCode>' + IntToStr(aDUE.IdeEmbarque.nTipo) + '</TypeCode>'  + sLineBreak;
          aXML  := aXML + '      </Warehouse>' + sLineBreak;
        aXML  := aXML + '    </ExitOffice>' + sLineBreak;

        // --> LOOP DOS ITENS
        for I := 0 to aDUE.Det.Count -1 do
        begin
          { Dados dos itens da DU-E agrupado por Nota Fiscal de Exportação }
          aXML  := aXML + '    <GoodsShipment>' + sLineBreak;
            { Itens da DU-E relacionados ao itens da Nota Fiscal de Exportação }
            aXML  := aXML + '      <GovernmentAgencyGoodsItem>' + sLineBreak;
              { Valor da mercadoria no local do embarque Num(17,2) }
              aXML  := aXML + '        <CustomsValueAmount languageID=' + QuotedStr('') + '>' + FormatCurr('################0.00', aDUE.Det.Items[I].Prod.vVMLE ) + '</CustomsValueAmount>'  + sLineBreak;
              { Número do Item da DU-E }
              aXML  := aXML + '        <SequenceNumeric>' + IntToStr(aDUE.Det.Items[I].Prod.nItem) + '</SequenceNumeric>'  + sLineBreak;

              { País de destino da mercadoria }
              aXML  := aXML + '        <Destination>' + sLineBreak;
                { Código do país qual a mercadoria foi enviada }
                aXML  := aXML + '          <CountryCode>' + aDUE.Det.Items[I].Prod.sPais + '</CountryCode>'  + sLineBreak;

                { Quantidade da mercadoria, na unidade de medida estatística, enviada ao país Num(19,5) }
                aXML  := aXML + '          <GoodsMeasure>' + sLineBreak;
                  aXML  := aXML + '            <TariffQuantity>' + FormatFloat('##################0.00000', aDUE.Det.Items[I].Prod.qUnidEst) + '</TariffQuantity>'  + sLineBreak;
                aXML  := aXML + '          </GoodsMeasure>' + sLineBreak;
              aXML  := aXML + '        </Destination>' + sLineBreak;

              { Commodity }
              aXML  := aXML + '        <Commodity>' + sLineBreak;
                { Descrição complementar da mercadoria 600 }
                aXML  := aXML + '          <Description>' + aDUE.Det.Items[I].Prod.DescCompl + '</Description>'  + sLineBreak;
                { Valor da mercadoria na condição de venda Num(17,2) }
                aXML  := aXML + '          <ValueAmount schemeID=' + QuotedStr('token')+'>' + FormatCurr('################0.00', aDUE.Det.Items[I].Prod.vVMCV) + '</ValueAmount>'  + sLineBreak;
                { Valor total da mercadoria Num(17,2) OPC }
                // <InvoiceBRLvalueAmount>154.45</InvoiceBRLvalueAmount>

                { Classificação da NCM }
                aXML  := aXML + '          <Classification>' + sLineBreak;
                  { NCM }
                  aXML  := aXML + '            <ID schemeID=' + QuotedStr('token')+'>' + StringReplace(aDUE.Det.Items[I].Prod.sNCM,'.','',[rfReplaceAll,rfIgnoreCase]) + '</ID>'  + sLineBreak;
                  aXML  := aXML + '            <IdentificationTypeCode>' + aDUE.Det.Items[I].Prod.sTpNCM + '</IdentificationTypeCode>'  + sLineBreak;
                aXML  := aXML + '          </Classification>' + sLineBreak;

                if aDUE.Det.Items[I].Prod.sTpAtributo <> '' then
                begin
                  aXML  := aXML + '          <ProductCharacteristics>' + sLineBreak;
                    { Atributos da NCM }
                    aXML  := aXML + '            <TypeCode>' + aDUE.Det.Items[I].Prod.sTpAtributo + '</TypeCode>'  + sLineBreak;
                    aXML  := aXML + '            <Description>' + aDUE.Det.Items[I].Prod.sAtributo + '</Description>'  + sLineBreak;
                  aXML  := aXML + '          </ProductCharacteristics>' + sLineBreak;
                end;

                { Item da Nota Fiscal }
                aXML  := aXML + '          <InvoiceLine>' + sLineBreak;
                  { Número do item da Nota Fiscal de Exportação máx. 999
                    Relaciona o item da DUE com o item da nota fiscal de exportação. Os números do item de nota iniciam em 1. }
                  aXML  := aXML + '            <SequenceNumeric>' + IntToStr(aDUE.Det.Items[I].Prod.NFRef.nItemNF) + '</SequenceNumeric>'  + sLineBreak;

                  // --> LOOP DAS NF-e REFERENCIADAS DO ITEM
                  for J := 0 to aDUE.Det.Items[I].Prod.NFRef.Count -1 do
                  begin
                    { tem da Nota Fiscal de Remessa ou Complementar }
                    aXML  := aXML + '            <ReferencedInvoiceLine>' + sLineBreak;
                      { Número do item da Nota Fiscal de Remessa ou Complementar }
                      aXML  := aXML + '              <SequenceNumeric>' + IntToStr(aDUE.Det.Items[I].Prod.NFRef.Items[J].nItem) + '</SequenceNumeric>'  + sLineBreak;
                      { Chave de Acesso da Nota Fiscal de Remessa ou Complementar }
                      aXML  := aXML + '              <InvoiceIdentificationID schemeID=' + QuotedStr('token') +'>' + aDUE.Det.Items[I].Prod.NFRef.Items[J].nChave + '</InvoiceIdentificationID>'  + sLineBreak;
                      { Quantidade na unidade de medida estatística a ser consumida da Nota Fisca de Remessa  Num(19,5) }
                      aXML  := aXML + '              <GoodsMeasure>' + sLineBreak;
                        aXML  := aXML + '                <TariffQuantity unitCode=' + QuotedStr('') + '>' + FormatFloat('##################0.00000', aDUE.Det.Items[I].Prod.NFRef.Items[J].qUnidEst) + '</TariffQuantity>'  + sLineBreak;
                      aXML  := aXML + '              </GoodsMeasure>' + sLineBreak;
                    aXML  := aXML + '            </ReferencedInvoiceLine>' + sLineBreak;
                  end;
                aXML  := aXML + '          </InvoiceLine>' + sLineBreak;

                { Atributos da NCM OPCIONAL }
                // ProductCharacteristitcs
                  {  }
                  // TypeCode ATT_1905 /TypeCode
                  {  }
                  // Description TRUE /Description
                // ProductCharacteristitcs end
              aXML  := aXML + '        </Commodity>' + sLineBreak;

              { Peso líquido total do item da DU-E em KG Num(19,5) }
              aXML  := aXML + '        <GoodsMeasure>' + sLineBreak;
                aXML  := aXML + '          <NetNetWeightMeasure>' + FormatFloat('##################0.00000', aDUE.Det.Items[I].Prod.vPesoLiqTotal) + '</NetNetWeightMeasure>'  + sLineBreak;
              aXML  := aXML + '        </GoodsMeasure>' + sLineBreak;

              { Enquadramentos }
              if aDUE.Det.Items[I].Prod.Enquad.nEnquad01 > 0 then
              begin
                aXML  := aXML + '        <GovernmentProcedure>' + sLineBreak;
                  { Enquadramento do item da DU-E }
                  aXML  := aXML + '          <CurrentCode>' + IntToStr(aDUE.Det.Items[I].Prod.Enquad.nEnquad01) + '</CurrentCode>'  + sLineBreak;
                aXML  := aXML + '        </GovernmentProcedure>' + sLineBreak;
              end;

              if aDUE.Det.Items[I].Prod.Enquad.nEnquad02 > 0 then
              begin
                aXML  := aXML + '        <GovernmentProcedure>' + sLineBreak;
                  { Enquadramento do item da DU-E }
                  aXML  := aXML + '          <CurrentCode>' + IntToStr(aDUE.Det.Items[I].Prod.Enquad.nEnquad02) + '</CurrentCode>'  + sLineBreak;
                aXML  := aXML + '        </GovernmentProcedure>' + sLineBreak;
              end;

              if aDUE.Det.Items[I].Prod.Enquad.nEnquad03 > 0 then
              begin
                aXML  := aXML + '        <GovernmentProcedure>' + sLineBreak;
                  { Enquadramento do item da DU-E }
                  aXML  := aXML + '          <CurrentCode>' + IntToStr(aDUE.Det.Items[I].Prod.Enquad.nEnquad03) + '</CurrentCode>'  + sLineBreak;
                aXML  := aXML + '        </GovernmentProcedure>' + sLineBreak;
              end;

              if aDUE.Det.Items[I].Prod.Enquad.nEnquad04 > 0 then
              begin
                aXML  := aXML + '        <GovernmentProcedure>' + sLineBreak;
                  { Enquadramento do item da DU-E }
                  aXML  := aXML + '          <CurrentCode>' + IntToStr(aDUE.Det.Items[I].Prod.Enquad.nEnquad04) + '</CurrentCode>'  + sLineBreak;
                aXML  := aXML + '        </GovernmentProcedure>' + sLineBreak;
              end;

            aXML  := aXML + '      </GovernmentAgencyGoodsItem>' + sLineBreak;

            { Dados da Nota Fiscal }
            aXML  := aXML + '      <Invoice>' + sLineBreak;
              { Chave de acesso da Nota Fiscal de Exportação }
              aXML  := aXML + '        <ID schemeID='+QuotedStr('token')+'>' + aDUE.Det.Items[I].Prod.Notas.nChave + '</ID>'  + sLineBreak;
              { Código que define que é uma nota fiscal }
              aXML  := aXML + '        <TypeCode>' + IntToStr(aDUE.Det.Items[I].Prod.Notas.nTipo) + '</TypeCode>'  + sLineBreak;

              // --> LOOP DAS NF REFERENCIADAS
              { Nota Fiscal Referenciada de Remessa ou Complementar }
              for K := 0 to aDUE.Det.Items[I].Prod.Notas.Count -1 do
              begin
                aXML  := aXML + '        <ReferencedInvoice>' + sLineBreak;
                  { Chave de acesso da Nota Fiscal Eletrônica }
                  aXML  := aXML + '          <ID schemeID='+QuotedStr('token') +'>' + aDUE.Det.Items[I].Prod.Notas.Items[K].nChave + '</ID>'  + sLineBreak;
                  { Código que define se é uma Nota Fiscal de Remessa ou Complementar }
                  aXML  := aXML + '          <TypeCode>' + aDUE.Det.Items[I].Prod.Notas.Items[K].nTipo + '</TypeCode>'  + sLineBreak;
                  // REM - Nota Fiscal de Remessa, COM - Nota Fiscal Complementar

                  { CPF/CNPJ do emissor da Nota Fiscal }
                  aXML  := aXML + '          <Submitter>' + sLineBreak;
                    aXML  := aXML + '            <ID schemeID=' + QuotedStr('token') + '>' + aDUE.Det.Items[I].Prod.Notas.Items[K].nDocto + '</ID>'  + sLineBreak;
                  aXML  := aXML + '          </Submitter>' + sLineBreak;
                aXML  := aXML + '        </ReferencedInvoice>' + sLineBreak;
              end;
            aXML  := aXML + '      </Invoice>' + sLineBreak;

            { Condição de Venda }
            aXML  := aXML + '      <TradeTerms>' + sLineBreak;
            aXML  := aXML + '        <ConditionCode>' + aDUE.Det.Items[I].Prod.sCondVenda + '</ConditionCode>'  + sLineBreak;
            aXML  := aXML + '      </TradeTerms>' + sLineBreak;
          aXML  := aXML + '    </GoodsShipment>' + sLineBreak;
        end;

        { UCR Código RUC }
        if not (aDUE.ReferenciaCarga.RUC = EmptyStr) then
        begin
          aXML  := aXML + '    <UCR>' + sLineBreak;
            aXML  := aXML + '      <TraderAssignedReferenceID schemeID=' + QuotedStr('token')+ '>' + aDUE.ReferenciaCarga.RUC + '</TraderAssignedReferenceID>'  + sLineBreak;
          aXML  := aXML + '    </UCR>' + sLineBreak;
        end;
      aXML  := aXML + '  </DeclarationNFe>' + sLineBreak;
    finally
      FormatSettings.DecimalSeparator := ',';
    end;

    aXML  := aXML + '</Declaration>';

    Result := True;
  except
    Result := False;
  end;
end;

function TUtils.LeResposta(aDUE: TDUE; const Response: String): Boolean;
var
  XMLDocto : IXMLDocument;
  I : Integer;
  nodeMessage, nodeInfo, nodeDetail : IXMLNode;
begin
  Result := True;
  XMLDocto := TXMLDocument.Create(nil);
  try
    try
      XMLDocto.LoadFromXML(Response);
          XMLDocto.Active := True;
      if Assigned(XMLDocto.ChildNodes.FindNode('pucomexReturn')) then
      begin
		nodeMessage := XMLDocto.ChildNodes.FindNode('pucomexReturn');
        if Assigned(nodeMessage.ChildNodes.FindNode('message')) then
          aDUE.IdeRetorno.Mensagem     := nodeMessage.ChildValues['message'];
        if Assigned(nodeMessage.ChildNodes.FindNode('due')) then
          aDUE.IdeRetorno.NroDUE       := nodeMessage.ChildValues['due'];
        if Assigned(nodeMessage.ChildNodes.FindNode('ruc')) then
          aDUE.IdeRetorno.NroRUC       := nodeMessage.ChildValues['ruc'];
        if Assigned(nodeMessage.ChildNodes.FindNode('chaveDeAcesso')) then
          aDUE.IdeRetorno.ChaveAcesso  := nodeMessage.ChildValues['chaveDeAcesso'];
        if Assigned(nodeMessage.ChildNodes.FindNode('date')) then
          aDUE.IdeRetorno.Data         := nodeMessage.ChildValues['date'];
        if Assigned(nodeMessage.ChildNodes.FindNode('cpf')) then
          aDUE.IdeRetorno.CPF          := nodeMessage.ChildValues['cpf'];
      end;
	  if Assigned(XMLDocto.ChildNodes.FindNode('error')) then
      begin
        nodeMessage := XMLDocto.ChildNodes.FindNode('error');
        if Assigned(nodeMessage.ChildNodes.FindNode('message')) then
          aDUE.IdeErro.Mensagem := nodeMessage.ChildValues['message'];
        if Assigned(nodeMessage.ChildNodes.FindNode('code')) then
          aDUE.IdeErro.Codigo       := nodeMessage.ChildValues['code'];
        if Assigned(nodeMessage.ChildNodes.FindNode('field')) then
          aDUE.IdeErro.Campo        := nodeMessage.ChildValues['field'];
        if Assigned(nodeMessage.ChildNodes.FindNode('tag')) then
          aDUE.IdeErro.Tag          := nodeMessage.ChildValues['tag'];
        if Assigned(nodeMessage.ChildNodes.FindNode('date')) then
          aDUE.IdeErro.Data         := nodeMessage.ChildValues['date'];
        if Assigned(nodeMessage.ChildNodes.FindNode('status')) then
          aDUE.IdeErro.Status       := nodeMessage.ChildValues['status'];
        if Assigned(nodeMessage.ChildNodes.FindNode('severity')) then
          aDUE.IdeErro.Gravidade:= nodeMessage.ChildValues['severity'];
 
        { Se houver node info }
        if Assigned(nodeMessage.ChildNodes.FindNode('info')) then
        begin
		  nodeInfo := nodeMessage.ChildNodes.FindNode('info');
 
          if Assigned(nodeInfo.ChildNodes.FindNode('ambiente')) then
            aDUE.IdeErro.Info.Ambiente  :=  nodeInfo.ChildValues['ambiente'];
          if Assigned(nodeInfo.ChildNodes.FindNode('mnemonico')) then
            aDUE.IdeErro.Info.MNemonico := nodeInfo.ChildValues['mnemonico'];
          if Assigned(nodeInfo.ChildNodes.FindNode('sistema')) then
            aDUE.IdeErro.Info.Sistema       := nodeInfo.ChildValues['sistema'];
          if Assigned(nodeInfo.ChildNodes.FindNode('url')) then
            aDUE.IdeErro.Info.Url               := nodeInfo.ChildValues['url'];
          if Assigned(nodeInfo.ChildNodes.FindNode('usuario')) then
            aDUE.IdeErro.Info.Usuario       := nodeInfo.ChildValues['usuario'];
          if Assigned(nodeInfo.ChildNodes.FindNode('visao')) then
            aDUE.IdeErro.Info.Visao         := nodeInfo.ChildValues['visao'];
        end;
		
		{ Se houver node detail }
        if Assigned(nodeMessage.ChildNodes.FindNode('detail')) then
        begin
          nodeDetail := nodeMessage.ChildNodes.FindNode('detail');
          for I := 0 to nodeDetail.ChildNodes.Count -1 do
          begin
            aDUE.IdeErro.Add;
            if Assigned(nodeDetail.ChildNodes[I].ChildNodes.FindNode('message')) then
              aDUE.IdeErro.Items[I].Detalhe.Mensagem    := nodeDetail.ChildNodes[I].ChildValues['message'];
            if Assigned(nodeDetail.ChildNodes[I].ChildNodes.FindNode('code')) then
              aDUE.IdeErro.Items[I].Detalhe.Codigo      := nodeDetail.ChildNodes[I].ChildValues['code'];
            if Assigned(nodeDetail.ChildNodes[I].ChildNodes.FindNode('tag')) then
              aDUE.IdeErro.Items[I].Detalhe.Tag := nodeDetail.ChildNodes[I].ChildValues['tag'];
            if Assigned(nodeDetail.ChildNodes[I].ChildNodes.FindNode('date')) then
              aDUE.IdeErro.Items[I].Detalhe.Data            := nodeDetail.ChildNodes[I].ChildValues['date'];
            if Assigned(nodeDetail.ChildNodes[I].ChildNodes.FindNode('status')) then
              aDUE.IdeErro.Items[I].Detalhe.Status      := nodeDetail.ChildNodes[I].ChildValues['status'];
            if Assigned(nodeDetail.ChildNodes[I].ChildNodes.FindNode('severity')) then
              aDUE.IdeErro.Items[I].Detalhe.Gravidade   := nodeDetail.ChildNodes[I].ChildValues['severity'];
          end;
        end;
	  end;
    except
      Result := False;
    end;
  finally
//    XMLDocto.Free;
  end;

{  if pos('<error>', Response) > 0 then
  begin
    aDUE.FIdeErro.Mensagem        := '';
    aDUE.FIdeErro.Codigo          := '';
    aDUE.FIdeErro.Campo           := '';
    aDUE.FIdeErro.Tag             := '';
    aDUE.FIdeErro.Data            := '';
    aDUE.FIdeErro.Status          := '';
    aDUE.FIdeErro.Gravidade       := '';
    aDUE.FIdeErro.Info.Ambiente   := '';
    aDUE.FIdeErro.Info.MNemonico  := '';
    aDUE.FIdeErro.Info.Sistema    := '';
    aDUE.FIdeErro.Info.Url        := '';
    aDUE.FIdeErro.Info.Usuario    := '';
    aDUE.FIdeErro.Info.Visao      := '';
  end
  else
    if pos('<pucomexReturn>', Response) > 0 then
    begin
      aDUE.FIdeRetorno.Mensagem     := '';
      aDUE.FIdeRetorno.NroDUE       := '';
      aDUE.FIdeRetorno.NroRUC       := '';
      aDUE.FIdeRetorno.ChaveAcesso  := '';
      aDUE.FIdeRetorno.Data         := '';
      aDUE.FIdeRetorno.CPF          := '';
    end;
	}
end;

function TUtils.NaoEstaVazio(const AValue: String): Boolean;
begin
  Result	:= NOT EstaVazio(AValue);
end;

function TUtils.PathWithDelim(const APath: String): String;
begin
  Result := Trim(APath);
  if Result <> '' then
  begin
    Result := IncludeTrailingPathDelimiter(Result);
  end;
end;

function TUtils.Valida(aDUE: TDUE): Boolean;
var
  Erros: String;
  I, J, K : Integer;
  Inicio : TDateTime;

  procedure GravaLog(AString: String);
  begin
    // DEBUG
    // Log := Log + FormatDateTime('hh:nn:ss:zzz',Now) + ' - ' + AString + sLineBreak;
  end;

  procedure AdicionaErro(const Erro: String);
  begin
    Erros := Erros + Erro + sLineBreak;
  end;

begin
  Inicio  := Now;
//  Agora   := IncMinute(Now, 5);  //Aceita uma tolerância de até 5 minutos, devido ao sincronismo de horário do servidor da Empresa e o servidor da SISCOMEX.

  GravaLog('Inicio da Validação');

  with aDUE do
  begin
    Erros := EmptyStr;

    GravaLog('Validar: 101-Local de Despacho');
    if IdeDespacho.nRFB = 0 then
      AdicionaErro('101-Rejeição: Local de Despacho Ausente');

    GravaLog('Validar: 102-Recinto de Despacho do Despacho');
    if IdeDespacho.nRecinto = 0 then
      AdicionaErro('102-Rejeição: Recinto de Despacho Ausente');

    GravaLog('Validar: 103-Indicador de Recinto Alfandegado do Despacho');
    if IdeDespacho.nTipo = 0 then
      AdicionaErro('103-Rejeição: Indicador de Recinto Alfandegado do Despacho Ausente ou Inválido');    //281, 22,19

    GravaLog('Validar: 104-Forma de Exportação');
    if FormaExportacao.nCodigo = 0 then
      AdicionaErro('104-Rejeição: Forma de Exportação Ausente ou Inválido');     //  1001, 1002, 1003

    GravaLog('Validar: 105-Tipo de Dados Inválido para Forma de Exportação CUS');
    if FormaExportacao.Tipo <> 'CUS' then
      AdicionaErro('105-Rejeição: Tipo de dados inválido para Forma de Exportação CUS');

    GravaLog('Validar: 106-Via Especial de Transporte');  // 4001, 4002, 4003, 4004, 4005
    if ViaTransporte.nCodigo = 0 then
      AdicionaErro('106-Rejeição: Via Especial de Transporte');

    GravaLog('Validar: 107-Tipo da Via Especial de Transporte');
    if ViaTransporte.Tipo <> 'TRA' then
      AdicionaErro('107-Rejeição: Tipo de dados inválido para Via Especial de Transporte');

    GravaLog('Validar: 108-Tipo de Observações Gerais');
    if Observacao.TipoDescricao <> 'AAI' then
      AdicionaErro('108-Rejeição: Tipo de Observações Gerais diferente de AAI');

    GravaLog('Validar: 109-Observações Gerais');
    if Observacao.Desc = EmptyStr then
      AdicionaErro('109-Rejeição: Observações Gerais deve conter alguma informação');

    GravaLog('Validar: 110-Moeda');
    if Cambio.Moeda = EmptyStr then
      AdicionaErro('110-Rejeição: Moeda precisa ser preenchida');

    GravaLog('Validar: 111-Declarante');
    if Declarante.Docto = EmptyStr then
      AdicionaErro('111-Rejeição: CNPJ ou CPF do Declarante precisa ser preenchido');

    GravaLog('Validar: 112-Local de Embarque');
    if IdeEmbarque.nRFB = 0 then
      AdicionaErro('112-Rejeição: Local de Embarque Ausente');

    GravaLog('Validar: 113-Recinto de Embarque');
    if IdeEmbarque.nRecinto = 0 then
      AdicionaErro('113-Rejeição: Recinto de Embarque Ausente');

    GravaLog('Validar: 114-Indicador de Recinto Alfandegado do Embarque');
    if IdeEmbarque.nTipo = 0 then
      AdicionaErro('114-Rejeição: Indicador de Recinto Alfandegado do Embarque Ausente ou Inválido');    //281, 22,19

    GravaLog('Validar: 115-Itens da DUE');
    for I := 0 to Det.Count -1 do
    begin
      with Det.Items[I] do
      begin
        GravaLog('Validar:116-Valor no Local do Embarque Cur(17,2)');
        if Prod.vVMLE <= 0.00 then
          AdicionaErro('116-Rejeição: Valor do Local do Embarque deve ser preenchido, para o item ' + InttoStr(Prod.nItem) + ' da DUE');

        GravaLog('Validar: 117-País de destino');
        if Prod.sPais = '' then
          AdicionaErro('117-Rejeição: País de destino do item ' + IntToStr(Prod.nItem) + ' da DUE precisa ser preenchido');

        { O somatório entre todos os GoodsMeasure.tariffQuantity de diferentes países deve ser igual
        a quantidade total na unidade de medida estatística do item da DU-E.}
        GravaLog('Validar: 118-Quantidade Estatistica enviada ao País de Destino'); // Num(19,5)
        if Prod.qUnidEst <= 0.00000 then
          AdicionaErro('118-Rejeição: Quantidade Estatística no País de Destino do item ' + IntToStr(Prod.nItem) + ' da DUE deve ser preenchida');

        GravaLog('Validar: 119-Descrição complementar da mercadoria');
        if Prod.DescCompl = EmptyStr then
          AdicionaErro('119-Rejeição: Descrição complementar da mercadoria precisa ser preenchida no item ' + IntToStr(Prod.nItem) + ' da DUE');

        GravaLog('Validar: 120-Valor da mercadoria na condição de venda');
        if Prod.vVMCV <= 0.00 then
          AdicionaErro('120-Rejeição: Valor da mercadoria na condição de venda deve ser preenchida no item '+ IntToStr(Prod.nItem) + ' da DUE');

        GravaLog('Validar: 121-Número do item da NF de Exportação não preenchido');
        if Prod.NFRef.nItemNF = 0 then
          AdicionaErro('121-Rejeição: Número do item da NF de Exportação não preenchido para o item ' + InttoStr(Prod.nItem) + ' da DUE');

        GravaLog('Validar: 122-Itens da NF de Remessa ou Complementar');
        for J := 0 to Prod.NFRef.Count -1 do
        begin
          GravaLog('Validar: 123-Número do item na NF de Remessa ou Complementar para o item ' + IntToStr(Prod.nItem)  + ' da DUE');
          if Prod.NFRef.Items[J].nItem = 0 then
            AdicionaErro('123-Rejeição: Número do item na NF de Remessa ou Complementar para o item ' + IntToStr(Prod.nItem)  + ' da DUE');

          GravaLog('Validar: 124-Chave de Acesso da NF de Remessa ou Complementar');
          if Prod.NFRef.Items[J].nChave = EmptyStr then
            AdicionaErro('124-Rejeição: Chave de Acesso da NF de Remessa ou Complementar precisa ser preenchida para o item ' + IntToStr(Prod.nItem) + ' da DUE');

          GravaLog('Validar: 125-Quantidade na unidade de medida estatística a ser consumida da Nota Fisca de Remessa'); // Num(19,5)
          if Prod.NFRef.Items[J].qUnidEst = 0.00000 then
            AdicionaErro('125-Rejeição: Quantidade na unidade de medida estatística a ser consumida da Nota Fisca de Remessa para o item ' + IntToStr(Prod.nItem) + ' da DUE');
        end;

        GravaLog('Validar: 126-Peso Líquido Total do item da DU-E em KG'); //Num(19,5)
        if Prod.vPesoLiqTotal = 0.00000 then
          AdicionaErro('126-Rejeição: Peso Líquido Total em KG do item ' + IntToStr(Prod.nItem) + ' da DUE deve ser preenchido');

        GravaLog('Validar: 127-Equadramento');
        if Prod.Enquad.nEnquad01 = 0 then
          AdicionaErro('127-Rejeição: Enquadramento deve ser preenchido para o item ' + IntToStr(Prod.nItem) + ' da DUE');

        GravaLog('Validar: 128-Nota Fiscal de Exportação');
        if Prod.Notas.nChave = EmptyStr then
          AdicionaErro('128-Rejeição: Chave da Nota Fiscal de Exportação deve ser preenchida para o item ' + IntToStr(Prod.nItem) + ' da DUE');

        GravaLog('Validar: 129-Tipo da Nota Fiscal de Exportação');
        if Prod.Notas.nTipo = 0 then
          AdicionaErro('129-Rejeição: Tipo da Nota Fiscal de Exportação deve ser preenchido para o item ' + IntToStr(Prod.nItem) + ' da DUE');

        GravaLog('Validar: 130-Notas Referenciadas');
        for K := 0 to Prod.Notas.Count -1 do
        begin
          GravaLog('Validar: 131-Chave da Nota referenciada');
          if Prod.Notas.Items[K].nChave = EmptyStr then
            AdicionaErro('131-Rejeição: Chave da Nota referenciada deve ser preenchida para o item ' + IntToStr(Prod.nItem) + ' da DUE');

          GravaLog('Validar: 132-Tipo da Nota Referenciada');
          if Prod.Notas.Items[K].nTipo = EmptyStr then
            AdicionaErro('132-Rejeição: Tipo da Nota Referenciada precisa ser preenchdia para o item ' + IntToStr(Prod.nItem) + ' da DUE');

          GravaLog('Validar: 133-CNPJ ou CPF do emissor da Nota Fiscal');
          if Prod.Notas.Items[K].FnDocto = EmptyStr then
            AdicionaErro('133-Rejeição: CNPJ ou CPF do emissor da Nota Fiscal deve ser preenchido para o item ' + IntToStr(Prod.nItem) + 'da DUE');
        end;

        GravaLog('Validar: 134-Condição de Venda');
        if Prod.FsCondVenda = EmptyStr then
          AdicionaErro('134-Rejeição: Condição de Venda deve ser preenchido para o item ' + IntToStr(Prod.nItem) + ' da DUE');
      end;

    end;

  end;

  Result := (Erros = EmptyStr);

  if not Result then
  begin
    Erros	:= 'Erro(s) nas Regras de negócios da DU-E' + sLineBreak + Erros;
  end;

  GravaLog('Fim da Validação. Tempo: ' + FormatDateTime('hh:nn:ss:zzz', Now - Inicio) + sLineBreak +
           'Erros:' + Erros);

  //DEBUG
  //WriteToTXT('c:\temp\DUE Validacao.txt', Log);

  FErroRegrasNegocios := Erros;
end;

procedure TUtils.WriteToTXT(const ArqTXT: String; const ABinaryString: AnsiString;
  const AppendIfExists, AddLineBreak, ForceDirectory: Boolean);
var
  FS : TFileStream;
  LineBreak : AnsiString;
  VDirectory : String;
  ArquivoExiste: Boolean;
begin
  if EstaVazio(ArqTXT) then
    Exit;

  ArquivoExiste := FileExists(ArqTXT);

  if ArquivoExiste then
  begin
    if (Length(ABinaryString) = 0) then
      Exit;
  end
  else
  begin
     if ForceDirectory then
     begin
       VDirectory := ExtractFileDir(ArqTXT);
       if NaoEstaVazio(VDirectory) and (not DirectoryExists(VDirectory)) then
         ForceDirectories(VDirectory);
     end;
  end;

  FS := TFileStream.Create( ArqTXT, fmCreate, fmShareDenyWrite);
  try
     FS.Seek(0, soEnd); // {$IFDEF COMPILER23_UP}soEnd{$ELSE}soFromEnd{$ENDIF});  // vai para EOF
     FS.Write(Pointer(ABinaryString)^,Length(ABinaryString));

     if AddLineBreak then
     begin
        LineBreak := sLineBreak;
        FS.Write(Pointer(LineBreak)^,Length(LineBreak));
     end ;
  finally
     FS.Free ;
  end;
end;

function TUtils.XmlEhUTF8(const AXML: String): Boolean;
begin
  Result := (pos('encoding="utf-8"', LowerCase(LeftStr(AXML, 50))) > 0);
end;

{ EDUEException }

constructor EDUEException.Create(const Msg: String);
begin
  inherited Create(Msg);
end;

{ TIdeRetorno }

procedure TIdeRetorno.Assign(Source: TPersistent);
begin
  if Source is TIdeRetorno then
  begin
    Mensagem    := TIdeRetorno(Source).Mensagem;
    NroDUE      := TIdeRetorno(Source).NroDUE;
    NroRUC      := TIdeRetorno(Source).NroRUC;
    ChaveAcesso := TIdeRetorno(Source).ChaveAcesso;
    Data        := TIdeRetorno(Source).Data;
    CPF         := TIdeRetorno(Source).CPF;
  end
  else
    inherited;
end;

{ TErroCollection }

function TErroCollection.Add: TErroCollectionItem;
begin
  Result := TErroCollectionItem(inherited Add);
end;

procedure TErroCollection.Assign(Source: TPersistent);
begin
  if Source is TErroCollection then
  begin
    Mensagem  := TErroCollection(Source).Mensagem;
    Codigo    := TErroCollection(Source).Codigo;
    Campo     := TErroCollection(Source).Campo;
    Tag       := TErroCollection(Source).Tag;
    Data      := TErroCollection(Source).Data;
    Status    := TErroCollection(Source).Status;
    Gravidade := TErroCollection(Source).Gravidade;
    Info      := TErroCollection(Source).Info;
  end
  else
    inherited;
end;

constructor TErroCollection.Create(AOwner: TDUE);
begin
  inherited Create(TErroCollectionItem);
  FInfo := TErroInfo.Create;
end;

function TErroCollection.GetItem(Index: Integer): TErroCollectionItem;
begin
  Result := TErroCollectionItem(inherited GetItem(Index));
end;

procedure TErroCollection.SetItem(Index: Integer;
  const Value: TErroCollectionItem);
begin
  inherited SetItem(Index, Value);
end;

{ TErroInfor }
procedure TErroInfo.Assign(Source: TPersistent);
begin
  if Source is TErroInfo then
  begin
    Ambiente := TErroInfo(Source).Ambiente;
    MNemonico := TErroInfo(Source).MNemonico;
    Sistema   := TErroInfo(Source).Sistema;
    Url       := TErroInfo(Source).Url;
    Usuario   := TErroInfo(Source).Usuario;
    Visao     := TErroInfo(Source).Visao;
  end
  else
    inherited;
end;

{ TErroCollectionItem }
procedure TErroCollectionItem.Assign(Source: TPersistent);
begin
  if Source is TErroCollectionItem then
  begin
    Mensagem  := TErroCollectionItem(Source).Mensagem;
    Codigo    := TErroCollectionItem(Source).Codigo;
    Tag       := TErroCollectionItem(Source).Tag;
    Data      := TErroCollectionItem(Source).Data;
    Status    := TErroCollectionItem(Source).Status;
    Gravidade := TErroCollectionItem(Source).Gravidade;
  end
  else
    inherited;
end;

constructor TErroCollectionItem.Create(Collection: TCollection);
begin
  inherited;
  FDetalhe := TDetalhe.Create(Self);
end;
 
destructor TErroCollectionItem.Destroy;
begin
  FDetalhe.Free;
  inherited;
end;

{ TErroDetalhe }
 
procedure TDetalhe.Assign(Source: TPersistent);
begin
  if Source is TProd then
  begin
     Mensagem   := TDetalhe(Source).Mensagem;
     Codigo     := TDetalhe(Source).Codigo;
     Tag          := TDetalhe(Source).Tag;
     Data             := TDetalhe(Source).Data;
     Status       := TDetalhe(Source).Status;
     Gravidade  := TDetalhe(Source).Gravidade;
  end
  else
    inherited;
end;
 
constructor TDetalhe.Create(AOwner: TErroCollectionItem);
begin
  inherited Create;
 
end;
 
destructor TDetalhe.Destroy;
begin
 
  inherited;
end;

end.