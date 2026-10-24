"""Gera o relatório paginado "Ficha diária de produção" (FichaDiaria.rdl) da demo B.

Fonte: o modelo da demo B aberto no Power BI Desktop (Analysis Services local). Como a porta muda
a cada vez que o Desktop abre, o script lê a porta atual e grava no .rdl. Rode de novo sempre que
reabrir o Desktop:  python gerar_ficha.py

Uma página por injetora para a data escolhida (parâmetro Data):
números do dia, produção por turno, paradas por motivo x turno, alerta de apontamentos incompletos,
espaço para observações e assinatura do supervisor.
"""
import base64
import glob
import os
import uuid
from xml.sax.saxutils import escape

PASTA = os.path.dirname(os.path.abspath(__file__))
LOGO = "C:/Users/Vitor/.claude/freelance/portfolio/logo/final/vitor-sato-logo-principal-claro.png"
WORKSPACES = os.path.expanduser("~/Microsoft/Power BI Desktop Store App/AnalysisServicesWorkspaces")

# cores da versão clara da marca (papel)
MARINHO, CIANO, LARANJA, VERMELHO = "#1E3A5F", "#0F8A80", "#DD6B20", "#C53030"
CINZA_TEXTO, CINZA_FUNDO, CINZA_LINHA = "#4A6A90", "#F4F6F9", "#D5DCE6"


def porta_atual():
    arquivos = glob.glob(os.path.join(WORKSPACES, "*", "Data", "msmdsrv.port.txt"))
    if not arquivos:
        raise SystemExit("Power BI Desktop não está aberto (nenhum msmdsrv.port.txt encontrado).")
    mais_recente = max(arquivos, key=os.path.getmtime)
    with open(mais_recente, "rb") as f:
        return f.read().decode("utf-16").strip()


# ---------- consulta DAX (validada contra o modelo) ----------
DAX = """DEFINE VAR Dia = TREATAS ( { DATE ( @ANO, @MES, @DIA ) }, dCalendario[Data] )
EVALUATE
VAR Turnos = CALCULATETABLE (
    ADDCOLUMNS (
        SUMMARIZE ( fApontamentos, dMaquina[Máquina], dMaquina[Tonelagem], dTurno[Turno], dTurno[Nome] ),
        "@OP", CALCULATE ( MAX ( fApontamentos[OP] ) ),
        "@Produto", CALCULATE ( MAXX ( fApontamentos, RELATED ( dProduto[Descrição] ) ) ),
        "@Produzido", [Produzido], "@Previsto", [Previsto], "@Refugo", [Refugo],
        "@MinParado", [Tempo Parado (min)], "@OEE", [OEE], "@Problemas", [Apontamentos com Problema],
        "@EficienciaDia", CALCULATE ( [Eficiência vs Previsto], REMOVEFILTERS ( dTurno ) ),
        "@RefugoPctDia", CALCULATE ( [% Refugo], REMOVEFILTERS ( dTurno ) ),
        "@OEEDia", CALCULATE ( [OEE], REMOVEFILTERS ( dTurno ) )
    ), Dia )
VAR Paradas = CALCULATETABLE (
    ADDCOLUMNS (
        SUMMARIZE ( fParadas, dMaquina[Máquina], dMaquina[Tonelagem], dTurno[Turno], dTurno[Nome], dMotivo[Motivo] ),
        "@Min", [Min Parada]
    ), Dia )
RETURN UNION (
    SELECTCOLUMNS ( Turnos, "Tipo", "T", "Maquina", dMaquina[Máquina], "Tonelagem", dMaquina[Tonelagem],
        "TurnoCod", dTurno[Turno], "Turno", dTurno[Nome], "OP", [@OP], "Produto", [@Produto],
        "Produzido", [@Produzido], "Previsto", [@Previsto], "Refugo", [@Refugo], "MinParado", [@MinParado],
        "OEE", [@OEE], "Problemas", [@Problemas], "EficienciaDia", [@EficienciaDia],
        "RefugoPctDia", [@RefugoPctDia], "OEEDia", [@OEEDia], "Motivo", BLANK (), "MinMotivo", BLANK () ),
    SELECTCOLUMNS ( Paradas, "Tipo", "P", "Maquina", dMaquina[Máquina], "Tonelagem", dMaquina[Tonelagem],
        "TurnoCod", dTurno[Turno], "Turno", dTurno[Nome], "OP", BLANK (), "Produto", BLANK (),
        "Produzido", BLANK (), "Previsto", BLANK (), "Refugo", BLANK (), "MinParado", BLANK (),
        "OEE", BLANK (), "Problemas", BLANK (), "EficienciaDia", BLANK (),
        "RefugoPctDia", BLANK (), "OEEDia", BLANK (), "Motivo", dMotivo[Motivo], "MinMotivo", [@Min] )
)"""

CAMPOS = [  # nome, tipo
    ("Tipo", "System.String"), ("Maquina", "System.String"), ("Tonelagem", "System.Int64"),
    ("TurnoCod", "System.String"), ("Turno", "System.String"), ("OP", "System.String"), ("Produto", "System.String"),
    ("Produzido", "System.Int64"), ("Previsto", "System.Int64"), ("Refugo", "System.Int64"), ("MinParado", "System.Int64"),
    ("OEE", "System.Double"), ("Problemas", "System.Int64"), ("EficienciaDia", "System.Double"),
    ("RefugoPctDia", "System.Double"), ("OEEDia", "System.Double"), ("Motivo", "System.String"), ("MinMotivo", "System.Int64"),
]


def comando_dax():
    """A consulta vira uma expressão VB: a data do parâmetro entra no DATE() do DAX."""
    vb = DAX.replace('"', '""')
    vb = vb.replace("@ANO", '" & Year(Parameters!Data.Value) & "')
    vb = vb.replace("@MES", '" & Month(Parameters!Data.Value) & "')
    vb = vb.replace("@DIA", '" & Day(Parameters!Data.Value) & "')
    vb = vb.replace("\r\n", "\n").replace("\n", '" & vbCrLf & "')
    return '="' + vb + '"'


def campo(nome, tipo):
    df = ('<?xml version="1.0" encoding="utf-8"?><Field xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
          f'xsi:type="Measure" UniqueName="[{nome}]" />')
    return f"<Field Name=\"{nome}\"><DataField>{escape(df)}</DataField><rd:TypeName>{tipo}</rd:TypeName></Field>"


# ---------- blocos de layout ----------
_n = {"i": 0}


def nome(prefixo):
    _n["i"] += 1
    return f"{prefixo}{_n['i']}"


def caixa(valor, top=None, left=None, h=None, w=None, tam="9pt", peso="Normal", cor=MARINHO, fundo=None,
          alinhar="Left", formato=None, borda_inf=None, cor_expr=None, prefixo="tx", vertical="Middle", pad="3pt"):
    """Textbox. Sem top/left/h/w quando vai dentro de célula de tabela."""
    v = escape(valor)
    estilo_run = f"<FontFamily>Segoe UI</FontFamily><FontSize>{tam}</FontSize><FontWeight>{peso}</FontWeight>"
    estilo_run += f"<Color>{escape(cor_expr) if cor_expr else cor}</Color>"
    if formato:
        estilo_run += f"<Format>{escape(formato)}</Format>"
    pos = "" if top is None else f"<Top>{top}cm</Top><Left>{left}cm</Left><Height>{h}cm</Height><Width>{w}cm</Width>"
    est = "<Border><Style>None</Style></Border>"
    if borda_inf:
        est += f"<BottomBorder><Color>{borda_inf}</Color><Style>Solid</Style><Width>0.75pt</Width></BottomBorder>"
    if fundo:
        est += f"<BackgroundColor>{fundo}</BackgroundColor>"
    est += f"<VerticalAlign>{vertical}</VerticalAlign><PaddingLeft>{pad}</PaddingLeft><PaddingRight>{pad}</PaddingRight><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom>"
    return (f'<Textbox Name="{nome(prefixo)}"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether>'
            f"<Paragraphs><Paragraph><TextRuns><TextRun><Value>{v}</Value><Style>{estilo_run}</Style></TextRun></TextRuns>"
            f"<Style><TextAlign>{alinhar}</TextAlign></Style></Paragraph></Paragraphs>{pos}<Style>{est}</Style></Textbox>")


def celula(conteudo):
    return f"<TablixCell><CellContents>{conteudo}</CellContents></TablixCell>"


def tabela(nome_t, colunas, cabecalhos, valores, grupo_expr, filtro_tipo, ordem, top, alinhamentos, formatos, cores=None):
    """Tabela com 1 linha de cabeçalho e 1 linha por grupo, filtrada por Tipo (T = turnos, P = paradas)."""
    cols = "".join(f"<TablixColumn><Width>{w}cm</Width></TablixColumn>" for w in colunas)
    cab = "".join(celula(caixa(c, tam="8pt", peso="Bold", cor="White", fundo=MARINHO, alinhar=a))
                  for c, a in zip(cabecalhos, alinhamentos))
    cores = cores or [None] * len(valores)
    lin = "".join(celula(caixa(v, tam="8.5pt", alinhar=a, formato=f, borda_inf=CINZA_LINHA, cor_expr=c))
                  for v, a, f, c in zip(valores, alinhamentos, formatos, cores))
    membros_col = "<TablixMember />" * len(colunas)
    largura = sum(colunas)
    return f"""<Tablix Name="{nome_t}"><TablixBody><TablixColumns>{cols}</TablixColumns><TablixRows>
<TablixRow><Height>0.55cm</Height><TablixCells>{cab}</TablixCells></TablixRow>
<TablixRow><Height>0.5cm</Height><TablixCells>{lin}</TablixCells></TablixRow>
</TablixRows></TablixBody>
<TablixColumnHierarchy><TablixMembers>{membros_col}</TablixMembers></TablixColumnHierarchy>
<TablixRowHierarchy><TablixMembers><TablixMember><KeepWithGroup>After</KeepWithGroup></TablixMember>
<TablixMember><Group Name="{nome_t}Grupo"><GroupExpressions><GroupExpression>{escape(grupo_expr)}</GroupExpression></GroupExpressions></Group>
<SortExpressions><SortExpression><Value>{escape(ordem[0])}</Value><Direction>{ordem[1]}</Direction></SortExpression></SortExpressions></TablixMember>
</TablixMembers></TablixRowHierarchy>
<DataSetName>Ficha</DataSetName>
<Filters><Filter><FilterExpression>=Fields!Tipo.Value</FilterExpression><Operator>Equal</Operator><FilterValues><FilterValue>{filtro_tipo}</FilterValue></FilterValues></Filter></Filters>
<Top>{top}cm</Top><Left>0cm</Left><Height>1.05cm</Height><Width>{largura}cm</Width><Style><Border><Style>None</Style></Border></Style></Tablix>"""


def kpi(rotulo, valor, left, formato, cor_expr=None):
    w = 3.48
    return (f'<Rectangle Name="{nome("kpi")}"><ReportItems>'
            + caixa(rotulo, 0.15, 0.2, 0.5, w - 0.4, tam="7.5pt", cor=CINZA_TEXTO)
            + caixa(valor, 0.65, 0.2, 0.8, w - 0.4, tam="15pt", peso="Bold", formato=formato, cor_expr=cor_expr)
            + f"</ReportItems><KeepTogether>true</KeepTogether><Top>1.1cm</Top><Left>{left}cm</Left><Height>1.6cm</Height><Width>{w}cm</Width>"
            f"<Style><Border><Style>None</Style></Border><BackgroundColor>{CINZA_FUNDO}</BackgroundColor></Style></Rectangle>")


def corpo_maquina():
    oee_cor = f'=IIF(Max(Fields!OEEDia.Value) < 0.75, "{VERMELHO}", "{CIANO}")'
    itens = [
        caixa('="Injetora " & Fields!Maquina.Value & "  ·  " & First(Fields!Tonelagem.Value) & " t"', 0, 0, 0.9, 12, tam="16pt", peso="Bold"),
        caixa('=Format(Parameters!Data.Value, "dddd, dd/MM/yyyy")', 0, 12, 0.9, 6.6, tam="10pt", cor=CINZA_TEXTO, alinhar="Right"),
        kpi("Produzido (peças)", "=Sum(Fields!Produzido.Value)", 0, "#,0"),
        kpi("Previsto (peças)", "=Sum(Fields!Previsto.Value)", 3.78, "#,0"),
        kpi("Eficiência vs previsto", "=Max(Fields!EficienciaDia.Value)", 7.56, "0.0%"),
        kpi("Refugo", "=Max(Fields!RefugoPctDia.Value)", 11.34, "0.0%"),
        kpi("OEE do dia (meta 75%)", "=Max(Fields!OEEDia.Value)", 15.12, "0.0%", oee_cor),
        caixa("Produção por turno", 3.0, 0, 0.6, 18.6, tam="10.5pt", peso="Bold"),
        tabela("Turnos", [2.3, 2.4, 4.7, 2.0, 2.0, 1.6, 1.8, 1.8],
               ["Turno", "OP", "Produto", "Produzido", "Previsto", "Refugo", "Min parado", "OEE"],
               ["=First(Fields!Turno.Value)", "=First(Fields!OP.Value)", "=First(Fields!Produto.Value)",
                "=Sum(Fields!Produzido.Value)", "=Sum(Fields!Previsto.Value)", "=Sum(Fields!Refugo.Value)",
                "=Sum(Fields!MinParado.Value)", "=Max(Fields!OEE.Value)"],
               "=Fields!TurnoCod.Value", "T", ("=Fields!TurnoCod.Value", "Ascending"), 3.65,
               ["Left", "Left", "Left", "Right", "Right", "Right", "Right", "Right"],
               [None, None, None, "#,0", "#,0", "#,0", "#,0", "0.0%"],
               [None] * 7 + [f'=IIF(Max(Fields!OEE.Value) < 0.75, "{VERMELHO}", "{MARINHO}")']),
        caixa("Paradas por motivo (minutos)", 5.6, 0, 0.6, 18.6, tam="10.5pt", peso="Bold"),
        tabela("Paradas", [6.6, 3.0, 3.0, 3.0, 3.0],
               ["Motivo", "1º turno", "2º turno", "3º turno", "Total"],
               ["=Fields!Motivo.Value",
                '=Sum(IIF(Fields!TurnoCod.Value = "T1", Fields!MinMotivo.Value, Nothing))',
                '=Sum(IIF(Fields!TurnoCod.Value = "T2", Fields!MinMotivo.Value, Nothing))',
                '=Sum(IIF(Fields!TurnoCod.Value = "T3", Fields!MinMotivo.Value, Nothing))',
                "=Sum(Fields!MinMotivo.Value)"],
               "=Fields!Motivo.Value", "P", ("=Sum(Fields!MinMotivo.Value)", "Descending"), 6.25,
               ["Left", "Right", "Right", "Right", "Right"], [None, "#,0", "#,0", "#,0", "#,0"]),
        caixa('=IIF(Sum(Fields!Problemas.Value) > 0, "Atenção: " & Sum(Fields!Problemas.Value) & " apontamento(s) incompleto(s) nesta injetora no dia. Conferir com os operadores antes de fechar a ficha.", "Todos os apontamentos do dia estão completos.")',
              7.6, 0, 0.7, 18.6, tam="8.5pt", peso="Bold",
              cor_expr=f'=IIF(Sum(Fields!Problemas.Value) > 0, "{LARANJA}", "{CIANO}")'),
        caixa("Observações do supervisor", 8.6, 0, 0.6, 18.6, tam="10.5pt", peso="Bold"),
        caixa(" ", 9.25, 0, 0.75, 18.6, borda_inf=CINZA_LINHA),
        caixa(" ", 10.0, 0, 0.75, 18.6, borda_inf=CINZA_LINHA),
        caixa(" ", 10.75, 0, 0.75, 18.6, borda_inf=CINZA_LINHA),
        caixa("Supervisor: ______________________________", 12.1, 0, 0.7, 9.3, tam="9pt", vertical="Bottom"),
        caixa("Assinatura: ______________________________", 12.1, 9.3, 0.7, 9.3, tam="9pt", vertical="Bottom", alinhar="Right"),
    ]
    return ("<Rectangle Name=\"Maquina\"><ReportItems>" + "".join(itens)
            + "</ReportItems><KeepTogether>true</KeepTogether><Style><Border><Style>None</Style></Border></Style></Rectangle>")


def lista():
    return f"""<Tablix Name="Lista"><TablixBody><TablixColumns><TablixColumn><Width>18.6cm</Width></TablixColumn></TablixColumns>
<TablixRows><TablixRow><Height>13cm</Height><TablixCells>{celula(corpo_maquina())}</TablixCells></TablixRow></TablixRows></TablixBody>
<TablixColumnHierarchy><TablixMembers><TablixMember /></TablixMembers></TablixColumnHierarchy>
<TablixRowHierarchy><TablixMembers><TablixMember>
<Group Name="PorMaquina"><GroupExpressions><GroupExpression>=Fields!Maquina.Value</GroupExpression></GroupExpressions>
<PageBreak><BreakLocation>Between</BreakLocation></PageBreak></Group>
<SortExpressions><SortExpression><Value>=Fields!Maquina.Value</Value></SortExpression></SortExpressions>
</TablixMember></TablixMembers></TablixRowHierarchy>
<DataSetName>Ficha</DataSetName><Top>0cm</Top><Left>0cm</Left><Height>13cm</Height><Width>18.6cm</Width>
<Style><Border><Style>None</Style></Border></Style></Tablix>"""


def cabecalho():
    return f"""<PageHeader><Height>2cm</Height><PrintOnFirstPage>true</PrintOnFirstPage><PrintOnLastPage>true</PrintOnLastPage><ReportItems>
<Image Name="Logo"><Source>Embedded</Source><Value>Logo</Value><Sizing>FitProportional</Sizing><Top>0cm</Top><Left>0cm</Left><Height>1.5cm</Height><Width>4.5cm</Width><Style><Border><Style>None</Style></Border></Style></Image>
{caixa("Ficha diária de produção", 0, 6.1, 0.9, 12.5, tam="17pt", peso="Bold", alinhar="Right")}
{caixa('="Fábrica Exemplo  ·  " & Format(Parameters!Data.Value, "dd/MM/yyyy") & "  ·  dados fictícios"', 0.9, 6.1, 0.6, 12.5, tam="9pt", cor=CINZA_TEXTO, alinhar="Right")}
<Line Name="LinhaTopo"><Top>1.75cm</Top><Left>0cm</Left><Height>0cm</Height><Width>18.6cm</Width><Style><Border><Color>{LARANJA}</Color><Style>Solid</Style><Width>1.5pt</Width></Border></Style></Line>
</ReportItems><Style><Border><Style>None</Style></Border></Style></PageHeader>"""


def rodape():
    return f"""<PageFooter><Height>0.8cm</Height><PrintOnFirstPage>true</PrintOnFirstPage><PrintOnLastPage>true</PrintOnLastPage><ReportItems>
{caixa("Vitor Sato  ·  Power BI e análise de dados", 0.2, 0, 0.5, 9.3, tam="7.5pt", cor=CINZA_TEXTO)}
{caixa('="Página " & Globals!PageNumber & " de " & Globals!TotalPages & "  ·  gerado em " & Format(Globals!ExecutionTime, "dd/MM/yyyy HH:mm")', 0.2, 9.3, 0.5, 9.3, tam="7.5pt", cor=CINZA_TEXTO, alinhar="Right")}
</ReportItems><Style><Border><Style>None</Style></Border></Style></PageFooter>"""


def rdl(porta):
    with open(LOGO, "rb") as f:
        logo64 = base64.b64encode(f.read()).decode()
    campos = "".join(campo(n, t) for n, t in CAMPOS)
    return f"""<?xml version="1.0" encoding="utf-8"?>
<Report MustUnderstand="df" xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition" xmlns:rd="http://schemas.microsoft.com/SQLServer/reporting/reportdesigner" xmlns:df="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily">
<rd:ReportUnitType>Cm</rd:ReportUnitType>
<rd:ReportID>{uuid.uuid4()}</rd:ReportID>
<df:DefaultFontFamily>Segoe UI</df:DefaultFontFamily>
<Description>Ficha diária de produção por injetora (demo B, dados fictícios)</Description>
<Author>Vitor Sato</Author>
<AutoRefresh>0</AutoRefresh>
<DataSources><DataSource Name="ModeloDemoB"><ConnectionProperties><DataProvider>OLEDB-MD</DataProvider>
<ConnectString>Data Source=localhost:{porta}</ConnectString><IntegratedSecurity>true</IntegratedSecurity></ConnectionProperties>
<rd:SecurityType>Integrated</rd:SecurityType><rd:DataSourceID>{uuid.uuid4()}</rd:DataSourceID></DataSource></DataSources>
<DataSets><DataSet Name="Ficha"><Query><DataSourceName>ModeloDemoB</DataSourceName>
<CommandText>{escape(comando_dax())}</CommandText></Query>
<Fields>{campos}</Fields></DataSet></DataSets>
<ReportSections><ReportSection><Body><ReportItems>{lista()}</ReportItems><Height>13cm</Height><Style><Border><Style>None</Style></Border></Style></Body>
<Width>18.6cm</Width>
<Page>{cabecalho()}{rodape()}
<PageHeight>29.7cm</PageHeight><PageWidth>21cm</PageWidth><InteractiveHeight>29.7cm</InteractiveHeight><InteractiveWidth>21cm</InteractiveWidth>
<LeftMargin>1.2cm</LeftMargin><RightMargin>1.2cm</RightMargin><TopMargin>1.2cm</TopMargin><BottomMargin>1.2cm</BottomMargin><Style /></Page>
</ReportSection></ReportSections>
<ReportParameters><ReportParameter Name="Data"><DataType>DateTime</DataType>
<DefaultValue><Values><Value>2026-09-30</Value></Values></DefaultValue><Prompt>Data da ficha</Prompt></ReportParameter></ReportParameters>
<ReportParametersLayout><GridLayoutDefinition><NumberOfColumns>1</NumberOfColumns><NumberOfRows>1</NumberOfRows>
<CellDefinitions><CellDefinition><ColumnIndex>0</ColumnIndex><RowIndex>0</RowIndex><ParameterName>Data</ParameterName></CellDefinition></CellDefinitions>
</GridLayoutDefinition></ReportParametersLayout>
<EmbeddedImages><EmbeddedImage Name="Logo"><MIMEType>image/png</MIMEType><ImageData>{logo64}</ImageData></EmbeddedImage></EmbeddedImages>
<Language>pt-BR</Language>
</Report>"""


if __name__ == "__main__":
    porta = porta_atual()
    destino = os.path.join(PASTA, "FichaDiaria.rdl")
    with open(destino, "w", encoding="utf-8-sig") as f:
        f.write(rdl(porta))
    print(f"{destino} (porta {porta})")
