import pandas as pd;
import numpy as np;
from datetime import datetime;
from datetime import date;
import os;
from openpyxl import load_workbook;
from openpyxl.worksheet.datavalidation import DataValidation;
from openpyxl.worksheet.table import Table, TableStyleInfo;

dataHoje = date.today().strftime("%d-%m-%Y");

tabelaTop1Caminho = r"\\base-bhz\MAT_PLE\02 Liberações na Bancada e Transferencias\Transferencias Firme" + f"\{dataHoje} T1.xlsx";

if os.path.isfile(tabelaTop1Caminho):
    tabelaTop1Extrac = pd.read_excel(tabelaTop1Caminho);
    top1Existe = True;
else:
    tabelaTop1Extrac = "";
    top1Existe = False;

tabelaTop2Caminho = r"\\base-bhz\MAT_PLE\02 Liberações na Bancada e Transferencias\Transferencias Firme" + f"\{dataHoje} T2.xlsx";

if os.path.isfile(tabelaTop2Caminho):
    tabelaTop2Extrac = pd.read_excel(tabelaTop2Caminho);
    top2Existe = True;
    
else:
    tabelaTop1Extrac = "";
    top2Existe = False;

listaTopExist = [top1Existe, top2Existe];
listaTopicos = [tabelaTop1Extrac, tabelaTop2Extrac];

diasTransf = int(input("Digite a quantidade de dias de transferência para análise: "));

diasCompra = int(input("Digite a quantidade de dias de compra para análise: "));

undNegoc = input("Digite a unidade de negocio: ");

undNegocSave = undNegoc;

if (undNegoc == "UOH"):
    undNegoc = "HELICOPTEROS";
elif (undNegoc == "UME"):
    undNegoc = "MANUTENCAO";

wbLiberacaoCaminho = r"C:\Users\ddias\Desktop\Liberações Novembro - 2023 - " + f"{undNegocSave}.xlsx";

dataComparacaoTransf = datetime.now()+pd.Timedelta(days=diasTransf);

dataComparacaoCompra = datetime.now()+pd.Timedelta(days=diasCompra);

dataHoje = date.today().strftime("%d-%m");

wsTopico1Nome = dataHoje + " T1"
wsTopico2Nome = dataHoje + " T2"

numLinhasT1 = 0;
numLinhasT2 = 0;

def filtraTopico1(wbLiberacaoCaminho, wsTopico1Nome):
    nomesColunas = [];

    for coluna in listaTopicos[0].columns.array:
        nomesColunas.append(coluna);
        
    arrayTop1Extrac = np.array(listaTopicos[0]);

    colDataPlan = 11;
    
    colUndNegoc = 1;

    arrayFiltroDataPlan = np.array([registro for registro in arrayTop1Extrac if registro[colDataPlan] <= dataComparacaoTransf]);
    
    arrayFiltroUndNegoc = np.array([registro for registro in arrayFiltroDataPlan if registro[colUndNegoc] == undNegoc]);

    if arrayFiltroUndNegoc.size == 0:
        print("Após a filtragem o tópico 1 ficou vazio.")

        tabelaTop1Filtrada = pd.DataFrame(index=range(1), columns=nomesColunas);

        tabelaTop1Filtrada["PN&PREFIXO&QTY"] = "";

        tabelaTop1Filtrada["Problema Raiz"] = "";

        tabelaTop1Filtrada["Ação"] = "";

        tabelaTop1Filtrada["OBS"] = "";

        tabelaTop1Filtrada["PN s/C"] = "";

        tabelaTop1Filtrada["Condição"] = "";
        
        tabelaTop1Filtrada["Desc."] = "";
    
    else:
        tabelaTop1Filtrada = pd.DataFrame(arrayFiltroUndNegoc, columns=nomesColunas);

        tabelaTop1Filtrada["PN&PREFIXO&QTY"] = tabelaTop1Filtrada["PN"] + tabelaTop1Filtrada["PREFIXO"].apply(str) + tabelaTop1Filtrada["QTDE. PENDENTE"].apply(int).apply(str);

        tabelaTop1Filtrada["Problema Raiz"] = "";

        tabelaTop1Filtrada["Ação"] = "";

        tabelaTop1Filtrada["OBS"] = "";

        tabelaTop1Filtrada["PN s/C"] = tabelaTop1Filtrada["PN"].str[:-2];

        tabelaTop1Filtrada["Condição"] = tabelaTop1Filtrada["PN"].str[-2:];
        
        tabelaTop1Filtrada["Desc."] = tabelaTop1Filtrada["DESCRIÇÃO"];
        
    with pd.ExcelWriter(wbLiberacaoCaminho, mode="a", if_sheet_exists="overlay") as writer:
        tabelaTop1Filtrada.to_excel(excel_writer=writer, sheet_name=wsTopico1Nome, index=False);

def filtraTopico2(wbLiberacaoCaminho, wsTopico2Nome):
    nomesColunas = [];

    for coluna in listaTopicos[1].columns.array:
        nomesColunas.append(coluna);
        
    listaTopicos[1]["DATA DA ORDEM SUGERIDA"] = pd.to_datetime(listaTopicos[1]["DATA DA ORDEM SUGERIDA"], format="%d/%m/%Y %H:%M:%S");

    arrayTop2Extrac = np.array(listaTopicos[1]);

    colDtOrdemSug = 15;

    colTipoOrdem = 5;

    colUndNegoc = 3;

    colOrgOrigem = 6;

    colPrefixo = 12;

    tabelaFiltroDataOrdemSugerida = np.array([registro for registro in arrayTop2Extrac if registro[colDtOrdemSug] <= dataComparacaoTransf and registro[colTipoOrdem] == "Entrega inbound planejada" or registro[colDtOrdemSug] <= dataComparacaoCompra and registro[5] == "Ordem planejada"]);

    tabelaFiltroUndNegoc = np.array([registro for registro in tabelaFiltroDataOrdemSugerida if registro[colUndNegoc] == undNegoc]);

    if (undNegoc == "MANUTENCAO"):
        tabelaFiltroCD = np.array([registro for registro in tabelaFiltroUndNegoc if registro[colOrgOrigem] != "CDC" and registro[colTipoOrdem] == "Entrega inbound planejada" or registro[colTipoOrdem] == "Ordem planejada"]);
    elif (undNegoc == "HELICOPTEROS"):
        tabelaFiltroCD = np.array([registro for registro in tabelaFiltroUndNegoc if registro[colOrgOrigem] != "CDJ" and registro[colTipoOrdem] == "Entrega inbound planejada" or registro[colTipoOrdem] == "Ordem planejada"]);

    tabelaFiltroPrefixo = np.array([registro for registro in tabelaFiltroCD if registro[colPrefixo] != "CONSUMO" and registro[colPrefixo] != "UNIFORME"]);

    for registro in tabelaFiltroPrefixo:
        if (type(registro[12]) != str):
            registro[12] = "";

    tabelaTop2Filtrada = pd.DataFrame(tabelaFiltroPrefixo, columns=nomesColunas);

    tabelaTop2Filtrada["PN&PREFIXO&QTY&TIPODEORDEM"] = tabelaTop2Filtrada["PN"] + tabelaTop2Filtrada["PREFIXO"].apply(str) + tabelaTop2Filtrada["QTDE"].apply(int).apply(str) + tabelaTop2Filtrada["TIPO ORDEM"];

    tabelaTop2Filtrada["Problema Raiz"] = "";

    tabelaTop2Filtrada["Ação"] = "";

    tabelaTop2Filtrada["OBS"] = "";

    tabelaTop2Filtrada["PN s/C"] = tabelaTop2Filtrada["PN"].str[:-2];

    tabelaTop2Filtrada["Condição"] = tabelaTop2Filtrada["PN"].str[-2:];
        
    tabelaTop2Filtrada["Desc."] = tabelaTop2Filtrada["DESCRIÇÃO"];

    with pd.ExcelWriter(wbLiberacaoCaminho, mode="a", if_sheet_exists="overlay") as writer:
        tabelaTop2Filtrada.to_excel(excel_writer=writer, sheet_name=wsTopico2Nome, index=False);

def insereValidacaoProbRaiz(wbLiberacaoCaminho, wsTopicoNome, listaTopExist, numLinhasT1, numLinhasT2):
    wbLiberacao = load_workbook(wbLiberacaoCaminho);

    wsTopico = wbLiberacao[wsTopicoNome];
    
    wsApoio = wbLiberacao["Apoio"];

    print("Chegou no apoio!");

    validacao_dados = DataValidation(type="list", formula1=f"=Apoio!$A$2:$A${wsApoio.max_row}", allow_blank=True);

    wsTopico.add_data_validation(validacao_dados);
    
    print(numLinhasT1);
    print(numLinhasT2);

    if (listaTopExist[0] and listaTopExist[1]):
        wsTopico.insert_rows(1, 4);
        wsTopico.cell(row=1, column=1, value="FOLLOW UP DE PLANEJAMENTO, COMPRAS, REPAROS E LOGÍSTICA");
        wsTopico.cell(row=2, column=1, value=date.today().strftime("%d/%m/%Y"));
        wsTopico.cell(row=4, column=1, value="TOPICO 1 - DEMANDAS SEM SUPRIMENTOS");
        linFimTop1 = numLinhasT1+5;
        linIniTop2 = numLinhasT1+10;
        linFimTop2 = 10+numLinhasT1+numLinhasT2;
        for i in range(len(listaTopExist)):
            if (numLinhasT1 > 1 and i == 0):
                validacao_dados.add(f"O6:O{numLinhasT1+5}");
            elif (numLinhasT2 > 1 and i == 1):
                wsTopico.cell(row=numLinhasT1+9, column=1, value="TOPICO 2 - MATERIAL NÃO LIBERADO PARA COMPRA OU TRANSFERÊNCIA");
                validacao_dados.add(f"R{numLinhasT1+11}:R{numLinhasT2+numLinhasT1+9}")
    else:
        if listaTopExist[0]:
            wsTopico.insert_rows(1, 4);
            wsTopico.cell(row=1, column=1, value="FOLLOW UP DE PLANEJAMENTO, COMPRAS, REPAROS E LOGÍSTICA");
            wsTopico.cell(row=2, column=1, value=date.today().strftime("%d/%m/%Y"));
            wsTopico.cell(row=4, column=1, value="TOPICO 1 - DEMANDAS SEM SUPRIMENTOS");
            validacao_dados.add(f"O6:O{wsTopico.max_row}");
        elif listaTopExist[1]:
            wsTopico.insert_rows(1, 4);
            wsTopico.cell(row=1, column=1, value="FOLLOW UP DE PLANEJAMENTO, COMPRAS, REPAROS E LOGÍSTICA");
            wsTopico.cell(row=2, column=1, value=date.today().strftime("%d/%m/%Y"));
            wsTopico.cell(row=1, column=1, value="TOPICO 2 - MATERIAL NÃO LIBERADO PARA COMPRA OU TRANSFERÊNCIA");
            validacao_dados.add(f"R3:R{wsTopico.max_row}");
    
    wbLiberacao.save(wbLiberacaoCaminho);
    wbLiberacao.close();
    return linFimTop1, linIniTop2, linFimTop2;

def uneTopicos(wbLiberacaoCaminho, wsTopico1Nome, wsTopico2Nome):
    wbLiberacao = load_workbook(wbLiberacaoCaminho);
    wsTopico1 = wbLiberacao[wsTopico1Nome];
    wsTopico2 = wbLiberacao[wsTopico2Nome];
    numLinhasT1 = wsTopico1.max_row;
    print(numLinhasT1)
    numLinhasT2 = wsTopico2.max_row;
    print(numLinhasT2)
    linhaInsTop2 = 5 + numLinhasT1
    for linha in range(1, numLinhasT2 + 1):
        linhaInsTop2 += 1;
        for coluna in range(1, wsTopico2.max_column + 1):
            wsTopico1.cell(linhaInsTop2, coluna).value = wsTopico2.cell(linha, coluna).value;

    wsTopico1.title = dataHoje;
    wbLiberacao.remove(wsTopico2);
    wbLiberacao.save(wbLiberacaoCaminho);
    wbLiberacao.close();
    return numLinhasT1, numLinhasT2;

def aplicaEstilo(wbLiberacaoCaminho, wsTopicoNome, linFimTop1, linIniTop2, linFimTop2, listaTopExist):
    wbLiberacao = load_workbook(wbLiberacaoCaminho);
    wsTopico = wbLiberacao[wsTopicoNome];
    estiloTabela = TableStyleInfo(name="TableStyleMedium19", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=True);
    if (listaTopExist[0] and listaTopExist[1]):
        wsTopico.add_table(Table(displayName=f"Tópico1_Planilha{wbLiberacao.worksheets.index(wsTopico)}", ref=f"A5:T{linFimTop1}", tableStyleInfo=estiloTabela));
        wsTopico.add_table(Table(displayName=f"Tópico2_Planilha{wbLiberacao.worksheets.index(wsTopico)}", ref=f"A{linIniTop2}:T{linFimTop2}", tableStyleInfo=estiloTabela));


if (listaTopExist[0] and listaTopExist[1]):
    
    filtraTopico1(wbLiberacaoCaminho, wsTopico1Nome);
    filtraTopico2(wbLiberacaoCaminho, wsTopico2Nome);
    numLinhasT1, numLinhasT2 = uneTopicos(wbLiberacaoCaminho, wsTopico1Nome, wsTopico2Nome);
    linFimTop1, linIniTop2, linFimTop2 = insereValidacaoProbRaiz(wbLiberacaoCaminho, dataHoje, listaTopExist, numLinhasT1, numLinhasT2);
    aplicaEstilo(wbLiberacaoCaminho, dataHoje, linFimTop1, linIniTop2, linFimTop2, listaTopExist);

else:

    i = 1;

    while i <= len(listaTopicos):

        topicoIdx = i-1;

        if listaTopExist[topicoIdx] and i == 1:

            filtraTopico1(wbLiberacaoCaminho, wsTopico1Nome);
            wsTopicoNome = wsTopico1Nome;

        elif listaTopExist[topicoIdx] and i == 2:

            filtraTopico2(wbLiberacaoCaminho, wsTopico2Nome);
            wsTopicoNome = wsTopico2Nome;

        i +=1;

    insereValidacaoProbRaiz(wbLiberacaoCaminho, wsTopicoNome, listaTopExist, numLinhasT1, numLinhasT2);


        


