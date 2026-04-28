from datetime import datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill


TAGS_MAP = {
    # Datas e identificadores
    "nDPS": "nDPS",
    "dhEmi": "dhEmi",
    "dCompet": "dCompet",
    # Emitente
    "emitente_cnpj": "emit/CNPJ",
    "emitente_nome": "emit/xNome",
    # Tomador
    "toma_cpf": "toma/CPF",
    "toma_cnpj": "toma/CNPJ",
    "toma_nome": "toma/xNome",
    # Servico
    "cTribNac": "cServ/cTribNac",
    "xDescServ": "cServ/xDescServ",
    "cNBS": "cServ/cNBS",
    "cIndOp": "IBSCBS/cIndOp",
    "CST": "gIBSCBS/CST",
    "cClassTrib": "gIBSCBS/cClassTrib",
    # Valores da NFS-e
    "vServ": "valores/vServPrest/vServ",
    "vCalcDR": "infNFSe/valores/vCalcDR",
    "vBC_ISSQN": "infNFSe/valores/vBC",
    "pAliqISSQN": "infNFSe/valores/pAliqAplic",
    "vISSQN": "infNFSe/valores/vISSQN",
    "vLiq": "infNFSe/valores/vLiq",
    # IBS/CBS
    "vBC_IBSCBS": "IBSCBS/valores/vBC",
    "pIBSUF": "IBSCBS/valores/uf/pIBSUF",
    "pRedAliqUF": "IBSCBS/valores/uf/pRedAliqUF",
    "pAliqEfetUF": "IBSCBS/valores/uf/pAliqEfetUF",
    "pIBSMun": "IBSCBS/valores/mun/pIBSMun",
    "pRedAliqMun": "IBSCBS/valores/mun/pRedAliqMun",
    "pAliqEfetMun": "IBSCBS/valores/mun/pAliqEfetMun",
    "pCBS": "IBSCBS/valores/fed/pCBS",
    "pRedAliqCBS": "IBSCBS/valores/fed/pRedAliqCBS",
    "pAliqEfetCBS": "IBSCBS/valores/fed/pAliqEfetCBS",
    "vTotNF": "IBSCBS/totCIBS/vTotNF",
    "vIBSTot": "IBSCBS/totCIBS/gIBS/vIBSTot",
    "vIBSUF": "IBSCBS/totCIBS/gIBS/gIBSUFTot/vIBSUF",
    "vIBSMun": "IBSCBS/totCIBS/gIBS/gIBSMunTot/vIBSMun",
    "vCBS": "IBSCBS/totCIBS/gCBS/vCBS",
}


COLUNAS_ORDEM = [
    "arquivo_zip",
    "arquivo",
    "nDPS",
    "dhEmi",
    "dCompet",
    "emitente_cnpj",
    "emitente_nome",
    "toma_documento",
    "toma_cpf",
    "toma_cnpj",
    "toma_nome",
    "cTribNac",
    "xDescServ",
    "cNBS",
    "cIndOp",
    "CST",
    "cClassTrib",
    "vServ",
    "vCalcDR",
    "vBC_ISSQN",
    "pAliqISSQN",
    "vISSQN",
    "vLiq",
    "vBC_IBSCBS",
    "pIBSUF",
    "pRedAliqUF",
    "pAliqEfetUF",
    "pIBSMun",
    "pRedAliqMun",
    "pAliqEfetMun",
    "pCBS",
    "pRedAliqCBS",
    "pAliqEfetCBS",
    "vTotNF",
    "vIBSTot",
    "vIBSUF",
    "vIBSMun",
    "vCBS",
]


NOMES_COLUNAS_EXCEL = {
    "arquivo_zip": "Arquivo ZIP",
    "arquivo": "Arquivo XML",
    "nDPS": "Numero DPS",
    "dhEmi": "Data Emissao",
    "dCompet": "Data Competencia",
    "emitente_cnpj": "CNPJ Emitente",
    "emitente_nome": "Empresa Emitente",
    "toma_documento": "Documento Tomador",
    "toma_cpf": "CPF Tomador",
    "toma_cnpj": "CNPJ Tomador",
    "toma_nome": "Nome Tomador",
    "cTribNac": "Codigo Tributacao Nacional",
    "xDescServ": "Descricao Servico XML",
    "cNBS": "NBS XML",
    "cIndOp": "Indicador Operacao XML",
    "CST": "CST XML",
    "cClassTrib": "Classificacao Tributaria XML",
    "vServ": "Valor Servico",
    "vCalcDR": "Valor Deducoes/Reducoes",
    "vBC_ISSQN": "Base ISSQN",
    "pAliqISSQN": "Aliquota ISSQN",
    "vISSQN": "Valor ISSQN",
    "vLiq": "Valor Liquido",
    "vBC_IBSCBS": "Base IBS/CBS",
    "pIBSUF": "Aliquota IBS UF",
    "pRedAliqUF": "Reducao IBS UF",
    "pAliqEfetUF": "Aliquota Efetiva IBS UF",
    "pIBSMun": "Aliquota IBS Municipio",
    "pRedAliqMun": "Reducao IBS Municipio",
    "pAliqEfetMun": "Aliquota Efetiva IBS Municipio",
    "pCBS": "Aliquota CBS",
    "pRedAliqCBS": "Reducao CBS",
    "pAliqEfetCBS": "Aliquota Efetiva CBS",
    "vTotNF": "Valor Total NF",
    "vIBSTot": "Valor IBS Total",
    "vIBSUF": "Valor IBS UF",
    "vIBSMun": "Valor IBS Municipio",
    "vCBS": "Valor CBS",
    "empresa_identificada": "Empresa Identificada",
    "servico_descricao_match": "Servico Encontrado na Descricao",
    "criterio_match": "Criterio de Comparacao",
    "status_comparacao": "Status Comparacao",
    "detalhe_comparacao": "Detalhe Comparacao",
    "servico_base": "Servico Base",
    "servico_chave": "Servico Chave",
    "tipo_atividade_base": "Tipo Atividade Base",
    "cnae_base": "CNAE Base",
    "nbs_base": "NBS Base",
    "CST_base": "CST Base",
    "cClassTrib_base": "Classificacao Tributaria Base",
    "cIndOp_base": "Indicador Operacao Base",
    "pRedAliqUF_base": "Reducao IBS UF Base",
    "pRedAliqMun_base": "Reducao IBS Municipio Base",
    "pRedAliqCBS_base": "Reducao CBS Base",
    "reducao_aliquota_percentual_base": "Reducao Aliquota Base",
}


COLUNAS_BASE = [
    "emitente_cnpj",
    "servico_base",
    "servico_chave",
    "tipo_atividade_base",
    "cnae_base",
    "nbs_base",
    "CST_base",
    "cClassTrib_base",
    "cIndOp_base",
    "pRedAliqUF_base",
    "pRedAliqMun_base",
    "pRedAliqCBS_base",
    "reducao_aliquota_percentual_base",
]


def remover_namespace(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def texto_limpo(elemento: ET.Element | None) -> str | None:
    if elemento is None or elemento.text is None:
        return None

    texto = elemento.text.strip()
    return texto or None


def encontrar_filho_por_nome(elemento: ET.Element, nome_tag: str) -> ET.Element | None:
    for filho in list(elemento):
        if remover_namespace(filho.tag) == nome_tag:
            return filho
    return None


def extrair_por_caminho(root: ET.Element, caminho: str) -> str | None:
    partes = [parte for parte in caminho.split("/") if parte]
    if not partes:
        return None

    for candidato in root.iter():
        if remover_namespace(candidato.tag) != partes[0]:
            continue

        atual = candidato
        for parte in partes[1:]:
            atual = encontrar_filho_por_nome(atual, parte)
            if atual is None:
                break

        valor = texto_limpo(atual)
        if valor is not None:
            return valor

    return None


def somente_digitos(valor) -> str | None:
    if valor is None or pd.isna(valor):
        return None

    digitos = re.sub(r"\D", "", str(valor))
    return digitos or None


def normalizar_cnpj(valor) -> str | None:
    digitos = somente_digitos(valor)
    if not digitos:
        return None

    return digitos.zfill(14)


def normalizar_codigo(valor, tamanho: int | None = None) -> str | None:
    digitos = somente_digitos(valor)
    if not digitos:
        return None

    return digitos.zfill(tamanho) if tamanho else digitos


def normalizar_percentual(valor) -> str | None:
    if valor is None or pd.isna(valor):
        return None

    texto = str(valor).replace("%", "").strip()
    if not texto:
        return None

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    else:
        texto = texto.replace(",", ".")

    try:
        return str(Decimal(texto).quantize(Decimal("0.01")))
    except InvalidOperation:
        return None


def percentuais_iguais(valor_xml, valor_base) -> bool:
    base = normalizar_percentual(valor_base)
    if base is None:
        return True

    xml = normalizar_percentual(valor_xml)
    return xml == base


def normalizar_texto_busca(valor) -> str:
    if valor is None or pd.isna(valor):
        return ""

    texto = str(valor).replace("\xa0", " ").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(char for char in texto if not unicodedata.combining(char))
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def servico_descricao_compativel(descricao_xml, servico_base) -> bool:
    descricao = normalizar_texto_busca(descricao_xml)

    if not descricao or not normalizar_texto_busca(servico_base):
        return False

    # Permite cadastrar sinonimos na base separados por ; ou |.
    # Exemplo: "Diarias; Hospedagem" casa com qualquer um dos termos.
    termos = []
    for termo in re.split(r"[;|]", str(servico_base)):
        termo_normalizado = normalizar_texto_busca(termo)
        if termo_normalizado:
            termos.append(termo_normalizado)

    stopwords = {"a", "ao", "aos", "as", "da", "das", "de", "do", "dos", "e", "em", "para"}
    descricao_tokens = set(descricao.split())

    for termo in termos:
        if re.search(rf"\b{re.escape(termo)}\b", descricao):
            return True

        termo_tokens = [token for token in termo.split() if token not in stopwords]
        if termo_tokens and all(token in descricao_tokens for token in termo_tokens):
            return True

    return False


def normalizar_cclass_trib(valor) -> str | None:
    digitos = somente_digitos(valor)
    if not digitos:
        return None

    # Na planilha, codigos como 200048 podem virar 20048 quando o Excel trata
    # como numero. Nesse caso, o zero perdido fica depois dos tres primeiros
    # digitos do grupo 200.
    if len(digitos) == 5 and digitos.startswith("200"):
        return f"{digitos[:3]}{digitos[3:].zfill(3)}"

    return digitos.zfill(6)


def formatar_data_iso(valor: str | None) -> str | None:
    if not valor:
        return None

    valor = str(valor).strip()
    try:
        return datetime.fromisoformat(valor.replace("Z", "+00:00")).date().isoformat()
    except ValueError:
        pass

    try:
        return datetime.strptime(valor[:10], "%Y-%m-%d").date().isoformat()
    except ValueError:
        return valor


def parsear_xml_nfse(xml_bytes: bytes) -> dict[str, str | None]:
    root = ET.fromstring(xml_bytes)
    dados = {
        campo: extrair_por_caminho(root, caminho)
        for campo, caminho in TAGS_MAP.items()
    }

    dados["emitente_cnpj"] = normalizar_cnpj(dados.get("emitente_cnpj"))
    dados["toma_cpf"] = normalizar_codigo(dados.get("toma_cpf"), 11)
    dados["toma_cnpj"] = normalizar_cnpj(dados.get("toma_cnpj"))
    dados["toma_documento"] = dados.get("toma_cnpj") or dados.get("toma_cpf")
    dados["cNBS"] = normalizar_codigo(dados.get("cNBS"))
    dados["cIndOp"] = normalizar_codigo(dados.get("cIndOp"), 6)
    dados["CST"] = normalizar_codigo(dados.get("CST"), 3)
    dados["cClassTrib"] = normalizar_cclass_trib(dados.get("cClassTrib"))
    dados["dhEmi"] = formatar_data_iso(dados.get("dhEmi"))
    dados["dCompet"] = formatar_data_iso(dados.get("dCompet"))

    return dados


def ler_zip_nfse(arquivo_zip) -> tuple[pd.DataFrame, pd.DataFrame]:
    registros = []
    erros = []
    nome_zip = getattr(arquivo_zip, "name", None)

    if hasattr(arquivo_zip, "seek"):
        arquivo_zip.seek(0)

    with zipfile.ZipFile(arquivo_zip, "r") as zip_ref:
        for nome in zip_ref.namelist():
            if not nome.lower().endswith(".xml") or nome.endswith("/"):
                continue

            try:
                with zip_ref.open(nome) as arquivo_xml:
                    dados = parsear_xml_nfse(arquivo_xml.read())
                    dados["arquivo_zip"] = nome_zip
                    dados["arquivo"] = nome
                    registros.append(dados)
            except Exception as exc:
                erros.append({"arquivo_zip": nome_zip, "arquivo": nome, "erro": str(exc)})

    return pd.DataFrame(registros), pd.DataFrame(erros)


def ler_zips_nfse(arquivos_zip) -> tuple[pd.DataFrame, pd.DataFrame]:
    dfs_xml = []
    dfs_erros = []

    for arquivo_zip in arquivos_zip:
        df_xml, df_erros = ler_zip_nfse(arquivo_zip)
        if not df_xml.empty:
            dfs_xml.append(df_xml)
        if not df_erros.empty:
            dfs_erros.append(df_erros)

    df_xml_final = pd.concat(dfs_xml, ignore_index=True) if dfs_xml else pd.DataFrame()
    df_erros_final = pd.concat(dfs_erros, ignore_index=True) if dfs_erros else pd.DataFrame()
    return df_xml_final, df_erros_final


def organizar_colunas(df: pd.DataFrame, colunas_ordem: list[str]) -> pd.DataFrame:
    for coluna in colunas_ordem:
        if coluna not in df.columns:
            df[coluna] = None

    extras = [coluna for coluna in df.columns if coluna not in colunas_ordem]
    return df[colunas_ordem + extras]


def linha_tem_cnpj(row: pd.Series) -> bool:
    primeira = normalizar_cnpj(row.iloc[0])
    outras_preenchidas = any(pd.notna(valor) and str(valor).strip() for valor in row.iloc[1:])
    return bool(primeira) and not outras_preenchidas


def ler_base_conferencia(arquivo_excel) -> pd.DataFrame:
    excel = pd.ExcelFile(arquivo_excel)

    if "Base_Normalizada" in excel.sheet_names:
        return ler_base_normalizada(excel)

    if "Parametros" not in excel.sheet_names:
        raise ValueError("A base deve ter a aba 'Base_Normalizada' ou 'Parametros'.")

    bruto = pd.read_excel(excel, sheet_name="Parametros", header=None, dtype=str)
    registros = []
    cnpj_atual = None

    for _, row in bruto.iterrows():
        primeira = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""

        if not primeira:
            continue

        if linha_tem_cnpj(row):
            cnpj_atual = normalizar_cnpj(row.iloc[0])
            continue

        if primeira.lower().startswith("servi"):
            continue

        if not cnpj_atual:
            continue

        registros.append(
            {
                "emitente_cnpj": cnpj_atual,
                "servico_base": primeira.replace("\xa0", " ").strip(),
                "servico_chave": normalizar_texto_busca(primeira),
                "tipo_atividade_base": normalizar_codigo(row.iloc[1], 4),
                "cnae_base": normalizar_codigo(row.iloc[2]),
                "nbs_base": normalizar_codigo(row.iloc[3]),
                "CST_base": None,
                "cClassTrib_base": normalizar_cclass_trib(row.iloc[4]),
                "cIndOp_base": normalizar_codigo(row.iloc[5], 6),
                "pRedAliqUF_base": None,
                "pRedAliqMun_base": None,
                "pRedAliqCBS_base": None,
                "reducao_aliquota_percentual_base": None,
            }
        )

    return pd.DataFrame(registros, columns=COLUNAS_BASE)


def obter_coluna(row: pd.Series, *nomes):
    for nome in nomes:
        if nome in row.index:
            return row.get(nome)
    return None


def ler_base_normalizada(excel: pd.ExcelFile) -> pd.DataFrame:
    bruto = pd.read_excel(excel, sheet_name="Base_Normalizada", dtype=str)
    registros = []

    for _, row in bruto.iterrows():
        emitente_cnpj = normalizar_cnpj(obter_coluna(row, "emitente_cnpj"))
        servico_base = obter_coluna(row, "servico_base")

        if not emitente_cnpj or pd.isna(servico_base) or not str(servico_base).strip():
            continue

        reducao_generica = normalizar_percentual(
            obter_coluna(
                row,
                "reducao_aliquota_percentual_base",
                "reducao_percentual_base",
                "percentual_reducao_base",
                "pRedAliq_base",
            )
        )

        registros.append(
            {
                "emitente_cnpj": emitente_cnpj,
                "servico_base": str(servico_base).replace("\xa0", " ").strip(),
                "servico_chave": normalizar_texto_busca(servico_base),
                "tipo_atividade_base": normalizar_codigo(obter_coluna(row, "tipo_atividade_base"), 4),
                "cnae_base": normalizar_codigo(obter_coluna(row, "cnae_base")),
                "nbs_base": normalizar_codigo(obter_coluna(row, "nbs_base")),
                "CST_base": normalizar_codigo(obter_coluna(row, "CST_base", "cst_base"), 3),
                "cClassTrib_base": normalizar_cclass_trib(obter_coluna(row, "cClassTrib_base")),
                "cIndOp_base": normalizar_codigo(obter_coluna(row, "cIndOp_base"), 6),
                "pRedAliqUF_base": normalizar_percentual(obter_coluna(row, "pRedAliqUF_base")),
                "pRedAliqMun_base": normalizar_percentual(obter_coluna(row, "pRedAliqMun_base")),
                "pRedAliqCBS_base": normalizar_percentual(obter_coluna(row, "pRedAliqCBS_base")),
                "reducao_aliquota_percentual_base": reducao_generica,
            }
        )

    return pd.DataFrame(registros, columns=COLUNAS_BASE)


def tipo_atividade_compativel(c_trib_nac: str | None, tipo_base: str | None) -> bool:
    if not c_trib_nac or not tipo_base:
        return True

    return str(c_trib_nac).startswith(str(tipo_base))


def comparar_com_base(df_xml: pd.DataFrame, df_base: pd.DataFrame) -> pd.DataFrame:
    registros = []
    base_por_empresa = {
        cnpj: grupo.reset_index(drop=True)
        for cnpj, grupo in df_base.groupby("emitente_cnpj", dropna=False)
    }

    for _, xml in df_xml.iterrows():
        cnpj = xml.get("emitente_cnpj")
        base_empresa = base_por_empresa.get(cnpj)
        registro = xml.to_dict()

        if base_empresa is None or base_empresa.empty:
            registro.update(
                {
                    "empresa_identificada": "Nao",
                    "status_comparacao": "Empresa nao encontrada na base",
                    "detalhe_comparacao": "Nao existe bloco para este emitente_cnpj na base.",
                }
            )
            registros.append(registro)
            continue

        candidatos_servico = base_empresa[
            base_empresa["servico_base"].apply(
                lambda servico: servico_descricao_compativel(xml.get("xDescServ"), servico)
            )
        ]
        encontrou_servico = not candidatos_servico.empty

        candidatos = candidatos_servico[
            (candidatos_servico["nbs_base"] == xml.get("cNBS"))
            & (candidatos_servico["cClassTrib_base"] == xml.get("cClassTrib"))
            & (candidatos_servico["cIndOp_base"] == xml.get("cIndOp"))
        ]

        if candidatos.empty:
            candidatos = candidatos_servico[
                (candidatos_servico["cClassTrib_base"] == xml.get("cClassTrib"))
                & (candidatos_servico["cIndOp_base"] == xml.get("cIndOp"))
                & (
                    candidatos_servico["tipo_atividade_base"].apply(
                        lambda tipo: tipo_atividade_compativel(xml.get("cTribNac"), tipo)
                    )
                )
            ]

        if candidatos.empty:
            candidatos = candidatos_servico

        if candidatos.empty:
            candidatos = base_empresa[
            (base_empresa["nbs_base"] == xml.get("cNBS"))
            & (base_empresa["cClassTrib_base"] == xml.get("cClassTrib"))
            & (base_empresa["cIndOp_base"] == xml.get("cIndOp"))
            ]

        if candidatos.empty:
            candidatos = base_empresa[
                (base_empresa["nbs_base"] == xml.get("cNBS"))
                & (base_empresa["cClassTrib_base"] == xml.get("cClassTrib"))
            ]

        if candidatos.empty:
            candidatos = base_empresa[base_empresa["nbs_base"] == xml.get("cNBS")]

        if candidatos.empty:
            candidatos = base_empresa[
                (base_empresa["cClassTrib_base"] == xml.get("cClassTrib"))
                & (base_empresa["cIndOp_base"] == xml.get("cIndOp"))
                & (
                    base_empresa["tipo_atividade_base"].apply(
                        lambda tipo: tipo_atividade_compativel(xml.get("cTribNac"), tipo)
                    )
                )
            ]

        if candidatos.empty:
            registro.update(
                {
                    "empresa_identificada": "Sim",
                    "servico_descricao_match": "Nao",
                    "criterio_match": "Nenhum",
                    "status_comparacao": "Parametro nao encontrado",
                    "detalhe_comparacao": "Empresa existe, mas nao houve match por servico/NBS/cClassTrib/cIndOp.",
                }
            )
            registros.append(registro)
            continue

        base = candidatos.iloc[0].to_dict()
        divergencias = []
        servico_match = servico_descricao_compativel(xml.get("xDescServ"), base.get("servico_base"))

        if not servico_match:
            divergencias.append("Servico_descricao")
        if base.get("nbs_base") != xml.get("cNBS"):
            divergencias.append("NBS")
        if base.get("CST_base") and base.get("CST_base") != xml.get("CST"):
            divergencias.append("CST")
        if base.get("cClassTrib_base") != xml.get("cClassTrib"):
            divergencias.append("cClassTrib")
        if base.get("cIndOp_base") != xml.get("cIndOp"):
            divergencias.append("Ind_Operacao")
        if not tipo_atividade_compativel(xml.get("cTribNac"), base.get("tipo_atividade_base")):
            divergencias.append("Tipo_Atividade/cTribNac")

        reducao_generica = base.get("reducao_aliquota_percentual_base")
        if reducao_generica:
            if not percentuais_iguais(xml.get("pRedAliqUF"), reducao_generica):
                divergencias.append("Reducao_UF")
            if not percentuais_iguais(xml.get("pRedAliqMun"), reducao_generica):
                divergencias.append("Reducao_Mun")
            if not percentuais_iguais(xml.get("pRedAliqCBS"), reducao_generica):
                divergencias.append("Reducao_CBS")
        else:
            if not percentuais_iguais(xml.get("pRedAliqUF"), base.get("pRedAliqUF_base")):
                divergencias.append("Reducao_UF")
            if not percentuais_iguais(xml.get("pRedAliqMun"), base.get("pRedAliqMun_base")):
                divergencias.append("Reducao_Mun")
            if not percentuais_iguais(xml.get("pRedAliqCBS"), base.get("pRedAliqCBS_base")):
                divergencias.append("Reducao_CBS")

        registro.update(base)
        registro["empresa_identificada"] = "Sim"
        registro["servico_descricao_match"] = "Sim" if servico_match else "Nao"
        registro["criterio_match"] = "Servico_descricao" if encontrou_servico else "Codigos fiscais"
        registro["status_comparacao"] = "Divergencia" if divergencias else "OK"
        registro["detalhe_comparacao"] = ", ".join(divergencias) if divergencias else "Parametros conferem"
        registros.append(registro)

    colunas_comparacao = [
        "empresa_identificada",
        "servico_descricao_match",
        "criterio_match",
        "status_comparacao",
        "detalhe_comparacao",
        "servico_base",
        "servico_chave",
        "tipo_atividade_base",
        "cnae_base",
        "nbs_base",
        "CST_base",
        "cClassTrib_base",
        "cIndOp_base",
        "pRedAliqUF_base",
        "pRedAliqMun_base",
        "pRedAliqCBS_base",
        "reducao_aliquota_percentual_base",
    ]

    return organizar_colunas(pd.DataFrame(registros), COLUNAS_ORDEM + colunas_comparacao)


def preparar_excel(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns=NOMES_COLUNAS_EXCEL)


def formatar_abas_excel(writer: pd.ExcelWriter) -> None:
    preenchimento = PatternFill("solid", fgColor="000000")
    fonte = Font(color="FFFFFF", bold=True)

    for worksheet in writer.book.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for cell in worksheet[1]:
            cell.fill = preenchimento
            cell.font = fonte

        for column_cells in worksheet.columns:
            tamanho = max(len(str(cell.value or "")) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(
                max(tamanho + 2, 12),
                60,
            )


def gerar_excel(
    df_xml: pd.DataFrame,
    df_erros: pd.DataFrame,
    df_base: pd.DataFrame | None = None,
    df_comparacao: pd.DataFrame | None = None,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        preparar_excel(df_xml).to_excel(writer, index=False, sheet_name="XML_Extraido")

        if df_comparacao is not None:
            preparar_excel(df_comparacao).to_excel(writer, index=False, sheet_name="Comparacao")

        if df_base is not None:
            preparar_excel(df_base).to_excel(writer, index=False, sheet_name="Base_Normalizada")

        if not df_erros.empty:
            preparar_excel(df_erros).to_excel(writer, index=False, sheet_name="Erros_XML")

        formatar_abas_excel(writer)

    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="Processador de NFS-e IBS/CBS", layout="wide")
    st.title("Processador de NFS-e IBS/CBS")

    col_zip, col_base = st.columns(2)
    with col_zip:
        arquivos_zip = st.file_uploader(
            "ZIPs com XMLs de NFS-e",
            type=["zip"],
            accept_multiple_files=True,
        )
    with col_base:
        arquivo_base = st.file_uploader("Base de conferencia em Excel", type=["xlsx"])

    if not arquivos_zip:
        return

    try:
        with st.spinner("Processando XMLs..."):
            df_xml, df_erros = ler_zips_nfse(arquivos_zip)
            df_xml = organizar_colunas(df_xml, COLUNAS_ORDEM)

        df_base = None
        df_comparacao = None

        if arquivo_base:
            with st.spinner("Lendo base e comparando parametros..."):
                df_base = ler_base_conferencia(arquivo_base)
                df_comparacao = comparar_com_base(df_xml, df_base)

        st.success(
            f"Processamento concluido: {len(df_xml)} XML(s) lido(s) "
            f"em {len(arquivos_zip)} ZIP(s)."
        )

        aba_xml, aba_comparacao, aba_base, aba_erros = st.tabs(
            ["XML extraido", "Comparacao", "Base normalizada", "Erros"]
        )

        with aba_xml:
            st.dataframe(df_xml, use_container_width=True, hide_index=True)

        with aba_comparacao:
            if df_comparacao is None:
                st.info("Envie a base de conferencia para gerar esta aba.")
            else:
                st.dataframe(df_comparacao, use_container_width=True, hide_index=True)

        with aba_base:
            if df_base is None:
                st.info("Envie a base de conferencia para visualizar a normalizacao.")
            else:
                st.dataframe(df_base, use_container_width=True, hide_index=True)

        with aba_erros:
            if df_erros.empty:
                st.info("Nenhum erro de XML encontrado.")
            else:
                st.dataframe(df_erros, use_container_width=True, hide_index=True)

        st.download_button(
            label="Baixar Excel",
            data=gerar_excel(df_xml, df_erros, df_base, df_comparacao),
            file_name="nfse_conferencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except zipfile.BadZipFile:
        st.error("O arquivo enviado nao e um ZIP valido.")
    except ValueError as exc:
        st.error(f"Erro na base de conferencia: {exc}")
    except Exception as exc:
        st.error(f"Erro inesperado: {exc}")


if __name__ == "__main__":
    main()
