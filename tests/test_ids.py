"""Leitura e validação dos IDs Lattes (.txt, .csv e .xlsx)."""

import openpyxl
import pytest
from falsos import ID_A, ID_B, ID_C

import coletor


# --- normalizar_ids -----------------------------------------------------------

def test_aceita_id_puro_e_link():
    ids, invalidos, duplicados = coletor.normalizar_ids(
        [ID_A, f"http://lattes.cnpq.br/{ID_B}", f"https://lattes.cnpq.br/{ID_C}/"])
    assert ids == [ID_A, ID_B, ID_C]
    assert invalidos == []
    assert duplicados == 0


def test_remove_duplicados_mantendo_a_ordem():
    ids, _, duplicados = coletor.normalizar_ids([ID_B, ID_A, ID_B, f"lattes.cnpq.br/{ID_A}"])
    assert ids == [ID_B, ID_A]
    assert duplicados == 2


def test_ignora_vazios_e_separa_invalidos():
    ids, invalidos, _ = coletor.normalizar_ids(["", "   ", "abc", " 123 ", ID_A])
    assert ids == [ID_A]
    assert invalidos == ["abc", "123"]


@pytest.mark.parametrize("valor", [
    "12345678901234567",       # 17 dígitos
    "123456789012345",         # 15 dígitos
    "1234 5678 9012 3456",     # com espaços
    "1,23457E+15",             # notação científica do Excel
])
def test_rejeita_o_que_nao_e_id_de_16_digitos(valor):
    ids, invalidos, _ = coletor.normalizar_ids([valor])
    assert ids == []
    assert invalidos == [valor]


def test_aceita_id_que_virou_float_no_excel():
    assert coletor.normalizar_ids([f"{ID_A}.0"])[0] == [ID_A]


# --- .txt ---------------------------------------------------------------------

@pytest.mark.parametrize("codificacao", ["utf-8", "utf-8-sig", "cp1252"])
def test_le_txt_nas_codificacoes_comuns(tmp_path, codificacao):
    arquivo = tmp_path / "ids.txt"
    arquivo.write_text(f"Pesquisadores da pós-graduação\r\n{ID_A}\r\n\r\nhttp://lattes.cnpq.br/{ID_B}\r\n",
                       encoding=codificacao)
    linhas = coletor.ler_ids_txt(str(arquivo))
    assert linhas[0] == "Pesquisadores da pós-graduação"
    assert coletor.normalizar_ids(linhas)[0] == [ID_A, ID_B]


# --- .csv ---------------------------------------------------------------------

def _escrever(caminho, conteudo, codificacao="utf-8-sig"):
    caminho.write_text(conteudo, encoding=codificacao, newline="")
    return str(caminho)


def _por_nome(colunas):
    return {c.nome: c.valores for c in colunas}


def test_csv_do_excel_em_portugues_com_ponto_e_virgula(tmp_path):
    conteudo = f"Nome;ID Lattes\r\nSILVA, Maria;{ID_A}\r\nSOUZA, João;{ID_B}\r\n"
    arquivo = _escrever(tmp_path / "ids.csv", conteudo)
    assert _por_nome(coletor.ler_colunas(arquivo)) == {
        "Nome": ["SILVA, Maria", "SOUZA, João"],
        "ID Lattes": [ID_A, ID_B],
    }


@pytest.mark.parametrize("separador", [",", "\t"])
def test_csv_com_virgula_ou_tabulacao(tmp_path, separador):
    arquivo = _escrever(tmp_path / "ids.csv", f"nome{separador}id\nMaria{separador}{ID_A}\n")
    assert _por_nome(coletor.ler_colunas(arquivo))["id"] == [ID_A]


def test_csv_sem_cabecalho_nao_perde_a_primeira_linha(tmp_path):
    arquivo = _escrever(tmp_path / "ids.csv", f"{ID_A}\n{ID_B}\n")
    colunas = coletor.ler_colunas(arquivo)
    assert [(c.nome, c.valores) for c in colunas] == [("Coluna A", [ID_A, ID_B])]


def test_formato_nao_suportado(tmp_path):
    with pytest.raises(ValueError, match="não suportado"):
        coletor.ler_colunas(_escrever(tmp_path / "ids.ods", "x"))


# --- .xlsx --------------------------------------------------------------------

def _salvar_xlsx(caminho, linhas):
    wb = openpyxl.Workbook()
    for linha in linhas:
        wb.active.append(linha)
    wb.save(caminho)
    return str(caminho)


def test_xlsx_com_cabecalho_e_coluna_vazia(tmp_path):
    arquivo = _salvar_xlsx(tmp_path / "ids.xlsx", [
        ["Nome", None, "ID"],
        ["Maria", None, ID_A],
        ["João", None, ID_B],
    ])
    colunas = coletor.ler_colunas(arquivo)
    assert [c.nome for c in colunas] == ["Nome", "ID"]
    assert colunas[1].valores == [ID_A, ID_B]
    assert not colunas[1].tem_numeros_grandes


def test_xlsx_avisa_quando_o_id_foi_salvo_como_numero(tmp_path):
    arquivo = _salvar_xlsx(tmp_path / "ids.xlsx", [["ID"], [int(ID_A)], [ID_B]])
    (coluna,) = coletor.ler_colunas(arquivo)
    assert coluna.tem_numeros_grandes
    assert coluna.valores == [ID_A, ID_B]


def test_xlsx_vazio(tmp_path):
    assert coletor.ler_colunas(_salvar_xlsx(tmp_path / "vazio.xlsx", [])) == []
