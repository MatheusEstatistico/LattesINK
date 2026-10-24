"""coletor.executar_lote: o laço que a janela roda numa thread separada."""

import csv
import threading

from falsos import ID_A, ID_B, ID_C, DriverFalso
from selenium.common.exceptions import WebDriverException

import coletor


class Mensagens(list):
    """Faz o papel da fila da janela."""

    def __call__(self, tipo, dado):
        self.append((tipo, dado))

    def do_tipo(self, tipo):
        return [dado for t, dado in self if t == tipo]


def test_lote_concluido(novo_coletor, tmp_path):
    driver = DriverFalso()
    c = novo_coletor(driver)
    mensagens = Mensagens()

    coletor.executar_lote(c, [ID_A, ID_B], str(tmp_path), mensagens, inicio=3, total=4)

    assert [r.id_lattes for r in mensagens.do_tipo("registro")] == [ID_A, ID_B]
    assert mensagens.do_tipo("log")[0] == f"🌐  [3/4]  Abrindo o currículo {ID_A}"
    assert mensagens[-1] == ("fim", ("concluido", ""))
    assert driver.fechado  # o navegador é fechado no fim

    with open(tmp_path / coletor.NOME_CSV, encoding="utf-8-sig") as f:
        assert [linha[1] for linha in csv.reader(f, delimiter=";")] == ["ID Lattes", ID_A, ID_B]


def test_lote_cancelado(novo_coletor, cancelado, tmp_path):
    cancelado.set()
    mensagens = Mensagens()
    coletor.executar_lote(novo_coletor(DriverFalso()), [ID_A], str(tmp_path), mensagens)
    assert mensagens == [("fim", ("cancelado", ""))]


def test_lote_sem_chrome(tmp_path):
    c = coletor.ColetorLattes(threading.Event(), log=lambda _: None)

    def sem_chrome():
        raise WebDriverException("unable to obtain driver for chrome")

    c._criar_driver = sem_chrome
    mensagens = Mensagens()
    coletor.executar_lote(c, [ID_A], str(tmp_path), mensagens)
    motivo, detalhe = mensagens[-1][1]
    assert motivo == "erro"
    assert detalhe.startswith("Não foi possível abrir o Chrome:\nunable to obtain driver for chrome")
    assert "Google Chrome está instalado" in detalhe


def test_csv_aberto_no_excel_nao_interrompe_o_lote(novo_coletor, tmp_path):
    (tmp_path / coletor.NOME_CSV).mkdir()  # abrir um diretório como arquivo dá OSError, como um CSV travado
    mensagens = Mensagens()
    coletor.executar_lote(novo_coletor(DriverFalso()), [ID_A], str(tmp_path), mensagens)
    assert any("CSV de backup" in m for m in mensagens.do_tipo("log"))
    assert len(mensagens.do_tipo("registro")) == 1
    assert mensagens[-1] == ("fim", ("concluido", ""))


def test_erro_inesperado_encerra_com_mensagem_e_vai_para_o_log(tmp_path, caplog):
    class Quebrado:
        cancelado = threading.Event()
        fechado = False

        def capturar(self, lattes_id, pasta):
            raise RuntimeError("falha inesperada")

        def fechar(self):
            self.fechado = True

    c = Quebrado()
    mensagens = Mensagens()
    coletor.executar_lote(c, [ID_A, ID_C], str(tmp_path), mensagens)
    assert mensagens[-1] == ("fim", ("erro", "Erro inesperado: falha inesperada"))
    assert c.fechado
    (registro,) = caplog.records
    assert registro.levelname == "ERROR"
    assert registro.exc_info[0] is RuntimeError  # o traceback completo vai para o log


# --- pausa depois de várias falhas seguidas --------------------------------------

class ColetorRoteirizado:
    """Devolve, para cada ID, um registro com o status do roteiro."""

    def __init__(self, status):
        self.status = list(status)
        self.cancelado = threading.Event()
        self.capturados = []

    def capturar(self, lattes_id, pasta):
        self.capturados.append(lattes_id)
        return coletor.Registro(lattes_id, self.status.pop(0))

    def fechar(self):
        pass


IDS = [f"{n:016d}" for n in range(1, 7)]
OK, SEM_DATA, CAPTCHA, PAGINA, ERRO = (coletor.STATUS_OK, coletor.STATUS_SEM_DATA, coletor.STATUS_CAPTCHA,
                                       coletor.STATUS_PAGINA_INESPERADA, coletor.STATUS_ERRO)


def test_pausa_depois_de_tres_falhas_seguidas(tmp_path):
    c = ColetorRoteirizado([OK, CAPTCHA, PAGINA, ERRO, OK, OK])
    mensagens = Mensagens()
    coletor.executar_lote(c, IDS, str(tmp_path), mensagens)
    assert c.capturados == IDS[:4]
    motivo, detalhe = mensagens[-1][1]
    assert motivo == "pausado"
    assert detalhe.startswith('3 currículos seguidos não abriram (o último: "Erro")')


def test_sucesso_ou_curriculo_sem_data_zeram_a_contagem(tmp_path):
    c = ColetorRoteirizado([CAPTCHA, CAPTCHA, OK, CAPTCHA, CAPTCHA, SEM_DATA])
    mensagens = Mensagens()
    coletor.executar_lote(c, IDS, str(tmp_path), mensagens)
    assert c.capturados == IDS
    assert mensagens[-1] == ("fim", ("concluido", ""))


def test_nao_pausa_se_as_falhas_sao_os_ultimos_ids(tmp_path):
    c = ColetorRoteirizado([OK, OK, OK, ERRO, ERRO, ERRO])
    mensagens = Mensagens()
    coletor.executar_lote(c, IDS, str(tmp_path), mensagens)
    assert mensagens[-1] == ("fim", ("concluido", ""))
