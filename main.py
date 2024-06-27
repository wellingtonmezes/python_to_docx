import os
import requests
from docxtpl import DocxTemplate, InlineImage  # type: ignore
from docx.shared import Mm, Pt
from docx import Document
from htmldocx import HtmlToDocx
from bs4 import BeautifulSoup

dict_spec = {
    "Sistema Operacional": "<p>Windows 11 Home 64-Bits</p>",
    "Memória": "<p>16 GB RAM DDR5 de até 4800 MHz (8 GB em módulo SO-DIMM + 8 GB em módulo SO-DIMM)</p><p>Expansível até 32 GB DDR5 de até 4800 MHz (2 Slots SO-DIMM com capacidade para até 16 GB cada)</p>",
    "Processador e Chipset": "<p>Intel® Core™ i7-13700HX</p><p>16 núcleos (8 P-cores e 8 E-cores)</p><p>24 threads</p><p>Frequência: até 5.00 GHz</p><p>30 MB Intel® Smart Cache</p><p>Chipset: HM770</p><p>Para maiores informações consultar o fabricante</p>",
    "Tela": "<p>16” LED com design ultrafino</p><p>Painel: IPS (in-Plane-Switching)</p><p>Resolução: WUXGA (Wide Ultra Extended Graphics Array) 1920 x 1200</p><p>Proporção: 16:10</p><p>Taxa de atualização: 165 Hz</p><p>Tempo de resposta: 9 ms</p><p>Brilho: 400 nits</p><p>Taxa de contraste: 1000</p><p>Espaço de cor (color gamut): sRGB 100%</p><p>Tecnologia antirreflexo Acer ComfyView™</p>",
    "Gráficos": "<p>Nvidia® GeForce® RTX 4050 com 6 GB de memória dedicada GDDR6 (TGP de 120W)</p><p>Suporte as tecnologias: NVIDIA® GeForce® Experience, 2nd Gen Ray Tracing Cores, 3rd Gen Tensor Cores, Microsoft DirectX® 12 Ultimate, OpenGL 4.6</p><p>NVIDIA® Dynamic Boost 2.0, Game Ready Drivers e DLSS 3</p><p>UHD para processadores Intel® com memória compartilhada com a memória RAM.</p>",
    "Áudio e Microfone": "<p>Alto-falantes duplos estéreo Acer TrueHarmony</p><p>Tecnologia DTS® X: Ultra Áudio</p><p>Suportado no Windows Spatial Sound para PC Gaming, com licença DTS Integrada</p><p>Renderização de áudio imersiva em fones de ouvido e alto-falantes internos</p><p>Microfone duplo</p><p>Acer Purified Voice</p><p>Compativel com Cortana com Voz</p>",
    "Armazenamento": "<p>512 GB SSD NVMe PCIe 4.0 x4 M.2 2280</p>",
    "Upgrades": "<p>Este modelo possuí capacidade para a instalação e/ou melhorias de SSDs NVMe:</p><p></p><p>Slot dedicado ocupado M.2 2280, compatível com barramento PCIe 4.0 x4 NVMe de até 2TB. (Não acompanha o produto)</p><p>Slot dedicado livre M.2 2280, compatível com barramento PCIe 4.0 x4 NVMe de até 2TB. (Não acompanha o produto)</p>",
    "Webcam": "<p>Webcam com resolução HD (1280 x 720) e gravação de áudio e vídeo em 720p a 30 FPS com tecnologia temporal noise reduction (TNR)</p>",
    "Wi-Fi e Rede": "<p>Wireless / Wi-Fi rede sem fio:</p><p></p><p>Killer™ Wi-Fi 6 AX 1650i</p><p>802.11 a/b/g/n/ac R2 + ax wireless</p><p>Dual band (2.4 GHz e 5 GHz)</p><p>Suporte ao Wi-Fi 6</p><p>Com tecnologia 2x2 MU-MIMO</p><p>Suporte ao Bluetooth® 5.1 ou superior</p><p>LAN / RJ-45 rede com fio:</p><p></p><p>Killer™ Ethernet E2600</p><p>10/100/1000 Mbps</p><p>Suporte ao modo hibernação</p><p>Suporte ao Wake On Lan</p><p>Suporte ao IPv4 (32 Bits) e IPv6 (128 Bits)</p>",
    "Controle": "<p>Senha para BIOS, HDD e Solução TPM em Firmware (fTPM)</p>",
    "Dimensões e Peso": "<p>Sem caixa:</p><p>360.1 (L) x 279.9 (P) x 28.25 (A) mm</p><p>2,6 kg</p><p>*Com caixa (Aproximadamente~):</p><p>540 (L) x 360 (P) x 85 (A) mm</p><p>4,7 kg</p>",
    "Bateria e Alimentação": "<p>Fonte de alimentação:</p><p>Adaptador AC bivolt de 3 pinos (230W) com cabo e certificação do INMETRO</p><p>Bateria:</p><p>Bateria de 4 células (Li-Íon) 90Wh</p><p>Autonomia de até 7 horas (dependendo das condições de uso)</p>",
    "Teclados e  Touchpad": "<p>Teclado:</p><p>Membrana em português do Brasil padrão (ABNT 2) retroiluminado RGB com 4 zonas de iluminação</p><p>Atalho multimídia e funções (Tecla FN) + (Play, pause, parar, voltar, avançar, aumentar volume, diminuir volume, mudo e etc)</p><p>Teclado numérico independente</p><p>Tecla de atalho Predator Sense</p><p>Tecla Turbo para controle dos modos variados de trabalho da ventoinha</p><p>Touchpad:</p><p>Multi gestual com dois botões suportando rolagem com dois dedos, gestos para abrir Cortana, Action Center, multitarefa e comandos de aplicativos</p><p>Resistente a umidade</p><p>Certificação Microsoft Precision Touchpad</p>",
    "Aplicativos": "<p>PredatorSense</p>",
    "Conteúdo da Embalagem": "",
    "Cor": "<p>Preto</p>",
    "Acer PN": "<p>NH.QN7AL.003</p>",
    "Garantia": "<p>1 ano</p>",
    "Observações do Produto": "<p>Este produto não possuí leitor de CD/DVD</p><p>Não é possível realizar upgrades de GPU e CPU pois são soldados na placa mãe</p>",
    "EAN": "<p>4711121690617</p>",
    "Saiba Mais": "<p>195133216838</p>",
    "Código ANATEL": "<p>069701804423 (Referente ao módulo de Wi-Fi)</p>",
}

overview = "<p>Característica 1</p><p>Característica 1</p><p>Característica 1</p><p>Característica 1</p>"

PATH_OUT_DIR = f"{os.path.dirname(os.path.abspath(__file__))}\\tmp"
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_PATH, "tmp")
IMAGE_URL = "https://www.parceirosacer.com.br/Conteudos-Especiais/Notebooks/Predator-Helios-Neo/PHN16-71-76PL/PHN16-71-76PL.jpg"


class SpecificationToDocx:

    _list_specifications = list()
    _context = {"content_specs": []}
    _templates_dir = os.path.join(BASE_PATH, "templates")
    _temp_dir = os.path.join(BASE_PATH, "tmp")

    @classmethod
    def execute(
        cls,
        specifications: dict,
        product_type: str,
        product_id: str,
        product_model: str,
        product_family: str,
        # product_image_name: str,
    ):
        template = DocxTemplate(
            os.path.join(cls._templates_dir, f"tpl_bg_{product_type.lower()}.docx")
        )
        specification_to_list = [(k, v) for k, v in specifications.items()]

        for spec in specification_to_list:
            sub_docx = Document()
            new_parser = HtmlToDocx()
            new_parser.add_html_to_document(spec[1], sub_docx)
            sub_docx.styles["Normal"].font.size = Pt(8)
            sub_docx.save(f"{spec[0]}.docx")

        for spec in specification_to_list:
            sub_doc = template.new_subdoc(f"{spec[0]}.docx")
            cls._context["content_specs"].append({"label": spec[0], "value": sub_doc})

        sub_docx_overview = Document()
        overview_parser = HtmlToDocx()
        overview_parser.add_html_to_document(overview, sub_docx_overview)
        sub_docx_overview.styles["Normal"].font.size = Pt(22)
        sub_docx_overview.save(f"overview.docx")

        sub_doc_overview = template.new_subdoc("overview.docx")

        # image_bg = InlineImage(
        #     template,
        #     image_descriptor=os.path.join(cls._temp_dir, product_image_name),
        #     width=Mm(125),
        #     height=Mm(125),
        # )

        # cls._context["image_bg"] = image_bg

        # cls._context["content_specs"] = cls._list_specifications
        cls._context["model"] = product_model
        cls._context["family"] = product_family
        cls._context["overview"] = sub_doc_overview

        template.render(cls._context)
        template.save(
            os.path.join(cls._temp_dir, f"bg_{product_model}_{product_id}.docx")
        )

        BASE_DIR = os.getcwd()

        for file in os.listdir():
            if file.endswith("docx"):
                os.remove(os.path.join(BASE_DIR, file))


# hero_image = requests.get(IMAGE_URL)

# with open(os.path.join(TEMP_DIR, "PHN16-71-76PL.jpg"), "wb") as image:
#     image.write(hero_image.content)


SpecificationToDocx.execute(
    dict_spec, "NOTEBOOK", "6662662662", "PHN16-71-76PL", "Predator Helios Neo"
)

# {%tr for item in content_specs %}
# {{ item.label }}	{{ item.value }}
# {%tr endfor %}
