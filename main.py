import os
import requests
from docxtpl import DocxTemplate, InlineImage  # type: ignore
from docx.shared import Mm

dict_spec = {
    "Sistema Operacional": "Windows 11 Pro",
    "Processador e Chipset": "Intel Core i7-1355U",
    "Memória": "2GB",
    "Tela": "16” LED com design ultrafino",
    "Gráficos": "Nvidia® GeForce® RTX 4050 com 6 GB de memória dedicada GDDR6 (TGP de 120W)",
    "Áudio e Microfone": "Alto-falantes duplos estéreo Acer TrueHarmony",
    "Armazenamento": "512 GB SSD NVMe PCIe 4.0 x4 M.2 2280",
    "Upgrades": "Não",
    "Webcam": "Webcam com resolução HD (1280 x 720) e gravação de áudio e vídeo em 720p a 30 FPS",
    "Wi-Fi e Rede": "Dual band (2.4 GHz e 5 GHz)",
    "Controle": "Não",
    "Dimensões e Peso": "360.1 (L) x 279.9 (P) x 28.25 (A) mm",
    "Bateria e Alimentação": "Bateria de 4 células (Li-Íon) 90Wh",
    "Teclados e  Touchpad": "Membrana em português do Brasil padrão (ABNT 2) retroiluminado",
    "Aplicativos": "PredatorSense",
    "Conteúdo da Embalagem": "Notebook Acer Predator Helios NEO",
    "Cor": "Preto",
    "Acer P/N": "NH.QN7AL.003",
    "Garantia": "12",
    "Observações do Produto": "",
    "EAN": "000011100011",
    "Saiba Mais": "Não",
    "Código ANATEL": "ANATEL00011122",
}

PATH_OUT_DIR = f"{os.path.dirname(os.path.abspath(__file__))}\\tmp"
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_PATH, "tmp")
IMAGE_URL = "https://www.parceirosacer.com.br/Conteudos-Especiais/Notebooks/Predator-Helios-Neo/PHN16-71-76PL/PHN16-71-76PL.jpg"


class SpecificationToDocx:

    _list_specifications = list()
    _context = dict()
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
        product_image_name: str,
    ):
        template = DocxTemplate(
            os.path.join(cls._templates_dir, f"tpl_bg_{product_type.lower()}.docx")
        )
        specification_to_list = [(k, v) for k, v in specifications.items()]

        for spec in specification_to_list:
            cls._list_specifications.append({"label": spec[0], "value": spec[1]})

        image_bg = InlineImage(
            template,
            image_descriptor=os.path.join(cls._temp_dir, product_image_name),
            width=Mm(125),
            height=Mm(125),
        )

        cls._context["image_bg"] = image_bg
        cls._context["content_specs"] = cls._list_specifications
        cls._context["model"] = product_model
        cls._context["family"] = product_family

        template.render(cls._context)
        template.save(
            os.path.join(cls._temp_dir, f"bg_{product_model}_{product_id}.docx")
        )


hero_image = requests.get(IMAGE_URL)

with open(os.path.join(TEMP_DIR, "PHN16-71-76PL.jpg"), "wb") as image:
    image.write(hero_image.content)


SpecificationToDocx.execute(
    dict_spec,
    "NOTEBOOK",
    "6662662662",
    "PHN16-71-76PL",
    "Predator Helios Neo",
    "PHN16-71-76PL.jpg",
)
