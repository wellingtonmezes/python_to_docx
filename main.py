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

content_specs = []  # type: ignore
context = {}
template = DocxTemplate("template_bg_notebook.docx")

list_specs = [(k, v) for k, v in dict_spec.items()]

for data in list_specs:
    content_specs.append({"label": data[0], "value": data[1]})

context["content_specs"] = content_specs

image_bg = InlineImage(
    template, image_descriptor="PHN16-71-76PL.jpg", width=Mm(100), height=Mm(100)
)

context["image_bg"] = image_bg

template.render(context)
template.save("bg_notebook_out.docx")
