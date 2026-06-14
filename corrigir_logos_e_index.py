from pathlib import Path
import re

logos = Path("public/logos")
index_path = Path("index.html")

# 1) Renomear ficheiros para nomes simples, sem acentos/espaços/ç
renomes = {
    "Africa do Sul.png": "africa-do-sul.png",
    "Alemanha.png": "alemanha.png",
    "Algeria.png": "argelia.png",
    "Arabia Saudita.png": "arabia-saudita.png",
    "Argentina.png": "argentina.png",
    "Australia.png": "australia.png",
    "Austria.png": "austria.png",
    "Belgica.png": "belgica.png",
    "Bosnia.png": "bosnia.png",
    "Brasil.png": "brasil.png",
    "Cabo Verde.png": "cabo-verde.png",
    "Canada.png": "canada.png",
    "Catar.png": "catar.png",
    "Chequia.png": "chequia.png",
    "Colombia.png": "colombia.png",
    "Congo DR.png": "congo-dr.png",
    "Coreia.png": "coreia.png",
    "Croacia.png": "croacia.png",
    "Curacao.png": "curacao.png",
    "Egipto.png": "egipto.png",
    "Equador.png": "equador.png",
    "Escocia.png": "escocia.png",
    "Espanha.png": "espanha.png",
    "França.png": "franca.png",
    "França.png": "franca.png",
    "Ghana.png": "ghana.png",
    "Haiti.png": "haiti.png",
    "Inglaterra.png": "inglaterra.png",
    "Iraque.png": "iraque.png",
    "Irão.png": "irao.png",
    "Irão.png": "irao.png",
    "Japão.png": "japao.png",
    "Japão.png": "japao.png",
    "Jordan.png": "jordania.png",
    "Nova Zelândia.png": "nova-zelandia.png",
    "Nova Zelândia.png": "nova-zelandia.png",
    "Noruega.png": "noruega.png",
    "Paises Baixos.png": "paises-baixos.png",
    "Panama.png": "panama.png",
    "Paraguai.png": "paraguai.png",
    "Portugal.png": "portugal.png",
    "Senegal.png": "senegal.png",
    "Suecia.png": "suecia.png",
    "Suiça.png": "suica.png",
    "Suíça.png": "suica.png",
    "Tunisia.png": "tunisia.png",
    "Turquia.png": "turquia.png",
    "Uruguai.png": "uruguai.png",
    "USA.png": "usa.png",
    "Uzbequistão.png": "uzbequistao.png",
    "Uzbequistão.png": "uzbequistao.png",
    "costa do marfim.png": "costa-do-marfim.png",
    "marrocos.png": "marrocos.png",
    "mexico.png": "mexico.png",
}

print("=== Renomear logos ===")
for antigo, novo in renomes.items():
    origem = logos / antigo
    destino = logos / novo

    if origem.exists():
        if destino.exists() and origem.resolve() != destino.resolve():
            print(f"Destino já existe, removo origem duplicada: {origem}")
            origem.unlink()
        else:
            origem.rename(destino)
            print(f"{antigo} -> {novo}")

# 2) Substituir o bloco LOGO_FILES no index.html
novo_bloco = '''const LOGO_FILES = {
      "africa-do-sul": "africa-do-sul.png",
      "south-africa": "africa-do-sul.png",

      "alemanha": "alemanha.png",
      "germany": "alemanha.png",

      "argelia": "argelia.png",
      "algeria": "argelia.png",

      "arabia-saudita": "arabia-saudita.png",
      "saudi-arabia": "arabia-saudita.png",

      "argentina": "argentina.png",
      "australia": "australia.png",
      "austria": "austria.png",

      "belgica": "belgica.png",
      "belgium": "belgica.png",

      "bosnia": "bosnia.png",
      "bosnia-e-herzegovina": "bosnia.png",
      "bosnia-and-herzegovina": "bosnia.png",

      "brasil": "brasil.png",
      "brazil": "brasil.png",

      "cabo-verde": "cabo-verde.png",
      "cape-verde": "cabo-verde.png",

      "canada": "canada.png",

      "catar": "catar.png",
      "qatar": "catar.png",

      "chequia": "chequia.png",
      "czechia": "chequia.png",

      "colombia": "colombia.png",

      "congo-dr": "congo-dr.png",
      "rd-congo": "congo-dr.png",
      "dr-congo": "congo-dr.png",
      "republica-democratica-do-congo": "congo-dr.png",

      "coreia": "coreia.png",
      "coreia-do-sul": "coreia.png",
      "south-korea": "coreia.png",
      "republica-da-coreia": "coreia.png",

      "croacia": "croacia.png",
      "croatia": "croacia.png",

      "curacau": "curacao.png",
      "curacao": "curacao.png",
      "curaçao": "curacao.png",

      "egito": "egipto.png",
      "egipto": "egipto.png",
      "egypt": "egipto.png",

      "equador": "equador.png",
      "ecuador": "equador.png",

      "escocia": "escocia.png",
      "scotland": "escocia.png",

      "espanha": "espanha.png",
      "spain": "espanha.png",

      "estados-unidos": "usa.png",
      "united-states": "usa.png",
      "usa": "usa.png",

      "franca": "franca.png",
      "france": "franca.png",

      "gana": "ghana.png",
      "ghana": "ghana.png",

      "haiti": "haiti.png",

      "inglaterra": "inglaterra.png",
      "england": "inglaterra.png",

      "irao": "irao.png",
      "iran": "irao.png",

      "iraque": "iraque.png",
      "iraq": "iraque.png",

      "japao": "japao.png",
      "japan": "japao.png",

      "jordania": "jordania.png",
      "jordan": "jordania.png",

      "marrocos": "marrocos.png",
      "morocco": "marrocos.png",

      "mexico": "mexico.png",

      "noruega": "noruega.png",
      "norway": "noruega.png",

      "nova-zelandia": "nova-zelandia.png",
      "new-zealand": "nova-zelandia.png",

      "paises-baixos": "paises-baixos.png",
      "netherlands": "paises-baixos.png",
      "holanda": "paises-baixos.png",

      "panama": "panama.png",

      "paraguai": "paraguai.png",
      "paraguay": "paraguai.png",

      "portugal": "portugal.png",

      "senegal": "senegal.png",

      "suecia": "suecia.png",
      "sweden": "suecia.png",

      "suica": "suica.png",
      "switzerland": "suica.png",

      "tunisia": "tunisia.png",

      "turquia": "turquia.png",
      "turkey": "turquia.png",

      "uruguai": "uruguai.png",
      "uruguay": "uruguai.png",

      "usbequistao": "uzbequistao.png",
      "uzbequistao": "uzbequistao.png",
      "uzbekistan": "uzbequistao.png",

      "costa-do-marfim": "costa-do-marfim.png",
      "ivory-coast": "costa-do-marfim.png",
      "cote-divoire": "costa-do-marfim.png",
      "cote-d-ivoire": "costa-do-marfim.png"
    };'''

html = index_path.read_text(encoding="utf-8")

html_novo, n = re.subn(
    r"const LOGO_FILES\s*=\s*\{.*?\n\s*\};",
    novo_bloco,
    html,
    flags=re.S
)

if n != 1:
    raise RuntimeError(f"Não consegui substituir o bloco LOGO_FILES. Substituições feitas: {n}")

index_path.write_text(html_novo, encoding="utf-8")
print("=== index.html atualizado ===")
print("Bloco LOGO_FILES substituído com sucesso.")
