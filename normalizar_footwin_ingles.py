from pathlib import Path
import json
import re
import unicodedata
import subprocess

LOGOS_DIR = Path("public/logos")
JSON_PATH = Path("public/previsoes.json")
INDEX_PATH = Path("index.html")

COUNTRY_MAP = {
    "mexico": ("Mexico", "mexico"),
    "africa-do-sul": ("South Africa", "south-africa"),
    "south-africa": ("South Africa", "south-africa"),

    "coreia": ("South Korea", "south-korea"),
    "coreia-do-sul": ("South Korea", "south-korea"),
    "republica-da-coreia": ("South Korea", "south-korea"),
    "south-korea": ("South Korea", "south-korea"),

    "chequia": ("Czechia", "czechia"),
    "czechia": ("Czechia", "czechia"),

    "canada": ("Canada", "canada"),

    "bosnia": ("Bosnia and Herzegovina", "bosnia-and-herzegovina"),
    "bosnia-e-herzegovina": ("Bosnia and Herzegovina", "bosnia-and-herzegovina"),
    "bosnia-and-herzegovina": ("Bosnia and Herzegovina", "bosnia-and-herzegovina"),

    "catar": ("Qatar", "qatar"),
    "qatar": ("Qatar", "qatar"),

    "suica": ("Switzerland", "switzerland"),
    "switzerland": ("Switzerland", "switzerland"),

    "brasil": ("Brazil", "brazil"),
    "brazil": ("Brazil", "brazil"),

    "marrocos": ("Morocco", "morocco"),
    "morocco": ("Morocco", "morocco"),

    "haiti": ("Haiti", "haiti"),

    "escocia": ("Scotland", "scotland"),
    "scotland": ("Scotland", "scotland"),

    "estados-unidos": ("United States", "united-states"),
    "usa": ("United States", "united-states"),
    "united-states": ("United States", "united-states"),

    "paraguai": ("Paraguay", "paraguay"),
    "paraguay": ("Paraguay", "paraguay"),

    "australia": ("Australia", "australia"),

    "turquia": ("Turkey", "turkey"),
    "turkey": ("Turkey", "turkey"),

    "alemanha": ("Germany", "germany"),
    "germany": ("Germany", "germany"),

    "curacao": ("Curacao", "curacao"),
    "curacau": ("Curacao", "curacao"),

    "costa-do-marfim": ("Ivory Coast", "ivory-coast"),
    "ivory-coast": ("Ivory Coast", "ivory-coast"),

    "equador": ("Ecuador", "ecuador"),
    "ecuador": ("Ecuador", "ecuador"),

    "paises-baixos": ("Netherlands", "netherlands"),
    "netherlands": ("Netherlands", "netherlands"),
    "holanda": ("Netherlands", "netherlands"),

    "japao": ("Japan", "japan"),
    "japan": ("Japan", "japan"),

    "suecia": ("Sweden", "sweden"),
    "sweden": ("Sweden", "sweden"),

    "tunisia": ("Tunisia", "tunisia"),

    "espanha": ("Spain", "spain"),
    "spain": ("Spain", "spain"),

    "cabo-verde": ("Cape Verde", "cape-verde"),
    "cape-verde": ("Cape Verde", "cape-verde"),

    "belgica": ("Belgium", "belgium"),
    "belgium": ("Belgium", "belgium"),

    "egipto": ("Egypt", "egypt"),
    "egito": ("Egypt", "egypt"),
    "egypt": ("Egypt", "egypt"),

    "arabia-saudita": ("Saudi Arabia", "saudi-arabia"),
    "saudi-arabia": ("Saudi Arabia", "saudi-arabia"),

    "uruguai": ("Uruguay", "uruguay"),
    "uruguay": ("Uruguay", "uruguay"),

    "irao": ("Iran", "iran"),
    "iran": ("Iran", "iran"),

    "nova-zelandia": ("New Zealand", "new-zealand"),
    "new-zealand": ("New Zealand", "new-zealand"),

    "franca": ("France", "france"),
    "france": ("France", "france"),

    "senegal": ("Senegal", "senegal"),

    "iraque": ("Iraq", "iraq"),
    "iraq": ("Iraq", "iraq"),

    "noruega": ("Norway", "norway"),
    "norway": ("Norway", "norway"),

    "argentina": ("Argentina", "argentina"),

    "argelia": ("Algeria", "algeria"),
    "algeria": ("Algeria", "algeria"),

    "austria": ("Austria", "austria"),

    "jordania": ("Jordan", "jordan"),
    "jordan": ("Jordan", "jordan"),

    "portugal": ("Portugal", "portugal"),

    "congo-dr": ("DR Congo", "dr-congo"),
    "rd-congo": ("DR Congo", "dr-congo"),
    "dr-congo": ("DR Congo", "dr-congo"),

    "uzbequistao": ("Uzbekistan", "uzbekistan"),
    "usbequistao": ("Uzbekistan", "uzbekistan"),
    "uzbekistan": ("Uzbekistan", "uzbekistan"),

    "colombia": ("Colombia", "colombia"),

    "inglaterra": ("England", "england"),
    "england": ("England", "england"),

    "croacia": ("Croatia", "croatia"),
    "croatia": ("Croatia", "croatia"),

    "ghana": ("Ghana", "ghana"),
    "gana": ("Ghana", "ghana"),

    "panama": ("Panama", "panama"),
}

def run(cmd):
    subprocess.run(cmd, check=True)

def norm(text):
    text = str(text or "").strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("ç", "c")
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")

def country_en(value):
    key = norm(value)
    return COUNTRY_MAP.get(key, (value, key))[0]

def country_slug(value):
    key = norm(value)
    return COUNTRY_MAP.get(key, (value, key))[1]

def normalizar_logos():
    print("=== Normalizar logos ===")

    for file in list(LOGOS_DIR.iterdir()):
        if not file.is_file():
            continue

        if file.suffix.lower() not in [".png", ".jpg", ".jpeg", ".webp", ".svg"]:
            continue

        slug = country_slug(file.stem)
        target = LOGOS_DIR / f"{slug}{file.suffix.lower()}"

        if file.name == target.name:
            continue

        if target.exists():
            print(f"Remover duplicado: {file.name}")
            file.unlink()
        else:
            print(f"{file.name} -> {target.name}")
            file.rename(target)

def normalizar_json():
    print("=== Normalizar JSON ===")

    data = json.loads(JSON_PATH.read_text(encoding="utf-8"))

    def walk(obj):
        if isinstance(obj, list):
            return [walk(x) for x in obj]

        if isinstance(obj, dict):
            novo = {}
            for k, v in obj.items():
                if k == "casa":
                    novo[k] = country_en(v)
                    novo["casa_slug"] = country_slug(v)
                elif k == "fora":
                    novo[k] = country_en(v)
                    novo["fora_slug"] = country_slug(v)
                elif k == "prognostico":
                    if isinstance(v, str):
                        if norm(v) == "empate":
                            novo[k] = "Draw"
                        else:
                            m = re.match(r"^Vit[oó]ria\s+(.+)$", v, flags=re.I)
                            if m:
                                novo[k] = f"{country_en(m.group(1))} win"
                            else:
                                novo[k] = v
                    else:
                        novo[k] = v
                else:
                    novo[k] = walk(v)
            return novo

        return obj

    data = walk(data)

    JSON_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

def atualizar_index():
    print("=== Atualizar index.html ===")

    html = INDEX_PATH.read_text(encoding="utf-8")

    entries = []
    done = set()

    for key, (_, slug) in sorted(COUNTRY_MAP.items()):
        if key not in done:
            entries.append(f'      "{key}": "{slug}.png",')
            done.add(key)

    bloco = "const LOGO_FILES = {\n" + "\n".join(entries) + "\n    };"

    html, n = re.subn(
        r"const LOGO_FILES\s*=\s*\{.*?\n\s*\};",
        bloco,
        html,
        flags=re.S
    )

    if n != 1:
        raise RuntimeError(f"Não consegui substituir LOGO_FILES. Substituições: {n}")

    html = html.replace(
        "const filename = LOGO_FILES[slug];",
        "const filename = LOGO_FILES[slug] || `${slug}.png`;"
    )

    INDEX_PATH.write_text(html, encoding="utf-8")

normalizar_logos()
normalizar_json()
atualizar_index()

print("Concluído.")
