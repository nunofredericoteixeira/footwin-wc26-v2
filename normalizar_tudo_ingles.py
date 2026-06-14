from pathlib import Path
import json
import re
import unicodedata
import shutil

BASE = Path(".")
LOGOS_DIR = BASE / "public" / "logos"
JSON_PATH = BASE / "public" / "previsoes.json"
INDEX_PATH = BASE / "index.html"

COUNTRIES = {
    "Mexico": {
        "slug": "mexico",
        "aliases": ["mexico", "méxico"]
    },
    "South Africa": {
        "slug": "south-africa",
        "aliases": ["south africa", "africa do sul", "áfrica do sul"]
    },
    "South Korea": {
        "slug": "south-korea",
        "aliases": ["south korea", "coreia", "coreia do sul", "república da coreia", "republica da coreia"]
    },
    "Czechia": {
        "slug": "czechia",
        "aliases": ["czechia", "chequia", "república checa", "republica checa"]
    },
    "Canada": {
        "slug": "canada",
        "aliases": ["canada", "canadá"]
    },
    "Bosnia and Herzegovina": {
        "slug": "bosnia-and-herzegovina",
        "aliases": ["bosnia", "bósnia", "bosnia e herzegovina", "bósnia e herzegovina", "bosnia and herzegovina"]
    },
    "Qatar": {
        "slug": "qatar",
        "aliases": ["qatar", "catar"]
    },
    "Switzerland": {
        "slug": "switzerland",
        "aliases": ["switzerland", "suica", "suíça", "suiça"]
    },
    "Brazil": {
        "slug": "brazil",
        "aliases": ["brazil", "brasil"]
    },
    "Morocco": {
        "slug": "morocco",
        "aliases": ["morocco", "marrocos"]
    },
    "Haiti": {
        "slug": "haiti",
        "aliases": ["haiti"]
    },
    "Scotland": {
        "slug": "scotland",
        "aliases": ["scotland", "escocia", "escócia"]
    },
    "United States": {
        "slug": "united-states",
        "aliases": ["united states", "estados unidos", "usa", "eua"]
    },
    "Paraguay": {
        "slug": "paraguay",
        "aliases": ["paraguay", "paraguai"]
    },
    "Australia": {
        "slug": "australia",
        "aliases": ["australia", "austrália"]
    },
    "Turkey": {
        "slug": "turkey",
        "aliases": ["turkey", "turquia"]
    },
    "Germany": {
        "slug": "germany",
        "aliases": ["germany", "alemanha"]
    },
    "Curacao": {
        "slug": "curacao",
        "aliases": ["curacao", "curaçao", "curacau"]
    },
    "Ivory Coast": {
        "slug": "ivory-coast",
        "aliases": ["ivory coast", "costa do marfim", "côte d'ivoire", "cote d'ivoire"]
    },
    "Ecuador": {
        "slug": "ecuador",
        "aliases": ["ecuador", "equador"]
    },
    "Netherlands": {
        "slug": "netherlands",
        "aliases": ["netherlands", "paises baixos", "países baixos", "holanda"]
    },
    "Japan": {
        "slug": "japan",
        "aliases": ["japan", "japao", "japão"]
    },
    "Sweden": {
        "slug": "sweden",
        "aliases": ["sweden", "suecia", "suécia"]
    },
    "Tunisia": {
        "slug": "tunisia",
        "aliases": ["tunisia", "tunísia"]
    },
    "Spain": {
        "slug": "spain",
        "aliases": ["spain", "espanha"]
    },
    "Cape Verde": {
        "slug": "cape-verde",
        "aliases": ["cape verde", "cabo verde"]
    },
    "Belgium": {
        "slug": "belgium",
        "aliases": ["belgium", "belgica", "bélgica"]
    },
    "Egypt": {
        "slug": "egypt",
        "aliases": ["egypt", "egito", "egipto"]
    },
    "Saudi Arabia": {
        "slug": "saudi-arabia",
        "aliases": ["saudi arabia", "arabia saudita", "arábia saudita"]
    },
    "Uruguay": {
        "slug": "uruguay",
        "aliases": ["uruguay", "uruguai"]
    },
    "Iran": {
        "slug": "iran",
        "aliases": ["iran", "irao", "irã", "irão"]
    },
    "New Zealand": {
        "slug": "new-zealand",
        "aliases": ["new zealand", "nova zelandia", "nova zelândia"]
    },
    "France": {
        "slug": "france",
        "aliases": ["france", "franca", "frança"]
    },
    "Senegal": {
        "slug": "senegal",
        "aliases": ["senegal"]
    },
    "Iraq": {
        "slug": "iraq",
        "aliases": ["iraq", "iraque"]
    },
    "Norway": {
        "slug": "norway",
        "aliases": ["norway", "noruega"]
    },
    "Argentina": {
        "slug": "argentina",
        "aliases": ["argentina"]
    },
    "Algeria": {
        "slug": "algeria",
        "aliases": ["algeria", "argelia", "argélia"]
    },
    "Austria": {
        "slug": "austria",
        "aliases": ["austria", "áustria"]
    },
    "Jordan": {
        "slug": "jordan",
        "aliases": ["jordan", "jordania", "jordânia"]
    },
    "Portugal": {
        "slug": "portugal",
        "aliases": ["portugal"]
    },
    "DR Congo": {
        "slug": "dr-congo",
        "aliases": ["dr congo", "congo dr", "rd congo", "republica democratica do congo", "república democrática do congo"]
    },
    "Uzbekistan": {
        "slug": "uzbekistan",
        "aliases": ["uzbekistan", "uzbequistao", "uzbequistão", "usbequistao", "usbequistão"]
    },
    "Colombia": {
        "slug": "colombia",
        "aliases": ["colombia", "colômbia"]
    },
    "England": {
        "slug": "england",
        "aliases": ["england", "inglaterra"]
    },
    "Croatia": {
        "slug": "croatia",
        "aliases": ["croatia", "croacia", "croácia"]
    },
    "Ghana": {
        "slug": "ghana",
        "aliases": ["ghana", "gana"]
    },
    "Panama": {
        "slug": "panama",
        "aliases": ["panama", "panamá"]
    },
}

def normalize_key(text):
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("ç", "c")
    text = re.sub(r"[^a-z0-9]+", "-", text)
    text = re.sub(r"^-+|-+$", "", text)
    return text

ALIAS_TO_EN = {}
ALIAS_TO_SLUG = {}

for en, info in COUNTRIES.items():
    slug = info["slug"]
    ALIAS_TO_EN[normalize_key(en)] = en
    ALIAS_TO_SLUG[normalize_key(en)] = slug
    ALIAS_TO_EN[normalize_key(slug)] = en
    ALIAS_TO_SLUG[normalize_key(slug)] = slug

    for alias in info["aliases"]:
        ALIAS_TO_EN[normalize_key(alias)] = en
        ALIAS_TO_SLUG[normalize_key(alias)] = slug

def country_to_en(value):
    key = normalize_key(value)
    return ALIAS_TO_EN.get(key, value)

def country_to_slug(value):
    key = normalize_key(value)
    return ALIAS_TO_SLUG.get(key, key)

def normalizar_logos():
    print("=== Normalizar logos para inglês ===")

    if not LOGOS_DIR.exists():
        print("Pasta de logos não encontrada.")
        return

    for file in list(LOGOS_DIR.iterdir()):
        if not file.is_file():
            continue

        if file.suffix.lower() not in [".png", ".jpg", ".jpeg", ".webp", ".svg"]:
            continue

        stem = file.stem
        slug = country_to_slug(stem)
        target = LOGOS_DIR / f"{slug}{file.suffix.lower()}"

        if file.name == target.name:
            continue

        if target.exists():
            print(f"Remover duplicado: {file.name} porque já existe {target.name}")
            file.unlink()
        else:
            print(f"{file.name} -> {target.name}")
            file.rename(target)

def traduzir_prognostico(texto):
    if not isinstance(texto, str):
        return texto

    t = texto.strip()
    key = normalize_key(t)

    if key == "empate":
        return "Draw"

    # Vitória Nome
    m = re.match(r"^Vit[oó]ria\s+(.+)$", t, flags=re.I)
    if m:
        equipa = m.group(1).strip()
        return f"{country_to_en(equipa)} win"

    return t

def normalizar_json():
    print("=== Normalizar public/previsoes.json para inglês ===")

    if not JSON_PATH.exists():
        print("previsoes.json não encontrado.")
        return

    data = json.loads(JSON_PATH.read_text(encoding="utf-8"))

    def walk(obj):
        if isinstance(obj, list):
            return [walk(x) for x in obj]

        if isinstance(obj, dict):
            novo = {}
            for k, v in obj.items():
                if k in ["casa", "home", "equipa_casa", "team_home"]:
                    novo[k] = country_to_en(v)
                    novo["casa_slug"] = country_to_slug(v)
                elif k in ["fora", "away", "equipa_fora", "team_away"]:
                    novo[k] = country_to_en(v)
                    novo["fora_slug"] = country_to_slug(v)
                elif k in ["pais", "país", "country", "seleção", "selecao"]:
                    novo[k] = country_to_en(v)
                elif k == "prognostico":
                    novo[k] = traduzir_prognostico(v)
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

    if not INDEX_PATH.exists():
        print("index.html não encontrado.")
        return

    html = INDEX_PATH.read_text(encoding="utf-8")

    logo_entries = []
    for en, info in sorted(COUNTRIES.items(), key=lambda x: x[0]):
        slug = info["slug"]
        file = f"{slug}.png"

        keys = set()
        keys.add(normalize_key(en))
        keys.add(normalize_key(slug))
        for alias in info["aliases"]:
            keys.add(normalize_key(alias))

        for key in sorted(keys):
            logo_entries.append(f'      "{key}": "{file}",')

    novo_bloco = "const LOGO_FILES = {\n" + "\n".join(logo_entries) + "\n    };"

    html2, n = re.subn(
        r"const LOGO_FILES\s*=\s*\{.*?\n\s*\};",
        novo_bloco,
        html,
        flags=re.S
    )

    if n != 1:
        raise RuntimeError(f"Não consegui substituir LOGO_FILES. Substituições: {n}")

    # Melhorar fallback para usar slug inglês se o JSON já trouxer casa_slug/fora_slug
    html2 = html2.replace(
        "const filename = LOGO_FILES[slug];",
        "const filename = LOGO_FILES[slug] || `${slug}.png`;"
    )

    INDEX_PATH.write_text(html2, encoding="utf-8")

normalizar_logos()
normalizar_json()
atualizar_index()

print("Concluído: nomes internos, JSON e logos normalizados para inglês.")
