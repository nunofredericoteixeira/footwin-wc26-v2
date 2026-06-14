import json
from pathlib import Path

ficheiro = Path("public/previsoes.json")

dados = json.loads(ficheiro.read_text(encoding="utf-8"))

for jogo in dados.get("jogos", []):
    casa = jogo.get("casa", "")
    fora = jogo.get("fora", "")

    if casa == "Australia" and fora == "Turkey":
        jogo["data"] = "2026-06-14"
        jogo["hora"] = "05:00"

    if casa == "Ivory Coast" and fora == "Ecuador":
        jogo["data"] = "2026-06-15"
        jogo["hora"] = "00:00"

    # garantir estado coerente
    rc = str(jogo.get("resultado_real_casa", "")).strip()
    rf = str(jogo.get("resultado_real_fora", "")).strip()

    if rc != "" and rf != "":
        jogo["estado"] = "final"
    elif str(jogo.get("previsao_casa", "")).strip() != "" and str(jogo.get("previsao_fora", "")).strip() != "":
        jogo["estado"] = "previsto"

ficheiro.write_text(
    json.dumps(dados, ensure_ascii=False, indent=2),
    encoding="utf-8"
)

print("Horas corrigidas em public/previsoes.json")
