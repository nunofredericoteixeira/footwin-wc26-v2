from pathlib import Path

logos = Path("public/logos")

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
    "Ghana.png": "ghana.png",
    "Haiti.png": "haiti.png",
    "Inglaterra.png": "inglaterra.png",
    "Iraque.png": "iraque.png",
    "Irão.png": "irao.png",
    "Japão.png": "japao.png",
    "Jordan.png": "jordania.png",
    "Nova Zelândia.png": "nova-zelandia.png",
    "Noruega.png": "noruega.png",
    "Paises Baixos.png": "paises-baixos.png",
    "Panama.png": "panama.png",
    "Paraguai.png": "paraguai.png",
    "Portugal.png": "portugal.png",
    "Senegal.png": "senegal.png",
    "Suecia.png": "suecia.png",
    "Suiça.png": "suica.png",
    "Tunisia.png": "tunisia.png",
    "Turquia.png": "turquia.png",
    "Uruguai.png": "uruguai.png",
    "USA.png": "usa.png",
    "Uzbequistão.png": "uzbequistao.png",
    "costa do marfim.png": "costa-do-marfim.png",
    "marrocos.png": "marrocos.png",
    "mexico.png": "mexico.png",
}

for antigo, novo in renomes.items():
    origem = logos / antigo
    destino = logos / novo

    if origem.exists():
        if destino.exists() and origem != destino:
            print(f"Já existe destino, vou manter: {destino}")
        else:
            origem.rename(destino)
            print(f"Renomeado: {antigo} -> {novo}")
    else:
        print(f"Não encontrado: {antigo}")

print("Concluído.")
