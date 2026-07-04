# Painel de Indicadores de Saúde

Site estático para compartilhar os dados da planilha `Demanda_Comparação Municipios`.

Quando publicado no GitHub Pages, o painel fica em:

https://maykesubhue.github.io/_dados_municipais/

## Rodar localmente

```powershell
python -m http.server 4173 -d painel-saude
```

Abra `http://127.0.0.1:4173`.

## Atualizar dados

1. Baixe/exporte a planilha atual para `tmp/demanda_comparacao_municipios.xlsx`.
2. Rode:

```powershell
python painel-saude/scripts/build_data.py
```

O script recria `painel-saude/data.json`.

## Publicar

### GitHub Pages

Este repositório inclui o workflow `.github/workflows/pages.yml`, que publica automaticamente a raiz do projeto no GitHub Pages a cada push na branch `main`.

```powershell
git push -u origin main
```

### Vercel

Use `painel-saude` como diretório raiz do projeto no Vercel. O site não precisa de build.

Com a Vercel CLI autenticada:

```powershell
cd painel-saude
vercel --prod --yes
```

Se publicar pelo painel web da Vercel, selecione esta pasta como o diretório raiz. Isso mantém o painel separado do app financeiro da pasta principal.
