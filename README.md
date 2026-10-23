# Certificado Analítico de Água Potável (São Martinho)

Ferramenta em Streamlit que lê a planilha **FO.GSM.LMI.03.08 - Análises Microbiológicas de Água
Potável** e preenche o **FO.GSM.LMI.07.01 - Certificado Analítico - Água Potável** de cada ponto
(ex.: *Vivência 417.xlsx*), **vários pontos de uma vez**. O preenchimento é 100% determinístico:
regras fixas em `data/certificado.json`, sem IA.

## Como usar

1. Carregue o FO.GSM.LMI.03.08 da safra (barra lateral, passo 1).
2. Carregue as planilhas dos pontos — quantas quiser, de uma vez (ex.: *Vivência 417*, *631*,
   *634*, *665*), ou direto o **.zip baixado no mês anterior** — e/ou digite os números dos pontos que ainda não têm planilha (`631, 634, 665`
   ou faixas como `600-610`). Sem nada carregado, o botão **“Criar certificado para todos os
   pontos coletados no mês”** preenche a lista com todos os pontos da origem escolhida.
3. Escolha o mês (vem pré-selecionado o da coleta mais recente).
4. Confira a tabela (dá para ajustar Nº do certificado, coletor e desmarcar pontos) e clique em
   **Baixar tudo (.zip)**: uma planilha por ponto e, se houver LibreOffice, o PDF de cada certificado.

### Todo mês, em 30 segundos

O .zip que a ferramenta entrega serve de entrada no mês seguinte: atualize o FO.GSM.LMI.03.08,
jogue-o no passo 1, jogue o .zip do mês anterior no passo 2, escolha o mês e baixe o .zip novo.
Cada planilha ganha a aba do mês e as anteriores ficam como estavam. Dentro do .zip valem só as
planilhas (.xlsx/.xlsm); a pasta `PDF/`, temporários do Excel (`~$...`) e `__MACOSX` são ignorados.

## O que é preenchido automaticamente

| Campo do certificado | De onde vem |
|---|---|
| Mês / Ano | Período escolhido |
| Nº certificado | Último nº do arquivo + 1 (ou o da aba, se o mês já existir) |
| Data da amostragem | Coleta do ponto no mês (Origem + Identificação no FO.GSM.LMI.03.08) |
| Validade | Data da amostragem + 6 meses |
| Coletor | Último certificado do ponto (pontos novos: campo na barra lateral) |
| Resultados | Cor aparente, Turbidez, pH, Cloro, Coliformes totais, Termotolerantes → *E. coli* |
| Aceitável | Regras do VMP (Portaria GM/MS 888/2021): “Aceitável”, “Não aceitável” ou “Conferir” |

Regras que evitam retrabalho:

- **Recoleta**: com mais de uma coleta no mês vale a mais recente (normalmente a recoleta após um
  “Presente”); dá para escolher outra na tela.
- **Aba**: o arquivo ganha uma aba nova com o nome do mês, copiada da última. Se a aba do mesmo
  mês/ano já existe, ela é atualizada; se o mês existe mas de outro ano, a nova vira “Abril 2027”.
- **Datas digitadas sem barra** (`0510/2026`) são lidas como 05/10/2026; linhas sem data legível são
  listadas na tela.
- Pontos sem coleta no mês ficam de fora e aparecem como pendência (ou use a coleta anterior).

## Como rodar

```bash
pip install -r requirements.txt
streamlit run app.py
```

Testes: `pip install -r requirements-dev.txt && pytest`

O PDF sai do LibreOffice quando ele está instalado (no Streamlit Cloud, via `packages.txt`); sem ele
a ferramenta entrega só os Excel.

## Por que o Excel é editado no XML

O openpyxl descarta, ao salvar, as caixas de texto do cabeçalho (Unidade, Mês, Ano...) e a imagem
EMF da assinatura do responsável técnico. Por isso `autofill/services/xlsx.py` clona a aba direto
no pacote .xlsx e troca só os valores: logotipo, assinatura, bordas e abas anteriores ficam
idênticos ao original.

## Estrutura (MVC)

| Camada | Arquivo | Papel |
|---|---|---|
| Model | `data/certificado.json` | Colunas do FO.GSM.LMI.03.08, células do FO.GSM.LMI.07.01, VMP e regras |
| Model | `assets/modelo_certificado.xlsx` | Modelo do certificado para pontos que ainda não têm planilha |
| Model | `autofill/model/` | Configuração e objetos (`Amostra`, `ArquivoCertificado`, `Certificado`) |
| Serviço | `autofill/services/analises.py` | Lê o FO.GSM.LMI.03.08 (colunas achadas pelo cabeçalho) |
| Serviço | `autofill/services/certificado.py` | Lê certificados, escolhe a coleta, avalia e preenche |
| Serviço | `autofill/services/xlsx.py` | Edição do .xlsx no XML (clonar/renomear aba, gravar células) |
| Serviço | `autofill/services/exportacao.py` | Lote: .zip com planilhas e PDFs (LibreOffice) |
| View | `autofill/view/` | Tema São Martinho e componentes de tela |
| Controller | `app.py` | Liga tudo no Streamlit |

## Ajustes sem mexer no código

Edite `data/certificado.json`: validade (meses), textos “Aceitável/Não aceitável”, faixas de cada
parâmetro, nome do arquivo de pontos novos (`Vivência {ponto}.xlsx`) e, se o formulário mudar, as
células e colunas. A faixa de pH (6,0 a 9,0) só serve para marcar “Aceitável”; o VMP impresso
continua “-”, como no formulário.
