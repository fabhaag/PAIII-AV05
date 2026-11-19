# Automação de Certificados de Calibração CMM — Schulz S.A.

Aplicação web transacional (MVP) desenvolvida em **Python** e **Streamlit** para automatizar a ingestão de planilhas de medição provenientes de Máquinas de Medição por Coordenadas (**MMC / CMM**), executar os cálculos metrológicos oficiais (média, desvio padrão amostral, erro sistemático, validação de tolerâncias, fator de abrangência $k$ e incerteza expandida $U$) e emitir certificados de calibração padronizados em **PDF**.

Projeto desenvolvido no âmbito da unidade curricular **Projeto Aplicado III (AV 05)** do **Centro Universitário SENAI/SC — Campus Florianópolis**, em parceria com o laboratório de metrologia da **Schulz S.A. (Joinville/SC)**.

---

## Funcionalidades Principais

* **Ingestão e Reparo Automático de Planilhas (`.xlsx` / `.xls`):** corrige incompatibilidades de *namespace* XML em memória (`repair_xlsx_stream`), extrai metadados do cabeçalho e higieniza automaticamente colunas residuais e valores nulos.
* **Motor de Cálculos Metrológicos:**
  * Agrupamento dinâmico das leituras (`Medida 1` e `Medida 2`) para número variável de características.
  * Cálculo automático da **média aritmética**, **desvio padrão amostral** (`ddof=1`) e **erro sistemático** ($\vert{}\bar{x} - \text{Nominal}\vert{}$).
  * Validação ponto a ponto contra o **Limite de Erro Admissível (LEA)** com destaque visual em tela (**DENTRO DA TOLERÂNCIA** em verde / **FORA DA TOLERÂNCIA** em vermelho) e parecer global (**APROVADO** / **REPROVADO**).
  * **Cálculo dinâmico da Incerteza Expandida ($U$) e Fator de Abrangência ($k$):** implementa integralmente o balanço de 8 componentes de incerteza da Schulz S.A., graus de liberdade efetivos (Welch-Satterthwaite) e tabela *t-Student* (95,45%).
* **Emissão Automatizada de Certificados em PDF:** geração em memória via **ReportLab** no layout oficial da Schulz S.A., com nomenclatura padronizada para download imediato.

---

## Estrutura do Repositório

```text
PAIII-AV05/
├── app.py                         # Aplicação principal (Streamlit + Motor Metrológico + ReportLab)
├── iniciar_aplicacao_schulz.bat   # Script executável de instalação e inicialização automática (Windows)
├── requirements.txt               # Lista de dependências Python
├── Imagens/
│   └── BarraSuperior.jpg          # Cabeçalho institucional utilizado na geração do PDF
└── README.md                      # Documentação do repositório