# Automação de Certificados de Calibração CMM — Schulz S.A.

Aplicação web transacional (MVP) desenvolvida em **Python** e **Streamlit** para automatizar a ingestão de planilhas de medição provenientes de Máquinas de Medição por Coordenadas (**MMC / CMM**), executar os cálculos metrológicos oficiais (média, desvio padrão amostral, erro sistemático, validação de tolerâncias, fator de abrangência $k$ e incerteza expandida $U$), persistir os registros com integridade em banco local **SQLite** e emitir certificados de calibração padronizados em **PDF** com rastreabilidade.

Projeto desenvolvido no âmbito da unidade curricular **Projeto Aplicado III (AV 05)** do **Centro Universitário SENAI/SC — Campus Florianópolis**, em parceria com o laboratório de metrologia da **Schulz S.A. (Joinville/SC)**.

---

## Funcionalidades Principais

* **Ingestão e Reparo Automático de Planilhas (`.xlsx` / `.xls`):** corrige incompatibilidades de *namespace* XML em memória (`repair_xlsx_stream`), extrai metadados do cabeçalho e higieniza automaticamente colunas residuais e valores nulos.
* **Motor de Cálculos Metrológicos:**
  * Agrupamento dinâmico das leituras (`Medida 1` e `Medida 2`) para número variável de características.
  * Cálculo automático da **média aritmética**, **desvio padrão amostral** (`ddof=1`) e **erro sistemático** ($|\bar{x} - \text{Nominal}|$).
  * Validação ponto a ponto contra o **Limite de Erro Admissível (LEA)** com destaque visual em tela (**DENTRO DA TOLERÂNCIA** em verde / **FORA DA TOLERÂNCIA** em vermelho) e parecer global (**APROVADO** / **REPROVADO**).
  * **Cálculo dinâmico da Incerteza Expandida ($U$) e Fator de Abrangência ($k$):** implementa integralmente o balanço de 8 componentes de incerteza da Schulz S.A., graus de liberdade efetivos (Welch-Satterthwaite) e tabela *t-Student* (95,45%).
* **Emissão Automatizada de Certificados em PDF com QR Code:** geração em memória via **ReportLab** no layout oficial da Schulz S.A., incluindo **QR Code** no rodapé para validação mobile, nomenclatura padronizada e salvamento automático em disco.
* **Versionamento Automático de PDFs (Auditoria):** ao reemitir um certificado já existente (por exemplo, após correção de metadados), o sistema move automaticamente a versão anterior para a subpasta `certificados/historico/` acrescentando data e hora (*timestamp*) ao nome do arquivo, preservando a rastreabilidade para auditorias **ISO 9001 / IATF 16949**.
* **Persistência em SQLite Anual + Exportação para Excel:**
  * Grava cada calibração em transação única (**ACID**) no banco nativo `calibracoes_{ano}.db` (ex.: `calibracoes_2026.db`, criando automaticamente um novo arquivo a cada virada de ano), espelhando o modelo relacional do PostgreSQL (`metadados`, `certificado` e `parametros`).
  * Disponibiliza na barra lateral a exportação sob demanda para **Excel** (`Calibrações-{ano}.xlsx`) com as 3 abas padronizadas e escrita atômica (evitando corrupção de arquivo).
* **Diretórios Configuráveis (`config.txt`):** permite personalizar facilmente os locais de gravação dos certificados e dos dados, inclusive para pastas de rede corporativas.

---

## Estrutura do Repositório

```text
PAIII-AV05/
├── app.py                         # Aplicação principal (Streamlit + Motor Metrológico + SQLite + ReportLab)
├── config.txt                     # Configuração de diretórios (gerado automaticamente na 1ª execução)
├── iniciar_aplicacao_schulz.bat   # Script executável de instalação e inicialização automática (Windows)
├── requirements.txt               # Lista de dependências Python
├── Imagens/
│   └── BarraSuperior.jpg          # Cabeçalho institucional utilizado na geração do PDF
├── certificados/                  # Certificados PDF emitidos (criada automaticamente)
│   └── historico/                 # Versões anteriores de certificados reemitidos (com timestamp)
├── dados/                         # Bancos SQLite anuais e planilhas Excel exportadas (criada automaticamente)
│   ├── calibracoes_2026.db        # Banco de dados SQLite principal do ano
│   └── Calibrações-2026.xlsx      # Planilha consolidada (abas: metadados, certificado, parametros)
└── README.md                      # Documentação do repositório
```

---

## Configuração de Diretórios e Segurança (`config.txt`)

Na primeira execução, o arquivo `config.txt` é criado automaticamente na raiz do projeto com os caminhos padrão relativos à pasta da aplicação:

```ini
# Configuração de diretórios do Gerador de Certificados CMM - Schulz
# Use caminhos relativos (ex: certificados) ou absolutos (ex: C:\Schulz\Certificados)
PASTA_CERTIFICADOS=certificados
PASTA_DADOS=dados
```

### Recomendação de Segurança: Uso de Pasta de Rede

Para ambientes corporativos sem banco de dados dedicado, **recomenda-se alterar os caminhos padrão no `config.txt` para uma pasta de rede compartilhada** (caminho UNC ou unidade mapeada). Isso incrementa significativamente a segurança e a governança dos dados ao:

1. **Proteger contra perdas locais:** evita que falhas de hardware na estação de trabalho comprometam o histórico de calibrações e certificados.
2. **Aproveitar o backup corporativo:** integra automaticamente os arquivos `.db`, `.xlsx` e `.pdf` às rotinas de backup e políticas de retenção da TI da empresa.
3. **Centralizar o repositório:** facilita o acesso dos arquivos por auditores, engenheiros da qualidade e outras estações autorizadas.

**Exemplo de configuração apontando para um servidor de rede:**

```ini
PASTA_CERTIFICADOS=\\servidor-arquivos\metrologia\certificados
PASTA_DADOS=\\servidor-arquivos\metrologia\dados
```

*(Também é possível utilizar unidades mapeadas do Windows, como `Z:\Metrologia\Certificados` e `Z:\Metrologia\Dados`.)*

---

## Instalação e Execução (Windows)

### Pré-requisitos
* **Windows 10** ou superior.
* **Python 3.10+** instalado e adicionado ao `PATH` do sistema *(o módulo `sqlite3` já é nativo do Python)*.
* Acesso à internet apenas na primeira execução (para instalação das dependências).

### Inicialização Automatizada
1. Clone ou baixe este repositório e extraia os arquivos em uma pasta local.
2. Dê um duplo clique no arquivo **`iniciar_aplicacao_schulz.bat`**.
3. O script executará automaticamente 5 etapas:
   * `[1/5]` Verificação do Python 3.10+ no `PATH`.
   * `[2/5]` Criação e ativação do ambiente virtual isolado (`venv`).
   * `[3/5]` Instalação/atualização das bibliotecas listadas em `requirements.txt`.
   * `[4/5]` Verificação/criação do `config.txt` e das pastas `certificados`, `certificados\historico` e `dados`.
   * `[5/5]` Inicialização do servidor local do Streamlit em `http://localhost:8501`.

---

## Como Utilizar

1. **Upload da Planilha:** arraste ou selecione o arquivo `.xlsx` / `.xls` exportado pela CMM na área de upload.
2. **Conferência dos Metadados:** revise os campos preenchidos automaticamente e ajuste, se necessário, informações complementares (como *Teste*, *Temperatura* ou *Cliente*).
3. **Análise Metrológica:** confira a tabela consolidada (verde = dentro da tolerância; vermelho = fora da tolerância), o Fator de abrangência ($k$), a Incerteza ($U$) e o parecer final (**APROVADO** ou **REPROVADO**).
4. **Emissão e Persistência:** clique em **"Salvar e Baixar Certificado em PDF"**. O sistema fará o download no navegador, gravará o PDF na pasta de certificados (versionando a cópia anterior em `historico/`, se houver) e registrará os dados no banco SQLite do ano correspondente.
5. **Exportação para Excel:** sempre que desejar gerar ou atualizar a planilha `Calibrações-{ano}.xlsx`, selecione o ano desejado na **barra lateral esquerda** e clique em **"Exportar Relatório Excel (.xlsx)"**.