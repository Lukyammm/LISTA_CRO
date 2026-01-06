Boa, isso aqui é importante pra **documentar direito** e evitar dor de cabeça depois. Vou escrever um texto que você pode usar tanto como **explicação interna**, **nota técnica** ou até colar numa aba “SOBRE / AJUDA” da planilha.

---

### 📋 **Finalidade da planilha – Separação de Óbitos (CRO)**

Esta planilha foi desenvolvida para **apoiar o trabalho da Comissão de Revisão de Óbitos (CRO)**, com o objetivo de **organizar, filtrar e preparar a lista de prontuários que ainda precisam ser avaliados**, facilitando o fluxo de trabalho entre a comissão e o NAC.

O foco principal é **garantir que apenas os óbitos pendentes de avaliação sejam encaminhados**, evitando retrabalho, duplicidade de análises e erros operacionais.

---

### 🧠 **Como a planilha funciona (visão geral)**

A solução se baseia em **três abas principais**, cada uma com um papel bem definido:

#### 1️⃣ **BASE_SCIH_NHE**

É a **base oficial de origem**, onde constam **todos os óbitos registrados**, com dados completos como:

* data do óbito
* prontuário
* nome do paciente
* unidade
* dados clínicos e administrativos

Essa aba **não é alterada pelo sistema**, ela apenas serve como fonte confiável de dados.

---

#### 2️⃣ **ANÁLISE_ÓBITOS**

É a aba que registra o **status da avaliação realizada pela comissão**, indicando se o óbito:

* já foi avaliado
* aguarda protocolo Londres
* já foi avaliado pelo Londres

A planilha utiliza essa aba para **identificar automaticamente quais prontuários já foram analisados**, evitando que eles apareçam novamente na lista de separação.

---

#### 3️⃣ **LISTA_SEPARAÇÃO_ÓBITOS**

Essa é a **aba final**, gerada automaticamente pelo sistema, que contém **somente os prontuários pendentes de avaliação**, no formato adequado para o NAC.

Ela é preenchida automaticamente com:

* prontuário
* nome do paciente
* unidade do óbito
* data do óbito
* status (quando existente)

E possui uma coluna de **Observações**, que é **preenchida manualmente pelo usuário**, diretamente pelo modal, antes da geração do PDF.

---

### 🖥️ **Modal CRO – por que ele existe**

O botão **“Gerar Listagem”** abre um modal moderno que centraliza todo o processo:

* seleção de **mês e ano**
* visualização clara de:

  * total de óbitos no período
  * quantos já foram avaliados
  * quantos ainda estão pendentes
* **pré-visualização da lista final**, exatamente como será enviada ao NAC
* possibilidade de **inserir ou editar observações** diretamente na lista, sem precisar mexer na planilha manualmente

Isso reduz erros, acelera o processo e mantém o padrão visual e operacional.

---

### 📄 **Geração do PDF**

Após a conferência:

* o sistema gera automaticamente um **PDF apenas com a área relevante da lista**
* ignora botões, linhas vazias e áreas fora da lista
* o arquivo é salvo diretamente na pasta oficial do Drive da CRO

O nome do arquivo segue o padrão:

```
MM_ANO_CRO.pdf
```

Garantindo rastreabilidade, padronização e histórico organizado.

---

### ✅ **Por que esse modelo foi adotado**

Esse fluxo foi pensado para:

* evitar retrabalho da comissão
* garantir que o NAC receba **somente o que precisa separar**
* manter os dados de origem íntegros
* padronizar o processo mês a mês
* reduzir erros humanos e dependência de filtros manuais
