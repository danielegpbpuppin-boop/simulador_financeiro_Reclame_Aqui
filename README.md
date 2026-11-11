# Simulador_Financeiro_Reclame_Aqui

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![DIO](https://img.shields.io/badge/DIO-Desafio_de_Projeto-0070f3?style=for-the-badge)

Projeto desenvolvido como parte do desafio prático da **DIO (Digital Innovation One)** sob orientação do expert Felipe Aguiar (Felipão). O objetivo principal é simular o acúmulo de patrimônio e a geração de renda passiva (dividendos mensais) através de aportes em Fundos Imobiliários.

---

## Perguntas de Negócio Respondidas

A ferramenta foi construída com foco em experiência de usuário (*planilha com cara de aplicativo*), permitindo responder a 5 perguntas centrais de planejamento financeiro:

1. **Quanto investir por mês?**  
    *Definido no painel de configurações pelo usuário (ou calculado via sugestão de 30% da renda).*
2. **Por quantos anos?**  
    *Selecionado no campo de prazo de investimento (com cenários projetados em aba/tabela dedicada).*
3. **Qual a taxa de rendimento mensal?**  
    *Informada no painel inicial conforme a expectativa média de retorno da carteira.*
4. **Quanto de patrimônio vai acumular?**  
    *Exibido no card principal de resultado (calculado dinamicamente via fórmula `VF`).*
5. **Quanto vai receber de dividendos por mês?**  
    *Exibido no card de renda passiva mensal projetada com base no patrimônio acumulado.*

---

##  Fórmulas e Lógica Utilizadas

### 1. Função `VF` (Valor Futuro)
A função `VF` é o coração financeiro do simulador. Ela calcula o valor final acumulado considerando aportes constantes e juros compostos.

$$\text{VF}(\text{taxa}, \text{nper}, \text{pgto}, [\text{vp}], [\text{pgto\_inicio}])$$

* **`taxa` (`taxa_mensal`):** Rendimento percentual esperado ao mês.
* **`nper` (`prazo_meses`):** Quantidade total de meses de investimento ($\text{anos} \times 12$).
* **`pgto` (`-aporte`):** Valor investido mensalmente (inserido com sinal negativo para ajustar a convenção do fluxo de caixa do Excel).
* **Dividendos Mensais:** Calculados multiplicando o patrimônio final obtido no `VF` pela `taxa_mensal`.

---

### 2. Função `PROCV` com Chave Composta
Para determinar a porcentagem de alocação de cada tipo de FII dependendo do perfil selecionado, foi utilizada a função `PROCV` combinada a uma **chave composta**:

```excel
=PROCV(perfil_selecionado & "_" & tipo_fundo; tabela_alocacao; 3; FALSO)
