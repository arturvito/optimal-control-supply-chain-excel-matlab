# SBPO 2026 - Optimal Control of Supply Chains

This folder contains the computational files associated with the SBPO 2026 paper:

**Optimal Control of Supply Chains: Integer-Variable Optimization and Excel Implementation**

The study proposes an integer optimal control formulation for supply chain management and bullwhip effect mitigation, implemented in MATLAB and Excel/VBA.

## Folder Structure

```text
SBPO2026/
├── main_ga_minmax_restricao.m
├── supply_chain_solver_excel_vba.xlsm
├── supply_chain_solver_macro.vba
└── README.md
```

## Files

### `main_ga_minmax_restricao.m`

MATLAB implementation of the integer optimal control problem using a genetic algorithm.

The script includes the demand profile, optimization horizon, integer decision variables, lower and upper bounds, initial population, genetic algorithm configuration, objective function, supply chain simulation routine, and result plots.

The optimization problem uses a 60-period horizon. The decision variables are the reference inventory trajectories for the wholesaler, distributor, and retailer. Therefore, the problem has 180 integer decision variables.

The MATLAB implementation exports the following figures:

```text
controles.png
estoques.png
vendas.png
```

### `supply_chain_solver_excel_vba.xlsm`

Excel spreadsheet with macro support.

This file contains the Excel implementation of the proposed model. The spreadsheet reproduces the discrete-time supply chain dynamics and allows the optimization problem to be solved using the Excel Solver with the Evolutionary method.

### `supply_chain_solver_macro.vba`

VBA source code associated with the Excel implementation.

This file contains the VBA routine used to create or support the Excel-based implementation of the supply chain optimization model.

## Problem Overview

The problem is formulated as an integer optimal control problem for a discrete-time supply chain model.

The supply chain is represented by multiple echelons. Consumer demand is treated as an external disturbance. The control strategy determines reference inventory positions for decentralized proportional controllers.

The optimization follows a shooting-based formulation. In this approach, the optimizer manipulates only the reference trajectories, while the state trajectories are obtained by direct simulation of the dynamic equations.

## Objective Function

The objective function includes terms associated with:

- demand tracking;
- inventory balancing using a MinMax criterion;
- penalty for negative inventories;
- penalty for fulfilled sales exceeding demand;
- smoothing of reference trajectories.

This structure aims to track consumer demand while avoiding excessive inventory concentration in a single supply chain echelon.

## MATLAB Requirements

To run the MATLAB implementation, the following are required:

- MATLAB;
- Global Optimization Toolbox;
- Parallel Computing Toolbox, if parallel execution is enabled.

The script uses the `ga` function with integer decision variables.

## Running the MATLAB Code

1. Open MATLAB.
2. Place `main_ga_minmax_restricao.m` in the MATLAB working directory.
3. Run:

```matlab
main_ga_minmax_restricao
```

The code executes the genetic algorithm, simulates the supply chain dynamics, and exports the result figures.

## Excel Requirements

To use the Excel implementation, the following are required:

- Microsoft Excel;
- Solver add-in enabled;
- macros enabled;
- Evolutionary Solver method.

## Using the Excel File

1. Open `supply_chain_solver_excel_vba.xlsm`.
2. Enable editing and macros, if requested.
3. Open the Excel Solver.
4. Select the objective cell corresponding to the total cost function.
5. Select the decision variable cells corresponding to the reference trajectories.
6. Add lower and upper bounds for the decision variables.
7. Add integer constraints to the decision variables.
8. Select the Evolutionary Solver method.
9. Run the optimization.

## Reproducibility Notes

The MATLAB and Excel implementations solve the same optimization problem. However, exact numerical equality between the results is not expected.

The implementations use different stochastic optimization engines and may converge to different local solutions. Therefore, the comparison should focus on feasibility, demand tracking, inventory levels, and qualitative consistency of the supply chain response.

## Citation

If you use these files, please cite the associated SBPO 2026 paper:

```bibtex
@inproceedings{author2026supplychaincontrol,
  title     = {Optimal Control of Supply Chains: Integer-Variable Optimization and Excel Implementation},
  author    = {Leonardi, L. and De Vito Jr., A. F. and Leonardi, F. and Zambuzi, N. C.},
  booktitle = {Proceedings of the Brazilian Symposium on Operations Research},
  year      = {2026}
}
```

## License

This material is intended for academic and research use. Please check the repository license for reuse conditions.

---

# SBPO 2026 - Controle Ótimo de Cadeias de Suprimentos

Esta pasta contém os arquivos computacionais associados ao artigo do SBPO 2026:

**Controle Ótimo de Cadeias de Suprimentos: Otimização com Variáveis Inteiras e Implementação em Excel**

O estudo propõe uma formulação de controle ótimo inteiro para gestão de cadeias de suprimentos e mitigação do efeito chicote, implementada em MATLAB e Excel/VBA.

## Estrutura da Pasta

```text
SBPO2026/
├── main_ga_minmax_restricao.m
├── supply_chain_solver_excel_vba.xlsm
├── supply_chain_solver_macro.vba
└── README.md
```

## Arquivos

### `main_ga_minmax_restricao.m`

Implementação em MATLAB do problema de controle ótimo inteiro usando algoritmo genético.

O script inclui o perfil de demanda, o horizonte de otimização, as variáveis inteiras de decisão, os limites inferiores e superiores, a população inicial, a configuração do algoritmo genético, a função objetivo, a rotina de simulação da cadeia de suprimentos e a geração dos gráficos de resultados.

O problema de otimização utiliza um horizonte de 60 períodos. As variáveis de decisão são as trajetórias de referência de estoque para atacadista, distribuidor e varejista. Portanto, o problema possui 180 variáveis inteiras de decisão.

A implementação em MATLAB exporta as seguintes figuras:

```text
controles.png
estoques.png
vendas.png
```

### `supply_chain_solver_excel_vba.xlsm`

Planilha em Excel com suporte a macros.

Este arquivo contém a implementação em Excel do modelo proposto. A planilha reproduz a dinâmica da cadeia de suprimentos em tempo discreto e permite resolver o problema de otimização usando o Solver do Excel com o método Evolucionário.

### `supply_chain_solver_macro.vba`

Código-fonte VBA associado à implementação em Excel.

Este arquivo contém a rotina em VBA usada para criar ou apoiar a implementação em Excel do modelo de otimização da cadeia de suprimentos.

## Visão Geral do Problema

O problema é formulado como um problema de controle ótimo inteiro para um modelo de cadeia de suprimentos em tempo discreto.

A cadeia de suprimentos é representada por múltiplos elos. A demanda do consumidor é tratada como uma perturbação externa. A estratégia de controle determina referências de posição de estoque para controladores proporcionais descentralizados.

A otimização segue uma formulação baseada em shooting. Nessa abordagem, o otimizador manipula apenas as trajetórias de referência, enquanto as trajetórias dos estados são obtidas por simulação direta das equações dinâmicas.

## Função Objetivo

A função objetivo inclui termos associados a:

- rastreamento da demanda;
- balanceamento dos estoques usando critério MinMax;
- penalização de estoques negativos;
- penalização de vendas realizadas acima da demanda;
- suavização das trajetórias de referência.

Essa estrutura busca rastrear a demanda do consumidor e evitar concentração excessiva de estoque em um único elo da cadeia de suprimentos.

## Requisitos para MATLAB

Para executar a implementação em MATLAB, são necessários:

- MATLAB;
- Global Optimization Toolbox;
- Parallel Computing Toolbox, caso a execução paralela esteja habilitada.

O script utiliza a função `ga` com variáveis inteiras de decisão.

## Execução do Código MATLAB

1. Abra o MATLAB.
2. Coloque `main_ga_minmax_restricao.m` no diretório de trabalho do MATLAB.
3. Execute:

```matlab
main_ga_minmax_restricao
```

O código executa o algoritmo genético, simula a dinâmica da cadeia de suprimentos e exporta as figuras de resultado.

## Requisitos para Excel

Para usar a implementação em Excel, são necessários:

- Microsoft Excel;
- suplemento Solver habilitado;
- macros habilitadas;
- método Evolucionário do Solver.

## Uso do Arquivo Excel

1. Abra `supply_chain_solver_excel_vba.xlsm`.
2. Habilite a edição e as macros, se solicitado.
3. Abra o Solver do Excel.
4. Selecione a célula objetivo correspondente à função custo total.
5. Selecione as células variáveis correspondentes às trajetórias de referência.
6. Adicione limites inferiores e superiores para as variáveis de decisão.
7. Adicione restrições de integralidade às variáveis de decisão.
8. Selecione o método Evolucionário.
9. Execute a otimização.

## Notas de Reprodutibilidade

As implementações em MATLAB e Excel resolvem o mesmo problema de otimização. Porém, igualdade numérica exata entre os resultados não é esperada.

As implementações usam motores de otimização estocásticos distintos e podem convergir para soluções locais diferentes. Portanto, a comparação deve considerar a viabilidade das trajetórias, o rastreamento da demanda, os níveis de estoque e a consistência qualitativa da resposta da cadeia de suprimentos.

## Citação

Se você utilizar estes arquivos, cite o artigo associado ao SBPO 2026:

```bibtex
@inproceedings{author2026supplychaincontrol,
  title     = {Optimal Control of Supply Chains: Integer-Variable Optimization and Excel Implementation},
  author    = {Leonardi, L. and De Vito Jr., A. F. and Leonardi, F. and Zambuzi, N. C.},
  booktitle = {Proceedings of the Brazilian Symposium on Operations Research},
  year      = {2026}
}
```

## Licença

Este material é disponibilizado para uso acadêmico e de pesquisa. Verifique a licença do repositório para condições de reutilização.
