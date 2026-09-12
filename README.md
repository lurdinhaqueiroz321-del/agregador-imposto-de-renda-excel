# Agregador de Dados para Imposto de Renda no Excel

Projeto desenvolvido como parte do desafio da DIO para aplicar recursos do Excel na criação de uma ferramenta prática, organizada e de fácil utilização.

## Sobre o projeto

O agregador reúne, em um único arquivo, as principais informações utilizadas na preparação da declaração de Imposto de Renda. A planilha foi criada sem dados pessoais reais e pode ser preenchida em uma cópia privada pelo usuário.

O arquivo oferece:

- painel inicial com resumo do preenchimento;
- menu de navegação entre as abas;
- campos destacados para digitação;
- listas suspensas e validações de dados;
- cálculos automáticos de valores líquidos, diferenças e totais;
- identificação automática de registros que precisam de revisão;
- checklist para conferência dos documentos;
- links rápidos para páginas oficiais e comprovantes digitais.

> **Atenção:** não publique no GitHub uma versão preenchida. CPF, rendimentos, patrimônio e despesas são informações pessoais e financeiras sensíveis.

## Demonstração

![Painel inicial](images/inicio.png)

![Checklist de conferência](images/checklist.png)

## Estrutura da planilha

| Aba | Finalidade |
| --- | --- |
| Início | Apresenta indicadores, pendências, progresso e menu de navegação. |
| Titular | Reúne dados de identificação e contato do contribuinte. |
| Dependentes | Registra dependentes e despesas relacionadas. |
| Rendimentos | Organiza rendimentos, fontes pagadoras e imposto retido. |
| Bens e Direitos | Controla bens, saldos declarados e alterações patrimoniais. |
| Pagamentos | Registra despesas, reembolsos e valores líquidos. |
| Dívidas | Controla credores, saldos e pagamentos realizados. |
| Checklist | Permite conferir documentos e etapas antes da declaração. |
| Apoio | Mantém as listas utilizadas nas validações de dados. |

## Recursos do Excel aplicados

- tabelas formatadas com filtros;
- validação de dados por listas;
- formatação condicional para pendências;
- referências absolutas entre abas;
- funções `SE`, `OU`, `CONT.SE`, `CONT.VALORES`, `SOMA`, `MÁXIMO` e `HIPERLINK`;
- preenchimento automático de fórmulas;
- formatos de moeda, percentual e data;
- congelamento de painéis nas bases maiores.

### Exemplos de fórmulas

Status de conferência:

```excel
=SE(B7="";"";SE(OU(A7="";C7="";E7="");"Revisar";"OK"))
```

Valor líquido após reembolso:

```excel
=SE(E7="";"";MÁXIMO(E7-SE(F7="";0;F7);0))
```

Progresso do checklist:

```excel
=CONT.SE(Checklist!D7:D24;"Concluído")/CONT.VALORES(Checklist!A7:A24)
```

## Como utilizar

1. Baixe o arquivo `agregador_imposto_de_renda.xlsx`.
2. Abra-o no Microsoft Excel.
3. Salve uma cópia privada antes de inserir dados pessoais.
4. Preencha as células amarelas de cada aba.
5. Corrija os registros identificados como **Revisar**.
6. Anexe ou informe links para os comprovantes quando desejar.
7. Finalize os itens da aba **Checklist**.
8. Confira o painel da aba **Início**.

## Organização dos arquivos

```text
.
├── agregador_imposto_de_renda.xlsx
├── images/
│   ├── inicio.png
│   └── checklist.png
├── .gitignore
├── LICENSE
└── README.md
```

## Aprendizados

O desenvolvimento permitiu praticar a organização de informações em diferentes abas, a criação de fórmulas reutilizáveis, o uso de validações para reduzir erros de digitação e a construção de uma interface simples para orientar o preenchimento. A documentação no GitHub também tornou o projeto mais fácil de apresentar, baixar e reproduzir.

## Observação

Esta planilha é uma ferramenta de organização. Ela não substitui o programa oficial da Receita Federal nem a orientação de um profissional habilitado. Regras tributárias podem mudar; confira sempre as orientações válidas para o exercício da declaração.

## Licença

Distribuído sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE).
