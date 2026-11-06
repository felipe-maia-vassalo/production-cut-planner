# Fictional sample Excel files

These files contain invented product names, identifiers and dimensions. They contain no company spreadsheet data. Each workbook includes an English and Portuguese instructions sheet.

| File | Use |
| --- | --- |
| `sample_structures.xlsx` | Bill of materials for product lookup and cut planning. Keep sheet `RELAÇÃO DE ESTRUTURAS` and its Portuguese headers. |
| `sample_product_mapping.xlsx` | Product name mapping / DE-PARA for product lookup. The mapping is the first sheet. |
| `sample_cut_plan.xlsx` | Ready-made cut plan for machine scheduling. Choose **Upload Excel cut plan** and upload its `Cuts` sheet. |

For product lookup, paste one name per line:

```text
DEMO CHAIR
DEMO TABLE
DEMO FRAME
```

For cut planning, paste codes `900001`, `900002`, `900003` on separate lines and quantities `100`, `50`, `25` on matching separate lines. Expected result: four cut rows totaling **800 pieces** (400, 200, 100, 100).

Additional lookup cases: `DEMO OLD CHAIR` refers to a discontinued product, whose name is red in the structures file. `DEMO MISSING` has no matching product.

All lengths are in millimetres. These invented dimensions demonstrate the app; they are not a manufacturing specification.

## Português

As três planilhas usam produtos, códigos e medidas fictícios. Nenhuma planilha da empresa foi usada. A aba `Instructions` traz o passo a passo em inglês e português. Na busca, cole os três nomes acima, um por linha. No planejamento, informe os três códigos e as quantidades na mesma ordem. O resultado esperado é **800 peças** em quatro linhas de corte. A planilha de cortes pronta pode ser enviada diretamente na programação das máquinas.
