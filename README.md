# 🛍️ Sistema de Venda e Consignação em VBA

Sistema feito no **Excel com VBA** para uma loja de roupas: registra **vendas** e **consignações** por um formulário, **baixa o estoque** automaticamente, controla o **prazo de devolução** e recebe as **devoluções** de peças consignadas.

| Arquivo | O que é |
|---|---|
| 📗 [`Sistema_Venda_Consignacao.xlsm`](Sistema_Venda_Consignacao.xlsm) | Planilha com os formulários e as macros já embutidas |

---

## 🗂️ Abas da planilha

| Aba | Conteúdo |
|---|---|
| **Produtos** | 100 peças de vestuário: código, descrição, preço de compra, preço de venda e **quantidade em estoque** (com os botões **Inserir** e **Devolver**) |
| **Venda-Consignação** | Cada movimento: código, tipo (Venda/Consignação), pessoa, produto, quantidade, valor, data e **data máxima de devolução** |

---

## ⚙️ Como funciona

### ➕ Botão **Inserir** → formulário `Movimento`

```
┌──────────────── Movimento ────────────────┐
│ Código:        [ 7        ]  (automático) │
│ Pessoa:        [ Maria              ]     │
│ Produto:       [ Vestido Floral...  ▼ ]   │
│ Quantidade:    [ 2  ]                     │
│ Valor unit.:   [ 143,00 ]                 │
│ Valor total:   [ 286,00 ]                 │
│ ( ) Venda   (•) Consignação               │
│      [ Inserir ]        [ Finalizar ]     │
└───────────────────────────────────────────┘
```

- O **código** do movimento é gerado sozinho (maior código da aba + 1).
- Ao escolher o produto, o **preço de venda** é buscado na aba Produtos e o **total** é calculado.
- Se a quantidade for **maior que o estoque**, o sistema avisa e corrige.
- Ao inserir, o item é gravado na aba **Venda-Consignação** e o **estoque é baixado**.
- Na **consignação**, a **data máxima de devolução** é a data de hoje **+ 3 dias**.
- Ao abrir o formulário, aparece um **aviso das consignações vencidas**.

### ↩️ Botão **Devolver** → formulário `Devo`

1. Lista só os movimentos do tipo **Consignação**.
2. Ao escolher o movimento, carrega os **produtos** daquele movimento.
3. Informa a **quantidade devolvida**, que não pode passar da quantidade levada.
4. A quantidade **volta para o estoque**, e a linha da consignação é **atualizada** ou **apagada** se tudo foi devolvido.

---

## 🧩 Estrutura do código

| Módulo | Rotinas |
|---|---|
| `Módulo1` | `AbrirFormulario`, `AbrirFormularioDEVO`, `VerificarConsignacoes` |
| `Movimento` (UserForm) | `UserForm_Initialize`, `CbProd_Change`, `EdtQuant_Change`, `CalcularTotal`, `BtnInserir_Click`, `VerificarConsignacoesVencidas`, `ValidarCampos`, `LimparCampos` |
| `Devo` (UserForm) | `CarregarMovimentos`, `CbMovimento_Change`, `BtnDevolver_Click` |

```vba
' Consignação: grava a data máxima de devolução (3 dias)
If OptionButton2.Value = True Then
    wsVC.Cells(novaLinha, "H").Value = Date + 3
End If
```

---

## 🚀 Como usar

1. Baixe **`Sistema_Venda_Consignacao.xlsm`** e abra no Excel.
2. Clique em **"Habilitar conteúdo"** para liberar as macros.
3. Na aba **Produtos**, use os botões **Inserir** e **Devolver**.

---

## 🎓 Informações acadêmicas

| | |
|---|---|
| **Instituição** | Fatec Taquaritinga |
| **Curso** | Análise e Desenvolvimento de Sistemas — 1º semestre |
| **Disciplina** | Microinformática |
| **Ano** | 2025 |

## 👤 Autor

**Juan Gabriel Pereira Augusto**
GitHub: [@JuanGabrielPereiraAugusto](https://github.com/JuanGabrielPereiraAugusto)
