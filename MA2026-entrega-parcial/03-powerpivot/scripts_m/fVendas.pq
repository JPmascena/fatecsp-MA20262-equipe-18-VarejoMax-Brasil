// fVendas — tabela fato (cópia fiel da fVendas da Etapa 2, com tipos explícitos)
let
    Fonte = cad_Etapa2{[Item = "fVendas", Kind = "Table"]}[Data],
    Tipos = Table.TransformColumnTypes(Fonte, {
        {"ID_Pedido", type text}, {"Item", Int64.Type}, {"Data_Pedido", type date},
        {"Data_Envio", type nullable date}, {"Data_Entrega_Prevista", type nullable date},
        {"Data_Entrega_Real", type nullable date},
        {"Prazo_Previsto_Dias", Int64.Type}, {"Prazo_Real_Dias", Int64.Type}, {"Atraso_Dias", Int64.Type},
        {"Status_Atraso", type text}, {"Status_Pedido", type text},
        {"Canal", type text}, {"Loja", type text}, {"Chave_Canal_Loja", type text},
        {"ID_Cliente", type text}, {"Cidade", type text}, {"CEP", type text},
        {"ID_Produto", type text}, {"Quantidade", Int64.Type},
        {"Preco_Unitario", Currency.Type}, {"Custo_Unitario", Currency.Type},
        {"Desconto_Pct", type number}, {"Valor_Bruto", Currency.Type}, {"Valor_Liquido", Currency.Type},
        {"Valor_Item_Original", Currency.Type}, {"Frete_Valor", Currency.Type},
        {"Forma_Pagamento", type text}, {"Parcelas", Int64.Type},
        {"Transportadora", type text}, {"Avaliacao_Cliente", Int64.Type},
        {"Arquivo_Origem", type text}, {"Linha_Origem", Int64.Type}})
in
    Tipos
