// dProduto — dimensão de produtos (com Custo_Ref para [Custo Total])
let
    Fonte = cad_Etapa2{[Item = "dProduto", Kind = "Table"]}[Data],
    Tipos = Table.TransformColumnTypes(Fonte, {
        {"ID_Produto", type text}, {"Produto", type text}, {"Categoria", type text},
        {"Marca", type text}, {"Preco_Ref_2025", Currency.Type},
        {"Preco_Ref_2026", Currency.Type}, {"Custo_Ref", Currency.Type}})
in
    Tipos
